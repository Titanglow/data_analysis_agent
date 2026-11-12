import os
import re
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from concurrent.futures import ThreadPoolExecutor, as_completed
import streamlit as st
import pandas as pd
import plotly.express as px
from data_agent.graph import build_agent_graph
from data_agent.excel_exporter import (
    export_analysis_to_excel,
    export_combined_analysis_to_excel,
    build_korean_stat_tables,
    analyze_color_cell_contexts,
    cluster_text_series_by_topic,
    shorten_long_label,
)
from data_agent.llm import get_installed_ollama_models
from data_agent.tools.data_loader import (
    load_and_sync_to_db,
    filter_by_keywords,
    run_sql_query,
    export_db_to_sql_script,
    DB_PATH,
    TABLE_NAME,
)

CONFIG_DIR = Path(".streamlit")
CONFIG_FILE = CONFIG_DIR / "config.toml"
if not CONFIG_FILE.exists():
    CONFIG_DIR.mkdir(exist_ok=True)
    CONFIG_FILE.write_text('[theme]\nprimaryColor = "#059669"\n', encoding="utf-8")

st.set_page_config(page_title="사내 데이터 분석 Agent", layout="wide")

st.markdown(
    """
    <style>
    :root { --primary-color: #059669 !important; }
    span[data-baseweb="tag"],
    div[data-baseweb="select"] span[data-baseweb="tag"],
    [data-testid="stMultiSelect"] span[data-baseweb="tag"] {
        background-color: #059669 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    span[data-baseweb="tag"] span { color: #ffffff !important; }
    div[role="radiogroup"] label > div:first-child { border-color: #059669 !important; }
    div[role="radiogroup"] label[data-baseweb="radio"] > div:first-child {
        background-color: #059669 !important;
        border-color: #059669 !important;
    }
    button[kind="primary"], [data-testid="stBaseButton-primary"] {
        background-color: #059669 !important;
        border-color: #059669 !important;
        color: white !important;
    }
    button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover {
        background-color: #047857 !important;
        border-color: #047857 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

DISTINCT_PALETTE = ["#2563EB", "#DC2626", "#059669", "#D97706", "#7C3AED", "#0891B2", "#DB2777"]

st.title("📊 자연어 기반 사내 데이터 분석 Agent")

MODEL_LABELS = {
    "cloud_groq": "⚡ Cloud: Groq (초고속 무료 API)",
    "cloud_gemini": "☁️ Cloud: Gemini (Google 무료 API)",
    "local_ollama": "🔒 Local: Ollama (사내보안 정밀 모드)",
    "mock": "🧪 테스트 모드 (Mock·0.2초 검증)",
}

if "saved_results" not in st.session_state:
    st.session_state["saved_results"] = {}
if "last_analyzed_df" not in st.session_state:
    st.session_state["last_analyzed_df"] = None


def style_dataframe_with_colors(df: pd.DataFrame):
    def _highlight_cell(val):
        s = str(val)
        if "빨강" in s or "FAIL" in s.upper():
            return "background-color: #FEE2E2; color: #991B1B; font-weight: bold;"
        if "초록" in s:
            return "background-color: #DCFCE7; color: #166534; font-weight: bold;"
        if "노랑" in s or "주황" in s:
            return "background-color: #FEF9C3; color: #854D0E; font-weight: bold;"
        if "파랑" in s or "보라" in s:
            return "background-color: #DBEAFE; color: #1E40AF; font-weight: bold;"
        return ""

    if len(df) <= 300:
        return df.style.map(_highlight_cell)
    return df


# =====================================================================
# 🔍 고도화된 한국어 자연어 -> SQL 변환 엔진
# ('Title칸에 문서 들어간거', '배달만 조회해줘', '쿠폰이나 내역 조회' 완벽 지원!)
# =====================================================================
COLUMN_ALIASES = {
    "title": ["title", "타이틀", "제목", "이름", "명칭", "시나리오명"],
    "id": ["id", "tc-id", "tc_id", "tc id", "아이디", "번호", "코드"],
    "precondition": ["precondition", "사전조건", "전제조건", "준비상태", "초기상태"],
    "procedure": ["procedure", "test step", "steps", "절차", "수행절차", "테스트절차", "동작"],
    "expected": ["expected result", "expected", "pass/fail criteria", "기대결과", "예상결과", "판정기준"],
    "actual": ["actual result", "actual", "execution status", "status", "실제결과", "수행결과", "판정", "결과"],
    "category": ["category", "domain", "app", "카테고리", "분류", "도메인", "앱", "구분"],
    "line": ["라인", "생산라인", "공정", "설비"],
}

NL_FILLER_WORDS = {
    "조회해줘", "조회해달라고", "조회해", "조회", "보여줘", "보여달라고", "찾아줘", "찾아",
    "검색해줘", "검색", "추출해줘", "추출", "출력해줘", "알려줘", "분석해줘", "데이터",
    "들어간거", "들어간", "포함된거", "포함된", "포함", "관련된", "관련", "있는거", "있는",
    "칸에", "칸에서", "칸", "열에", "열에서", "열", "컬럼에", "컬럼에서", "컬럼", "필드에",
    "항목에", "항목", "행만", "행", "부분", "내용", "중에서", "중에", "에서", "만", "좀",
    "그거", "이거", "저거", "전부", "모두", "다", "또는", "이나", "나", "이랑", "랑", "하고",
}


def build_smart_sql(nl_query: str, df: pd.DataFrame, table_name: str) -> str:
    """
    1) 사용자가 'Title칸에 문서 들어간거'처럼 특정 컬럼을 지정하면 해당 컬럼만 정확히 LIKE 검색합니다.
    2) '배달만 조회해줘', '쿠폰이나 내역 조회'처럼 부분 키워드만 입력해도 조사/어미를 떼어내고
       실제 데이터가 들어있는 컬럼(또는 주요 텍스트 컬럼)에 대해 정확한 WHERE ... LIKE '%키워드%' SQL을 생성합니다.
    """
    if not nl_query or not nl_query.strip():
        return f"SELECT * FROM {table_name} LIMIT 100;"

    q_orig = nl_query.strip()
    if q_orig in ("전체", "전체 조회", "모두 조회"):
        return f"SELECT * FROM {table_name} LIMIT 100;"

    all_cols = list(df.columns)
    where_clauses = []

    # 1단계: 색상 키워드('빨강', '초록', '노랑', '배경', '강조') 우선 감지
    if any(k in q_orig for k in ["빨강", "빨간", "초록", "노랑", "배경", "색상", "색깔", "강조"]):
        color_kw = "빨강" if ("빨강" in q_orig or "빨간" in q_orig) else ("초록" if "초록" in q_orig else ("노랑" if "노랑" in q_orig else "강조"))
        or_likes = [f'"{col}" LIKE \'%{color_kw}%\'' for col in all_cols]
        return f"SELECT * FROM {table_name} WHERE (" + " OR ".join(or_likes) + ") LIMIT 100;"

    # 2단계: 질문에서 특정 컬럼 지정 여부 감지 (예: 'Title칸에', '제목에', 'Precondition에서')
    target_cols = []
    matched_col_tokens = set()
    q_lower = q_orig.lower()

    for col in all_cols:
        c_low = col.lower()
        if c_low in q_lower:
            target_cols.append(col)
            matched_col_tokens.add(c_low)
        else:
            for _, aliases in COLUMN_ALIASES.items():
                if any(a == c_low or a in c_low for a in aliases):
                    for alias in aliases:
                        if re.search(rf"(?:^|\s){re.escape(alias)}(?:칸|열|컬럼|에서|에|\s|$)", q_lower):
                            target_cols.append(col)
                            matched_col_tokens.add(alias)

    target_cols = list(dict.fromkeys(target_cols))

    # 3단계: 수치형 조건 감지 (예: '불량수 20 넘는 행')
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    nums = re.findall(r"\d+(?:\.\d+)?", q_orig)
    used_num_tokens = set()
    if nums and num_cols:
        target_num = nums[0]
        num_col = next((col for col in num_cols if col.lower() in q_lower), None)
        if num_col is None:
            for col in num_cols:
                if ("불량" in q_orig and "불량" in col) or ("생산" in q_orig and "생산" in col) or ("가동" in q_orig and "가동" in col):
                    num_col = col
                    break
        if num_col:
            used_num_tokens.add(target_num)
            if any(k in q_orig for k in ["넘", "초과", "보다 큰", "많은", "높은", ">"]):
                op = ">=" if "이상" in q_orig else ">"
            elif any(k in q_orig for k in ["이상", "부터", ">="]):
                op = ">="
            elif any(k in q_orig for k in ["미만", "보다 작은", "적은", "낮은", "<"]):
                op = "<"
            elif any(k in q_orig for k in ["이하", "이내", "까지", "<="]):
                op = "<="
            else:
                op = "="
            where_clauses.append(f'"{num_col}" {op} {target_num}')

    # 4단계: 핵심 검색 키워드 추출 (따옴표 안의 단어 우선 + 한국어 조사 제거)
    quoted_words = re.findall(r"['\"“”‘’]([^'\"“”‘’]+)['\"“”‘’]", q_orig)
    keywords = [w.strip() for w in quoted_words if w.strip()]

    # 따옴표가 없거나 추가 키워드가 있는 경우 토큰 분리
    q_cleaned = re.sub(r"['\"“”‘’][^'\"“”‘’]+['\"“”‘’]", " ", q_orig)
    raw_tokens = re.findall(r"[가-힣A-Za-z0-9_-]+", q_cleaned)

    for tok in raw_tokens:
        tok_low = tok.lower()
        if tok in used_num_tokens or tok_low in matched_col_tokens or tok_low in NL_FILLER_WORDS:
            continue
        # 컬럼명+조사 형태('Title칸에', '제목칸에', '라인에서') 제거
        is_col_word = False
        for c_token in matched_col_tokens:
            if tok_low.startswith(c_token):
                is_col_word = True
                break
        if is_col_word:
            continue

        # 단어 끝에 붙은 한국어 조사/어미('배달만' -> '배달', '쿠폰이나' -> '쿠폰', '내역만' -> '내역', 'A라인랑' -> 'A라인') 제거
        stem = re.sub(
            r"(만|이나|나|이랑|랑|하고|과|와|을|를|이|가|은|는|에서|에게|으로|로|에|도|씩|들)$",
            "",
            tok,
        ).strip()
        if not stem or stem.lower() in NL_FILLER_WORDS or stem.lower() in matched_col_tokens:
            continue
        if len(stem) >= 1 and stem not in keywords:
            keywords.append(stem)

    # 5단계: 추출된 키워드로 SQL WHERE 조건 조립
    str_cols = df.select_dtypes(include=["object", "category"]).columns.tolist() or all_cols
    search_scope_cols = target_cols if target_cols else str_cols

    kw_sql_parts = []
    for kw in keywords:
        kw_safe = kw.replace("'", "''")
        # 현재 데이터프레임의 검색 대상 컬럼 중 해당 키워드를 부분 포함(contains)하는 컬럼 탐색
        matching_cols = []
        for col in search_scope_cols:
            if df[col].astype(str).str.contains(kw, case=False, na=False, regex=False).any():
                matching_cols.append(col)

        if matching_cols:
            if len(matching_cols) == 1:
                kw_sql_parts.append(f'"{matching_cols[0]}" LIKE \'%{kw_safe}%\'')
            else:
                or_cols = " OR ".join(f'"{c}" LIKE \'%{kw_safe}%\'' for c in matching_cols[:4])
                kw_sql_parts.append(f"({or_cols})")
        else:
            # 데이터에 해당 키워드가 없더라도 절대 SELECT * 로 무시하지 않고, 지정 컬럼(또는 대표 텍스트 컬럼)에 LIKE 쿼리 생성!
            fallback_cols = target_cols if target_cols else search_scope_cols[:3]
            if len(fallback_cols) == 1:
                kw_sql_parts.append(f'"{fallback_cols[0]}" LIKE \'%{kw_safe}%\'')
            else:
                or_cols = " OR ".join(f'"{c}" LIKE \'%{kw_safe}%\'' for c in fallback_cols)
                kw_sql_parts.append(f"({or_cols})")

    if kw_sql_parts:
        # 여러 키워드('A라인랑 B라인', '쿠폰이나 내역')는 OR로 묶어 둘 다 조회되게 처리
        if len(kw_sql_parts) == 1:
            where_clauses.append(kw_sql_parts[0])
        else:
            where_clauses.append("(" + " OR ".join(kw_sql_parts) + ")")

    sql = f"SELECT * FROM {table_name}"
    if where_clauses:
        sql += " WHERE " + " AND ".join(where_clauses)
    sql += " LIMIT 100;"
    return sql


def generate_smart_assistant_question(
    target_df: pd.DataFrame,
    aux_sheets: dict[str, pd.DataFrame] | None,
    user_filter_text: str,
) -> str:
    if len(target_df) == 0:
        return "현재 데이터의 핵심 현황, 주요 분류별 특징, 원인 가설 및 실무 조치 제안을 정리해 줘."

    all_cols = target_df.columns.tolist()
    scope_parts = []
    if user_filter_text and user_filter_text.strip():
        scope_parts.append(f"검색조건('{user_filter_text.strip()}')")

    if "시트명" in all_cols:
        main_s = target_df["시트명"].unique().tolist()
        scope_parts.append(f"대상 시트[{', '.join(map(str, main_s))}]({len(target_df)}건)")
    else:
        scope_parts.append(f"조회 데이터({len(target_df)}건)")

    if aux_sheets:
        scope_parts.append(f"보조 시트[{', '.join(aux_sheets.keys())}] 통합")

    stat_snippets = []
    if "생산량" in all_cols and "불량수" in all_cols:
        p_sum = int(target_df["생산량"].sum())
        d_sum = int(target_df["불량수"].sum())
        rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
        stat_snippets.append(f"총 생산량 {p_sum:,}개, 총 불량수 {d_sum:,}개(평균 불량률 {rate:.2f}%)")
    else:
        title_col = next((c for c in all_cols if any(k in c.lower() for k in ["title", "제목", "타이틀", "시나리오"])), None)
        if title_col:
            topic_df = cluster_text_series_by_topic(target_df[title_col], max_groups=4)
            if len(topic_df) > 0:
                top_topics = ", ".join(
                    f"{r['주제/분류 그룹']}({r['건수']}건)" for _, r in topic_df.head(3).iterrows()
                )
                stat_snippets.append(f"주요 주제분류[{top_topics}]")

        color_df = analyze_color_cell_contexts(target_df)
        if len(color_df) > 0:
            stat_snippets.append(f"🎨색상강조 {len(color_df)}건 감지")

    scope_title = " + ".join(scope_parts)
    stat_str = " / ".join(stat_snippets) if stat_snippets else f"총 {len(target_df)}건"

    return (
        f"[AI 비서 자동 맞춤 질문] {scope_title} ({stat_str}) 데이터를 바탕으로 "
        f"① 주제/분류별 핵심 현황 및 비중(%), ② 중점 관리 시나리오(또는 이상·강조 항목) 상세 분석, "
        f"③ 기술적·업무적 원인 가설 및 취약점, ④ 담당 부서별 즉시 조치(24h)·근본 대책(1개월)을 정리해 줘."
    )


@st.cache_data(ttl=600, show_spinner=False)
def fetch_related_news_links(query_keyword: str, max_items: int = 4) -> list[dict[str, str]]:
    encoded = urllib.parse.quote(query_keyword)
    rss_url = f"https://news.google.com/rss/search?q={encoded}&hl=ko&gl=KR&ceid=KR:ko"
    articles = []
    try:
        req = urllib.request.Request(rss_url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req, timeout=2.5) as resp:
            xml_bytes = resp.read()
        root = ET.fromstring(xml_bytes)
        for item in root.findall(".//item")[:max_items]:
            title = item.findtext("title", default="관련 뉴스 기사")
            link = item.findtext("link", default="#")
            pub_date = item.findtext("pubDate", default="")[:16]
            articles.append({"title": title, "link": link, "date": pub_date})
    except Exception:
        pass
    return articles


def render_related_links_box(target_df: pd.DataFrame):
    all_cols = target_df.columns.tolist()
    keywords = []
    if "불량유형" in all_cols:
        keywords.extend(target_df["불량유형"].dropna().astype(str).unique()[:2].tolist())
    else:
        title_col = next((c for c in all_cols if any(k in c.lower() for k in ["title", "제목", "app", "domain"])), None)
        if title_col:
            t_df = cluster_text_series_by_topic(target_df[title_col], max_groups=3)
            for _, r in t_df.head(2).iterrows():
                kw_clean = re.sub(r"[📁📌📄\[\]관련항목등\s]+", " ", str(r["주제/분류 그룹"])).strip()
                if kw_clean:
                    keywords.append(kw_clean)
        keywords.append("SW QA 품질 테스트 자동화")

    search_query = " ".join(keywords[:3]) if keywords else "사내 데이터 품질 관리 자동화"

    with st.expander(f"🌐 분석 데이터 연관 뉴스 & 외부 기술 자료 링크 자동 추천 (키워드: `{search_query}`)", expanded=False):
        custom_kw = st.text_input("🔎 연관 뉴스/기술자료 검색 키워드 (직접 수정 가능):", value=search_query)
        news_items = fetch_related_news_links(custom_kw, max_items=4)

        c_news, c_ref = st.columns([2, 1])
        with c_news:
            st.markdown("**📰 실시간 연관 뉴스 헤드라인**")
            if news_items:
                for item in news_items:
                    st.markdown(f"- 🔗 [{item['title']}]({item['link']}) `{item['date']}`")
            else:
                st.caption("우측 포털 원클릭 검색 링크를 이용해 주세요.")

        with c_ref:
            q_enc = urllib.parse.quote(custom_kw)
            st.markdown("**📚 전문 기술 자료 & 포털 원클릭 검색**")
            st.markdown(
                f"- 🟢 [네이버 뉴스: '{custom_kw}' 검색](https://search.naver.com/search.naver?where=news&query={q_enc})\n"
                f"- 🔵 [구글 뉴스: '{custom_kw}' 검색](https://news.google.com/search?q={q_enc}&hl=ko&gl=KR)\n"
                f"- 📑 [Google 학술/기술논문(Scholar) 검색](https://scholar.google.co.kr/scholar?q={q_enc})"
            )


def _classify_verification_goal(text: str) -> str:
    """Expected Result처럼 문장이 모두 다른 열을 의미 있는 '검증 목적 유형'으로 자동 분류합니다."""
    s = str(text)
    if any(k in s for k in ["오류", "에러", "않는다", "없음", "실패", "차단", "불가"]):
        return "🚫 오류·예외 방지 및 미표시 검증"
    if any(k in s for k in ["이동", "전환", "열리", "열린", "돌아", "접속"]):
        return "🔗 화면/문서 이동 및 링크 전환 검증"
    if any(k in s for k in ["숨겨", "펼쳐", "접기", "펼치", "유지", "상태", "토글"]):
        return "⚙️ UI 상태 변화(접기/펼치기/유지) 검증"
    if any(k in s for k in ["표시된", "나타난", "보인다", "출력", "목록"]):
        return "🖥️ 화면/데이터 정상 표시 검증"
    return "📌 기타 기능 동작 검증"


def _count_procedure_steps(text: str) -> str:
    """Procedure/Test Step의 단계 수(복잡도)를 계산해 분류합니다."""
    s = str(text).strip()
    if not s or s == "-":
        return "미기재"
    # 줄바꿈, 화살표(->, →), 번호(1., 2., PR-01) 등으로 단계 수 추정
    steps = max(
        len([l for l in s.splitlines() if l.strip()]),
        len(re.findall(r"(?:->|→|PR-\d+|\d+\.)", s)),
        1,
    )
    if steps <= 1:
        return "1단계 (단순 동작)"
    if steps == 2:
        return "2단계 (기본 절차)"
    if steps == 3:
        return "3단계 (중간 절차)"
    return f"4단계 이상 (복합 절차)"


# =====================================================================
# 📊 스마트 다차원 시각화: 1건씩 쪼개지는 현상 방지 + 긴 문장 가로 막대 & 폰트 확대!
# =====================================================================
def render_smart_charts(target_df: pd.DataFrame, key_prefix: str = "main"):
    if len(target_df) == 0:
        st.info("조회된 데이터가 없습니다.")
        return

    num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
    all_cols = target_df.columns.tolist()

    # [모드 A] 제조·수치형 데이터
    if num_cols and ("생산량" in num_cols or "불량수" in num_cols):
        cat_cols = target_df.select_dtypes(include=["object", "category"]).columns.tolist()
        line_col = "라인" if "라인" in all_cols else ("생산라인" if "생산라인" in all_cols else (cat_cols[0] if cat_cols else all_cols[0]))
        defect_type_col = next((c for c in ["불량유형", "불량원인", "유형", "분류", "카테고리"] if c in all_cols), None)
        time_col = next((c for c in ["일자", "날짜", "주차", "월"] if c in all_cols), None)
        shift_col = next((c for c in ["시간대", "교대", "근무조", "시간"] if c in all_cols), None)
        val_col = "불량수" if "불량수" in num_cols else num_cols[0]

        row1_c1, row1_c2 = st.columns(2)
        with row1_c1:
            if "생산량" in num_cols and "불량수" in num_cols and line_col in all_cols:
                agg_line = target_df.groupby(line_col, as_index=False)[["생산량", "불량수"]].sum()
                agg_line["불량률(%)"] = (agg_line["불량수"] / agg_line["생산량"] * 100).round(2)
                agg_line["표시라벨"] = agg_line.apply(lambda r: f"{int(r['불량수'])}개 ({r['불량률(%)']}%)", axis=1)
                fig1 = px.bar(
                    agg_line, x=line_col, y="불량수", color=line_col, text="표시라벨",
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"① {line_col}별 총 불량수 및 평균 불량률(%) 비교",
                )
                fig1.update_traces(textposition="outside", textfont_size=13)
                fig1.update_layout(height=420, font=dict(size=13))
                st.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_num1")

        with row1_c2:
            if defect_type_col:
                agg_defect = target_df.groupby([defect_type_col, line_col], as_index=False)[val_col].sum()
                fig2 = px.bar(
                    agg_defect, x=defect_type_col, y=val_col, color=line_col, barmode="group", text_auto=True,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"② {defect_type_col}별 · {line_col}별 {val_col} 분류",
                )
                fig2.update_traces(textposition="outside", textfont_size=13)
                fig2.update_layout(height=420, font=dict(size=13))
                st.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_num2")

        row2_c1, row2_c2 = st.columns(2)
        with row2_c1:
            if time_col and "생산량" in target_df.columns and "불량수" in target_df.columns:
                agg_time = target_df.groupby([time_col, line_col], as_index=False)[["생산량", "불량수"]].sum()
                agg_time["불량률(%)"] = (agg_time["불량수"] / agg_time["생산량"] * 100).round(2)
                fig3 = px.line(
                    agg_time, x=time_col, y="불량률(%)", color=line_col, markers=True, text="불량률(%)",
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"③ {time_col}별 불량률(%) 추이 (사내 기준 1.5% 선 표시)",
                )
                fig3.add_hline(y=1.5, line_dash="dash", line_color="#DC2626", annotation_text="사내 기준 (1.5%)")
                fig3.update_layout(height=420, font=dict(size=13))
                st.plotly_chart(fig3, use_container_width=True, key=f"{key_prefix}_num3")
        with row2_c2:
            if shift_col:
                agg_shift = target_df.groupby([shift_col, defect_type_col or line_col], as_index=False)[val_col].sum()
                fig4 = px.bar(
                    agg_shift, x=shift_col, y=val_col, color=defect_type_col or line_col, barmode="group", text_auto=True,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"④ {shift_col}(주/야간)별 비교",
                )
                fig4.update_layout(height=420, font=dict(size=13))
                st.plotly_chart(fig4, use_container_width=True, key=f"{key_prefix}_num4")
        return

    # =========================================================================
    # [모드 B] TC / 검증 시트 및 범주형 문서 전용 시각화
    # (Title 1건씩 쪼개짐 방지 -> 유사 키워드 그룹핑! / 긴 문장 -> 가로형 막대 & 요약 라벨!)
    # =========================================================================
    clean_df = target_df.copy()
    for c in clean_df.columns:
        clean_df[c] = clean_df[c].astype(str).apply(lambda x: re.sub(r"\s*\[[^\]]+\]", "", x).strip())

    title_col = next((c for c in all_cols if any(k in c.lower() for k in ["title", "제목", "타이틀", "시나리오", "기능"])), None)
    cat_col = next((c for c in ["Category", "Test Technique", "분류", "Domain", "App"] if c in all_cols), None)
    precond_col = next((c for c in all_cols if any(k in c.lower() for k in ["precondition", "사전조건", "전제조건", "환경"])), None)
    proc_col = next((c for c in all_cols if any(k in c.lower() for k in ["procedure", "step", "절차", "동작"])), None)
    exp_col = next((c for c in all_cols if any(k in c.lower() for k in ["expected", "criteria", "기대결과", "예상결과"])), None)
    real_status_col = next(
        (c for c in all_cols if any(k in c.lower() for k in ["execution status", "actual result", "status", "판정", "수행결과"])),
        None,
    )

    r1_c1, r1_c2 = st.columns(2)

    # [차트 1] Title(또는 Category) 유사 주제 키워드 자동 그룹핑 차트 (1건씩 나뉘지 않고 '문서', '검색', '목차' 등으로 합산!)
    with r1_c1:
        target_c1 = cat_col if (cat_col and clean_df[cat_col].nunique() <= 12) else (title_col or all_cols[min(1, len(all_cols) - 1)])
        grouped_c1 = cluster_text_series_by_topic(clean_df[target_c1], max_groups=8)
        if len(grouped_c1) > 0:
            fig_c1 = px.bar(
                grouped_c1,
                x="주제/분류 그룹",
                y="건수",
                color="주제/분류 그룹",
                text="표시라벨",
                hover_data=["포함된 상세 항목"],
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"① [{target_c1}] 유사 주제(키워드)별 묶음 건수 및 비중(%)",
            )
            fig_c1.update_traces(textposition="outside", textfont_size=13)
            fig_c1.update_layout(
                height=440,
                font=dict(size=13),
                xaxis_title="자동 분류된 핵심 주제 그룹 (마우스 오버 시 상세 목록 표시)",
                yaxis_title="묶음 건수 (건)",
                showlegend=False,
                margin=dict(t=55, b=40),
            )
            st.plotly_chart(fig_c1, use_container_width=True, key=f"{key_prefix}_cat1")

    # [차트 2] 실제 판정 상태(Pass/Fail/Not Run)가 있으면 표시하고, 없으면 Expected Result를 '검증 목적 유형'으로 묶어 표시!
    with r1_c2:
        if real_status_col and clean_df[real_status_col].nunique() <= 10:
            vc_st = clean_df[real_status_col].value_counts().reset_index()
            vc_st.columns = [real_status_col, "건수"]
            fig_c2 = px.pie(
                vc_st, names=real_status_col, values="건수", hole=0.42,
                color_discrete_sequence=["#059669", "#2563EB", "#D97706", "#DC2626", "#7C3AED"],
                title=f"② [{real_status_col}] 실행·판정 상태 분포",
            )
            fig_c2.update_traces(textinfo="label+percent+value", textfont_size=13)
            fig_c2.update_layout(height=440, font=dict(size=13))
            st.plotly_chart(fig_c2, use_container_width=True, key=f"{key_prefix}_cat2")
        elif exp_col:
            # Expected Result의 긴 문장들을 1건씩 쪼개지 않고 '검증 목적 유형'으로 스마트 분류!
            goal_series = clean_df[exp_col].apply(_classify_verification_goal)
            vc_goal = goal_series.value_counts().reset_index()
            vc_goal.columns = ["검증 목적 유형", "건수"]
            fig_c2 = px.pie(
                vc_goal, names="검증 목적 유형", values="건수", hole=0.42,
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"② [{exp_col}] 기대결과 검증 목적 유형별 비중 (유사 문장 자동 분류)",
            )
            fig_c2.update_traces(textinfo="label+percent+value", textfont_size=13)
            fig_c2.update_layout(height=440, font=dict(size=13))
            st.plotly_chart(fig_c2, use_container_width=True, key=f"{key_prefix}_cat2")

    r2_c1, r2_c2 = st.columns(2)

    # [차트 3] Precondition(사전조건) 등 긴 문장 컬럼 -> 깔끔한 축약 라벨 + 가로형 막대그래프(Horizontal Bar)로 가독성 해결!
    with r2_c1:
        color_ctx_df = analyze_color_cell_contexts(target_df)
        if len(color_ctx_df) > 0:
            ctx_agg = color_ctx_df.groupby(["색상 표시된 컬럼(열)", "셀 실제 내용"], as_index=False).size()
            ctx_agg.columns = ["색상 표시된 컬럼(열)", "셀 실제 내용", "건수"]
            fig_c3 = px.bar(
                ctx_agg, x="색상 표시된 컬럼(열)", y="건수", color="셀 실제 내용", barmode="group", text_auto=True,
                color_discrete_sequence=["#DC2626", "#D97706", "#2563EB", "#059669", "#7C3AED"],
                title="③ 색상 강조 셀(🎨)의 컬럼별·실제 값별 분포",
            )
            fig_c3.update_traces(textposition="outside", textfont_size=13)
            fig_c3.update_layout(height=440, font=dict(size=13))
            st.plotly_chart(fig_c3, use_container_width=True, key=f"{key_prefix}_cat3")
        elif precond_col:
            vc_pre = clean_df[precond_col].value_counts().head(8).reset_index()
            vc_pre.columns = ["원문_사전조건", "건수"]
            vc_pre["비중(%)"] = (vc_pre["건수"] / len(clean_df) * 100).round(1)
            vc_pre["요약_라벨"] = vc_pre["원문_사전조건"].apply(lambda x: shorten_long_label(x, max_len=28))
            vc_pre["표시라벨"] = vc_pre.apply(lambda r: f"{r['건수']}건 ({r['비중(%)']}%)", axis=1)
            # 긴 문장이 아래로 깔려 그래프가 찌그러지지 않도록 가로형(orientation='h') 막대로 표시!
            fig_c3 = px.bar(
                vc_pre.iloc[::-1],
                x="건수",
                y="요약_라벨",
                orientation="h",
                color="요약_라벨",
                text="표시라벨",
                hover_data=["원문_사전조건"],
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"③ [{precond_col}] 사전조건 유형별 건수 (긴 문장 가로형 요약 차트)",
            )
            fig_c3.update_traces(textposition="outside", textfont_size=13)
            fig_c3.update_layout(
                height=440,
                font=dict(size=13),
                xaxis_title="건수 (건)",
                yaxis_title="사전조건 요약 (마우스 오버 시 전체 문장 표시)",
                showlegend=False,
                margin=dict(l=20, r=60, t=55, b=30),
            )
            st.plotly_chart(fig_c3, use_container_width=True, key=f"{key_prefix}_cat3")

    # [차트 4] 테스트 절차(Procedure/Step) 단계 복잡도 분포 또는 교차 분포
    with r2_c2:
        if proc_col:
            step_series = clean_df[proc_col].apply(_count_procedure_steps)
            vc_step = step_series.value_counts().reset_index()
            vc_step.columns = ["절차 복잡도(단계 수)", "건수"]
            vc_step["비중(%)"] = (vc_step["건수"] / len(clean_df) * 100).round(1)
            vc_step["표시라벨"] = vc_step.apply(lambda r: f"{r['건수']}건 ({r['비중(%)']}%)", axis=1)
            fig_c4 = px.bar(
                vc_step,
                x="절차 복잡도(단계 수)",
                y="건수",
                color="절차 복잡도(단계 수)",
                text="표시라벨",
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"④ [{proc_col}] 수행 절차 단계별(복잡도) 분포 현황",
            )
            fig_c4.update_traces(textposition="outside", textfont_size=13)
            fig_c4.update_layout(height=440, font=dict(size=13), showlegend=False, margin=dict(t=55, b=40))
            st.plotly_chart(fig_c4, use_container_width=True, key=f"{key_prefix}_cat4")


@st.dialog("📊 AI 모델별 상세 시각화 대시보드 & 인사이트 리포트", width="large")
def show_model_visual_popup(m_code: str, res_data: dict, target_df: pd.DataFrame):
    st.markdown(f"## {MODEL_LABELS.get(m_code, m_code)}")
    st.caption(f"⏱️ 분석 소요 시간: **{res_data['elapsed_time']}초** | 핵심 분석 데이터: **{len(target_df)}건**")

    color_df = analyze_color_cell_contexts(target_df)
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("핵심 데이터 건수", f"{len(target_df):,} 건")
    k2.metric("분석 컬럼 수", f"{len(target_df.columns)} 개")
    k3.metric("🎨 색상강조 셀 감지", f"{len(color_df):,} 건")
    k4.metric("AI 응답 속도", f"{res_data['elapsed_time']} 초")

    st.divider()
    render_smart_charts(target_df, key_prefix=f"popup_{m_code}")
    st.divider()

    st.markdown("#### 📝 AI 생성 전체 리포트")
    st.markdown(res_data["report"])

    st.download_button(
        label=f"📥 {MODEL_LABELS.get(m_code, m_code)} 엑셀 리포트 다운로드 (차트 포함 .xlsx)",
        data=res_data["excel_bytes"],
        file_name=f"report_{m_code}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"popup_dl_{m_code}",
        type="primary",
    )


@st.fragment
def render_visualization_fragment(target_df: pd.DataFrame, aux_sheets: dict[str, pd.DataFrame] | None):
    st.subheader("2️⃣ 조회된 데이터 다차원 시각화 & 시트별 종합 집계")
    v_tab1, v_tab2, v_tab3 = st.tabs([
        "🎯 핵심 다차원 대시보드 (유사 제목 자동 묶음 · 긴 문장 가로 차트)",
        "🛠️ 자유 조합 커스텀 차트 (글씨 크기 확대 & 유사 키워드 묶기 지원)",
        "📋 시트 통합 요약 & 유사 주제별 핵심 집계표 (엑셀 2번 시트)",
    ])

    with v_tab1:
        render_smart_charts(target_df, key_prefix="main_dash")

    with v_tab2:
        st.caption(
            "💡 `Precondition`이나 `Title`처럼 문장이 긴 항목을 선택하면 **자동으로 라벨을 깔끔하게 요약하고 글씨와 차트 높이를 키워서** 한눈에 보이게 그려줍니다."
        )
        num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
        all_cols = target_df.columns.tolist()
        if len(target_df) > 0:
            c1, c2, c3, c4 = st.columns(4)
            default_x_idx = next((i for i, c in enumerate(all_cols) if "title" in c.lower() or "precondition" in c.lower()), min(1, len(all_cols) - 1))
            x_col = c1.selectbox("X축 (기준 항목)", options=all_cols, index=default_x_idx, key="frag_x")
            y_options = ["(건수 집계 - COUNT)"] + num_cols
            y_sel = c2.selectbox("Y축 (집계 대상)", options=y_options, index=0 if not num_cols else min(1, len(y_options) - 1), key="frag_y")
            chart_dir = c3.selectbox("차트 방향 (긴 문장은 가로형 추천)", options=["자동 선택 (긴 문장 가로형)", "세로 막대형 (Vertical)", "가로 막대형 (Horizontal)"], key="frag_dir")
            use_topic_cluster = c4.checkbox("🧩 유사 키워드로 묶어서 보기 (1건씩 쪼개질 때 체크)", value=("title" in x_col.lower()), key="frag_cluster")

            # 유사 키워드 클러스터링 적용 여부
            if use_topic_cluster and y_sel == "(건수 집계 - COUNT)":
                clustered = cluster_text_series_by_topic(target_df[x_col], max_groups=10)
                if len(clustered) > 0:
                    fig_custom = px.bar(
                        clustered,
                        x="주제/분류 그룹",
                        y="건수",
                        color="주제/분류 그룹",
                        text="표시라벨",
                        hover_data=["포함된 상세 항목"],
                        color_discrete_sequence=DISTINCT_PALETTE,
                        title=f"[{x_col}] 유사 키워드 그룹별 건수 및 비중(%) 비교",
                    )
                    fig_custom.update_traces(textposition="outside", textfont_size=14)
                    fig_custom.update_layout(height=480, font=dict(size=14), showlegend=False, margin=dict(t=60, b=50))
                    st.plotly_chart(fig_custom, use_container_width=True, key="frag_custom_chart")
            else:
                temp_df = target_df.copy()
                temp_df["_원문_"] = temp_df[x_col].astype(str)
                avg_str_len = temp_df["_원문_"].map(len).mean()
                temp_df["_요약라벨_"] = temp_df["_원문_"].apply(lambda x: shorten_long_label(x, max_len=28))

                if y_sel == "(건수 집계 - COUNT)":
                    custom_df = temp_df.groupby(["_요약라벨_", "_원문_"], as_index=False).size()
                    custom_df.columns = ["_요약라벨_", "_원문_", "건수"]
                    custom_df = custom_df.sort_values(by="건수", ascending=False).head(20)
                    y_plot = "건수"
                else:
                    custom_df = temp_df.groupby(["_요약라벨_", "_원문_"], as_index=False)[y_sel].sum().head(20)
                    y_plot = y_sel

                is_horizontal = (chart_dir == "가로 막대형 (Horizontal)") or (chart_dir.startswith("자동") and avg_str_len > 18)

                if is_horizontal:
                    fig_custom = px.bar(
                        custom_df.iloc[::-1],
                        x=y_plot,
                        y="_요약라벨_",
                        orientation="h",
                        color="_요약라벨_",
                        text_auto=True,
                        hover_data=["_원문_"],
                        color_discrete_sequence=DISTINCT_PALETTE,
                        title=f"[{x_col}]별 {y_plot} 비교 (긴 문장 가독성 최적화 가로 차트)",
                    )
                    fig_custom.update_traces(textposition="outside", textfont_size=14)
                    fig_custom.update_layout(
                        height=max(420, len(custom_df) * 38),
                        font=dict(size=14),
                        xaxis_title=y_plot,
                        yaxis_title=f"{x_col} (마우스 오버 시 전체 문장 표시)",
                        showlegend=False,
                        margin=dict(l=20, r=60, t=60, b=30),
                    )
                else:
                    fig_custom = px.bar(
                        custom_df,
                        x="_요약라벨_",
                        y=y_plot,
                        color="_요약라벨_",
                        text_auto=True,
                        hover_data=["_원문_"],
                        color_discrete_sequence=DISTINCT_PALETTE,
                        title=f"[{x_col}]별 {y_plot} 비교",
                    )
                    fig_custom.update_traces(textposition="outside", textfont_size=14)
                    fig_custom.update_layout(
                        height=480,
                        font=dict(size=14),
                        xaxis_title=x_col,
                        yaxis_title=y_plot,
                        showlegend=False,
                        margin=dict(t=60, b=80),
                    )
                st.plotly_chart(fig_custom, use_container_width=True, key="frag_custom_chart")

    with v_tab3:
        st.caption("유사 제목(Title) 키워드 그룹핑 결과 및 시트 통합 요약이 엑셀 `2_시트통합요약_및_핵심집계` 시트에 그대로 저장됩니다.")
        korean_tables = build_korean_stat_tables(target_df, aux_sheets=aux_sheets)
        for t_title, t_df in korean_tables.items():
            st.markdown(f"**■ {t_title}**")
            st.dataframe(style_dataframe_with_colors(t_df), use_container_width=True, hide_index=True)


@st.fragment
def render_agent_analysis_fragment(
    target_df: pd.DataFrame,
    aux_sheets: dict[str, pd.DataFrame] | None,
    summary_df: pd.DataFrame,
    selected_models: list[str],
    active_filter_text: str,
):
    st.subheader("3️⃣ AI 비서 맞춤 질문 & 에이전트 동시 비교 분석 (Groq vs Gemini vs Ollama 정밀 모드)")

    auto_assistant_q = generate_smart_assistant_question(target_df, aux_sheets, active_filter_text)

    q_mode = st.radio(
        "🤖 질문 생성 방식 선택",
        options=[
            "🤖 AI 비서 자동 맞춤 질문 (유사 주제 그룹핑 + 보조 시트 + 색상 문맥 자동 조합)",
            "✍️ 사용자 직접 질문 입력/수정",
        ],
        horizontal=True,
    )

    if q_mode.startswith("🤖"):
        st.info(f"**💡 AI 비서가 현재 조회 데이터({len(target_df)}건)의 주제 묶음과 구성을 분석해 조합한 질문입니다:**\n\n{auto_assistant_q}")
        question = auto_assistant_q
    else:
        question = st.text_input("💡 에이전트에게 지시할 분석 질문을 입력하세요:", value=auto_assistant_q)

    render_related_links_box(target_df)

    if st.button("🚀 에이전트 정밀 분석 실행 (선택 모델 동시 병렬 실행)", type="primary"):
        if not selected_models:
            st.warning("좌측 사이드바에서 최소 1개 이상의 모델을 선택해주세요.")
        else:
            agent_app = build_agent_graph()
            stat_tables = build_korean_stat_tables(target_df, aux_sheets=aux_sheets)
            stat_summary_str = "\n\n".join(
                f"■ {k}\n{v.to_string(index=False)}" for k, v in stat_tables.items()
            )
            sample_rows_df = target_df.head(35)
            data_info = (
                f"[조회 조건]: {active_filter_text or '핵심 데이터 시트 전체'}\n"
                f"[핵심 분석 대상 건수]: {len(target_df)}건\n"
                f"[보조 시트 목록]: {list(aux_sheets.keys()) if aux_sheets else '없음'}\n"
                f"[컬럼 목록]: {list(target_df.columns)}\n\n"
                f"[파이썬 정밀 집계 및 주제 그룹핑 요약표]\n{stat_summary_str}\n\n"
                f"[핵심 데이터 상세 행 내역]\n{sample_rows_df.to_string(index=False)}"
            )

            new_results = {}
            status_box = st.status("⚡ 선택한 AI 모델들을 동시에 병렬 실행하고 있습니다...", expanded=True)

            def _run_single_model(m_code: str):
                res = agent_app.invoke({
                    "question": question,
                    "mode": m_code,
                    "data_summary": data_info,
                    "analysis_result": {},
                    "report": "",
                    "elapsed_time": 0.0,
                })
                excel_bytes = export_analysis_to_excel(
                    target_df, summary_df, res["report"], res["elapsed_time"], m_code, aux_sheets=aux_sheets
                )
                res["excel_bytes"] = excel_bytes
                return m_code, res

            with ThreadPoolExecutor(max_workers=len(selected_models)) as executor:
                future_map = {executor.submit(_run_single_model, m): m for m in selected_models}
                for future in as_completed(future_map):
                    m_done, res_done = future.result()
                    new_results[m_done] = res_done
                    status_box.write(f"✅ **{MODEL_LABELS.get(m_done, m_done)}** 분석 완료! (`{res_done['elapsed_time']}초`)")

            status_box.update(label="🎉 모든 모델 분석 완료!", state="complete", expanded=False)
            st.session_state["saved_results"] = new_results
            st.session_state["last_analyzed_df"] = target_df

    saved = st.session_state.get("saved_results", {})
    if saved:
        display_models = [m for m in selected_models if m in saved] or list(saved.keys())
        popup_df = st.session_state.get("last_analyzed_df") if st.session_state.get("last_analyzed_df") is not None else target_df

        st.markdown("##### ⚖️ AI 모델별 성능(응답 속도) 및 핵심 진단 한눈에 비교")
        matrix_rows = []
        for m_code in display_models:
            r_item = saved[m_code]
            first_line = next(
                (l.strip("> ") for l in r_item["report"].splitlines() if any(k in l for k in ["확인 필요", "이상 없음", "점검 안내"])),
                "본문 참조",
            )
            matrix_rows.append({
                "AI 모델": MODEL_LABELS.get(m_code, m_code),
                "응답 소요 시간": f"{r_item['elapsed_time']} 초",
                "자동 판정 요약": first_line,
            })
        st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True, hide_index=True)

        combined_excel_bytes = export_combined_analysis_to_excel(popup_df, saved, MODEL_LABELS, aux_sheets=aux_sheets)
        top_c1, top_c2 = st.columns([2, 1])
        with top_c1:
            st.caption("👆 **각 모델의 제목 버튼을 클릭하면 상세 다차원 시각화 팝업창이 열립니다.**")
        with top_c2:
            st.download_button(
                label="📚 전체 모델 종합 비교 엑셀 추출 (주제그룹핑·차트 완비 .xlsx)",
                data=combined_excel_bytes,
                file_name="report_ALL_models_combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_combined_master",
                type="primary",
                use_container_width=True,
            )

        cols = st.columns(len(display_models))
        for idx, m_code in enumerate(display_models):
            res_item = saved[m_code]
            with cols[idx]:
                if st.button(f"📊 {MODEL_LABELS.get(m_code, m_code)}\n(클릭 시 시각화 팝업 열기 🔍)", key=f"title_btn_{m_code}", use_container_width=True):
                    show_model_visual_popup(m_code, res_item, popup_df)

                st.info(f"⏱️ 소요 시간: **{res_item['elapsed_time']}초**")
                st.markdown(res_item["report"])

                st.download_button(
                    label=f"📥 {m_code} 개별 엑셀 추출 (주제그룹핑·차트 포함 .xlsx)",
                    data=res_item["excel_bytes"],
                    file_name=f"report_{m_code}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_persist_{m_code}_{idx}",
                    use_container_width=True,
                )


with st.sidebar:
    st.header("⚙️ 실행 및 비교 모델 설정")
    run_type = st.radio("실행 방식 선택", options=["다중 모델 동시 비교 (PoC)", "단일 모델 실행"], index=0)

    if run_type == "단일 모델 실행":
        selected_models = [
            st.selectbox("사용할 AI 모델", options=list(MODEL_LABELS.keys()), format_func=lambda x: MODEL_LABELS[x])
        ]
    else:
        selected_models = st.multiselect(
            "비교할 AI 모델들을 선택하세요 (병렬 동시 실행)",
            options=list(MODEL_LABELS.keys()),
            default=["cloud_groq", "cloud_gemini", "local_ollama"],
            format_func=lambda x: MODEL_LABELS[x],
        )

    st.divider()
    ollama_models = get_installed_ollama_models()
    with st.expander("🖥️ 로컬 Ollama 상세 설정 & 모델 안내 (클릭)", expanded=False):
        st.markdown(
            "**💡 로컬 Ollama 정밀 모드**\n"
            "- 파이썬이 먼저 시트별 통계와 유사 제목 그룹핑을 0.01초 만에 계산하고, Ollama가 리포트 문장을 작성합니다."
        )
        chosen_ollama = st.selectbox("구동할 로컬 모델 변경", options=ollama_models, index=0)
        os.environ["LOCAL_MODEL_NAME"] = chosen_ollama

    if st.session_state["saved_results"]:
        if st.button("🗑️ 이전 분석 결과 초기화", use_container_width=True):
            st.session_state["saved_results"] = {}
            st.rerun()

uploaded_file = st.file_uploader("📂 사내 데이터 파일(CSV, Excel) 업로드 (다중 시트 분리 & 유사 제목 자동 그룹핑)", type=["csv", "xlsx"])

if uploaded_file is not None:
    raw_df_main, sheet_names, sheet_counts, aux_sheets = load_and_sync_to_db(uploaded_file, selected_sheet="AUTO_MAIN")

    if len(sheet_names) > 1:
        main_sheet_names = [s for s in sheet_names if s not in aux_sheets]
        aux_sheet_names = list(aux_sheets.keys())
        st.info(
            f"📑 **다중 시트 엑셀 스마트 분리 완료!**\n"
            f"- **🎯 핵심 분석 데이터 시트:** `{', '.join(main_sheet_names)}` (총 **{len(raw_df_main)}건**)\n"
            f"- **📋 요약·환경·설명 보조 시트:** `{', '.join(aux_sheet_names) if aux_sheet_names else '없음'}`"
        )

        if aux_sheets:
            with st.expander(f"🔎 보조·요약 시트({', '.join(aux_sheet_names)}) 내용 바로 보기 (클릭)", expanded=False):
                aux_cols = st.columns(len(aux_sheets))
                for idx_a, (a_name, a_df) in enumerate(aux_sheets.items()):
                    with aux_cols[idx_a]:
                        st.markdown(f"**📄 [{a_name}] 시트 ({len(a_df)}행)**")
                        st.dataframe(style_dataframe_with_colors(a_df), use_container_width=True, hide_index=True)

        sheet_options = ["AUTO_MAIN"] + sheet_names + ["ALL_MERGE"]
        chosen_sheet = st.radio(
            "조회할 시트 모드 선택:",
            options=sheet_options,
            index=0,
            format_func=lambda x: (
                f"🎯 핵심 데이터 시트만 정확히 조회 ({len(raw_df_main)}건 · 추천)"
                if x == "AUTO_MAIN"
                else (
                    f"📚 모든 시트 강제 합쳐서 보기 ({sum(sheet_counts.values())}건)"
                    if x == "ALL_MERGE"
                    else f"📄 {x} 시트 단독 보기 ({sheet_counts.get(x, 0)}건)"
                )
            ),
            horizontal=True,
        )
        if chosen_sheet == "AUTO_MAIN":
            raw_df = raw_df_main
        else:
            raw_df, _, _, _ = load_and_sync_to_db(uploaded_file, selected_sheet=chosen_sheet)
    else:
        raw_df = raw_df_main

    # 1. 자연어 쿼리 자동 생성 및 SQLite DB 조회
    st.subheader("1️⃣ 자연어 기반 데이터 수집 및 SQLite DB 자동 조회")
    tab_nl, tab_kw, tab_sql = st.tabs(["🤖 자연어 → SQL 자동 변환", "🔍 키워드 조합 필터", "🗄️ 수동 SQL 편집기 (.db / .sql 추출)"])

    active_filter_text = ""

    with tab_nl:
        st.caption(
            f"테이블(`{TABLE_NAME}`) 컬럼: `{list(raw_df.columns)}` | "
            "입력 예시: **'Title칸에 문서 들어간거 조회해줘'**, **'목차이나 검색 조회해줘'**, **'배달만 조회해줘'**, **'쿠폰이나 내역 조회'**"
        )
        nl_question = st.text_input("조회 조건을 자연어로 입력하세요 (특정 컬럼 지정 및 부분 단어 검색 지원):", value="")
        auto_sql = build_smart_sql(nl_question, raw_df, TABLE_NAME)
        st.code(auto_sql, language="sql")
        try:
            df = run_sql_query(auto_sql)
            if len(df) > 0:
                st.success(f"✅ SQL 자동 실행 완료: 전체 **{len(raw_df)}건** 중 조건에 일치하는 **{len(df)}건**이 조회되었습니다.")
                st.dataframe(style_dataframe_with_colors(df), use_container_width=True, hide_index=True)
            else:
                st.warning(
                    f"⚠️ 입력하신 조건(`{nl_question}`)으로 SQL 쿼리를 실행했으나 현재 파일에서 일치하는 데이터가 **0건**입니다. "
                    "(아래 차트와 분석은 전체 원본 데이터를 기준으로 유지됩니다.)"
                )
            if nl_question.strip():
                active_filter_text = nl_question.strip()
        except Exception as e:
            st.error(f"SQL 실행 오류: {e}")
            df = raw_df

    with tab_kw:
        c1, c2 = st.columns([3, 1])
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,` 구분, 예: `문서, 목차` 또는 `배달, 쿠폰`)", value="")
        match_type = c2.radio("조건", options=["OR (하나라도 포함)", "AND (모두 포함)"], horizontal=True)
        if kw_input.strip():
            df = filter_by_keywords(raw_df, kw_input, match_all=match_type.startswith("AND"))
            active_filter_text = f"키워드({kw_input.strip()})"
        st.dataframe(style_dataframe_with_colors(df), use_container_width=True, hide_index=True)

    with tab_sql:
        sql_input = st.text_area("수동 SQL 입력", value=auto_sql, height=80)
        if st.button("▶️ 수동 SQL 실행"):
            try:
                df = run_sql_query(sql_input)
                st.success(f"쿼리 실행 성공: {len(df)}건 조회됨")
                st.dataframe(style_dataframe_with_colors(df), use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"오류: {e}")

        if DB_PATH.exists():
            dl_c1, dl_c2 = st.columns(2)
            with dl_c1:
                with open(DB_PATH, "rb") as f:
                    st.download_button(
                        label="💾 SQLite DB 바이너리 파일 (.db) 다운로드",
                        data=f,
                        file_name="internal_analysis.db",
                        mime="application/octet-stream",
                        use_container_width=True,
                    )
            with dl_c2:
                sql_dump_bytes = export_db_to_sql_script(current_sql=sql_input)
                st.download_button(
                    label="📝 Notepad++ 조회용 SQL 덤프 (.sql) 다운로드",
                    data=sql_dump_bytes,
                    file_name="internal_analysis_dump.sql",
                    mime="application/sql",
                    type="primary",
                    use_container_width=True,
                )

    target_df = df if len(df) > 0 else raw_df
    summary_df = target_df.describe(include="all").fillna("")
    st.divider()

    render_visualization_fragment(target_df, aux_sheets)
    st.divider()
    render_agent_analysis_fragment(target_df, aux_sheets, summary_df, selected_models, active_filter_text)