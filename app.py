# app.py
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

from data_agent.config import (
    PRIMARY_COLOR,
    DISTINCT_PALETTE,
    DEFAULT_DEFECT_THRESHOLD,
    MODEL_LABELS,
)
from data_agent.graph import build_agent_graph
from data_agent.excel_exporter import (
    export_analysis_to_excel,
    export_combined_analysis_to_excel,
    build_korean_stat_tables,
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
from data_agent.tools.visualizer import (
    render_validation_banner,
    render_smart_charts,
)

CONFIG_DIR = Path(".streamlit")
CONFIG_FILE = CONFIG_DIR / "config.toml"
if not CONFIG_FILE.exists():
    CONFIG_DIR.mkdir(exist_ok=True)
    CONFIG_FILE.write_text(f'[theme]\nprimaryColor = "{PRIMARY_COLOR}"\n', encoding="utf-8")

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

st.title("📊 자연어 기반 사내 데이터 분석 Agent")

if "saved_results" not in st.session_state:
    st.session_state["saved_results"] = {}
if "last_analyzed_df" not in st.session_state:
    st.session_state["last_analyzed_df"] = None
if "last_aux_sheets" not in st.session_state:
    st.session_state["last_aux_sheets"] = {}
if "pipeline_logs" not in st.session_state:
    st.session_state["pipeline_logs"] = []


# =====================================================================
# ⚡ 실시간 파이프라인 가동현황 (1번 ~ 5번 배지)
# =====================================================================
def render_pipeline_status(has_file: bool, has_query_df: bool, saved_results: dict):
    has_analysis = bool(saved_results)
    has_verified_report = has_analysis and any(bool(v.get("report")) for v in saved_results.values())
    has_excel_pkg = has_analysis and any(bool(v.get("excel_bytes")) for v in saved_results.values())

    steps = [
        (1, "1. 문서·시트·색상 스캔", has_file, "다중 시트·양식 검증 완료" if has_file else "파일 업로드 대기"),
        (2, "2. SQLite DB·정밀 집계", has_query_df, "전체 건수 DB 적재·통계 완료" if has_query_df else "데이터 집계 대기"),
        (3, "3. 3대 AI 병렬 추론", has_analysis, f"{len(saved_results)}개 모델 추론 완료" if has_analysis else "분석 실행 대기"),
        (4, "4. 리포트 품질 검증·보정", has_verified_report, "수치·양식 자동 정제 완료" if has_verified_report else "리포트 검증 대기"),
        (5, "5. 통합 리포트·엑셀 완료", has_excel_pkg, "다중 시트·차트 패키징 완료" if has_excel_pkg else "엑셀 생성 대기"),
    ]

    cards_html = []
    for num, title, is_done, desc in steps:
        if is_done:
            bg, border, title_col, sub_col, icon = "#ecfdf5", "#10b981", "#065f46", "#047857", "✅"
        else:
            bg, border, title_col, sub_col, icon = "#f8fafc", "#cbd5e1", "#64748b", "#94a3b8", "⏳"

        cards_html.append(
            f'<div style="flex:1; min-width:155px; background:{bg}; border:2px solid {border}; '
            f'border-radius:10px; padding:10px 12px; text-align:center; box-shadow:0 1px 3px rgba(0,0,0,0.05);">'
            f'<div style="font-size:0.86rem; font-weight:700; color:{title_col};">{icon} {title}</div>'
            f'<div style="font-size:0.74rem; color:{sub_col}; margin-top:4px; font-weight:600;">{desc}</div>'
            f'</div>'
        )

    st.markdown(
        '<div style="margin: 4px 0 14px 0;">'
        '<div style="font-size:0.92rem; font-weight:700; color:#1e293b; margin-bottom:8px;">'
        '⚡ 실시간 파이프라인 가동현황</div>'
        f'<div style="display:flex; gap:10px; flex-wrap:wrap;">{"".join(cards_html)}</div>'
        '</div>',
        unsafe_allow_html=True,
    )

    if st.session_state.get("pipeline_logs"):
        with st.expander("📋 파이프라인 단계별 상세 실행 로그 보기 (클릭)", expanded=False):
            for log_line in st.session_state["pipeline_logs"]:
                st.markdown(f"- {log_line}")


def style_dataframe_with_colors(df: pd.DataFrame):
    def _highlight_cell(val):
        s = str(val)
        if "빨강" in s or "🔴" in s or "FAIL" in s.upper():
            return "background-color: #FEE2E2; color: #991B1B; font-weight: bold;"
        if "초록" in s or "🟢" in s:
            return "background-color: #DCFCE7; color: #166534; font-weight: bold;"
        if "노랑" in s or "주황" in s or "🟡" in s or "🟠" in s:
            return "background-color: #FEF9C3; color: #854D0E; font-weight: bold;"
        if "🎨" in s or "🔵" in s:
            return "background-color: #DBEAFE; color: #1E40AF; font-weight: bold;"
        return ""

    if len(df) <= 300:
        return df.style.map(_highlight_cell)
    return df


def build_smart_sql(nl_query: str, df: pd.DataFrame, table_name: str) -> str:
    """
    [전체 건수 조회 완벽 복원]
    - 검색창이 비어 있거나 '전체조회', '전체', '모두' 입력 시 LIMIT 없이 전체 데이터(예: 592건 전체)를 조회합니다.
    - '1000건 조회', '상위 50개', '100건만'처럼 건수 제한을 입력하면 정확히 LIMIT N으로 변환합니다.
    """
    if not nl_query or not nl_query.strip():
        return f"SELECT * FROM {table_name};"

    q = nl_query.strip()
    if q.replace(" ", "") in ("전체", "전체조회", "전체데이터", "모두조회", "다보여줘", "전체보여줘", "all"):
        return f"SELECT * FROM {table_name};"

    # 1. 건수 제한(LIMIT N) 표현 먼저 추출 (예: "1000건 조회", "상위 50개", "limit 200")
    limit_val = None
    limit_match = re.search(r"(?:상위\s*|최대\s*|limit\s*)?(\d+)\s*(?:건|개|행|줄)(?:\s*만|\s*조회|\s*보여|\s*출력)?", q, re.IGNORECASE)
    if not limit_match:
        limit_match = re.search(r"limit\s+(\d+)", q, re.IGNORECASE)
    if limit_match:
        limit_val = int(limit_match.group(1))
        # 건수 표현을 제거하여 아래 숫자 비교 조건(WHERE)과 충돌하지 않게 함
        q_without_limit = (q[:limit_match.start()] + " " + q[limit_match.end():]).strip()
    else:
        q_without_limit = q

    # 건수만 입력한 경우 (예: "1000건 조회", "500건") 바로 LIMIT 쿼리 반환
    rem_cleaned = re.sub(r"(전체|조회|해줘|보여줘|출력|데이터|만|에서|중)", "", q_without_limit).strip()
    if not rem_cleaned and limit_val is not None:
        return f"SELECT * FROM {table_name} LIMIT {limit_val};"

    # 2. 배경색 키워드 검색
    if any(k in q_without_limit for k in ["빨강", "빨간", "초록", "노랑", "배경", "색상", "색깔"]):
        color_kw = "빨강" if ("빨강" in q_without_limit or "빨간" in q_without_limit) else (
            "초록" if "초록" in q_without_limit else ("노랑" if "노랑" in q_without_limit else "강조")
        )
        or_likes = [f'"{col}" LIKE \'%{color_kw}%\'' for col in df.columns]
        sql = f"SELECT * FROM {table_name} WHERE " + " OR ".join(or_likes)
        return f"{sql} LIMIT {limit_val};" if limit_val else f"{sql};"

    # 3. 범주형 컬럼 값 매칭 (같은 컬럼 내 복수 값은 OR로 연결)
    col_matches: dict[str, list[str]] = {}
    str_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    for col in str_cols:
        matched_vals = []
        for val in df[col].dropna().astype(str).unique():
            val_clean = re.sub(r"\s*\[[^\]]+\]", "", val).strip()
            if not val_clean or val_clean == "-" or len(val_clean) > 45:
                continue
            if val_clean.lower() in q_without_limit.lower() or f"{val_clean}{col}".lower() in q_without_limit.lower():
                matched_vals.append(val_clean)
        if matched_vals:
            col_matches[col] = list(dict.fromkeys(matched_vals))

    where_clauses = []
    for col, vals in col_matches.items():
        if len(vals) == 1:
            where_clauses.append(f'"{col}" LIKE \'%{vals[0]}%\'')
        else:
            or_part = " OR ".join(f'"{col}" LIKE \'%{v}%\'' for v in vals)
            where_clauses.append(f"({or_part})")

    # 4. 수치형 컬럼 범위 조건 매칭 (예: "생산량 1200 이상", "불량수 20 초과")
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    nums = re.findall(r"\d+(?:\.\d+)?", q_without_limit)
    if nums and num_cols:
        target_num = float(nums[0])
        target_col = next((col for col in num_cols if col.lower() in q_without_limit.lower()), None)
        if target_col is None:
            for col in num_cols:
                if ("불량" in q_without_limit and "불량" in col) or ("생산" in q_without_limit and "생산" in col) or ("가동" in q_without_limit and "가동" in col):
                    target_col = col
                    break
            if target_col is None:
                # 데이터 실제 최소~최대 범위에 속하는 수치 컬럼 자동 추론
                for col in num_cols:
                    s_num = df[col].dropna()
                    if len(s_num) > 0 and float(s_num.min()) <= target_num <= float(s_num.max()):
                        target_col = col
                        break
        if target_col:
            num_str = nums[0]
            if any(k in q_without_limit for k in ["넘", "초과", "보다 큰", "많은", "높은", ">"]):
                op = ">=" if "이상" in q_without_limit else ">"
            elif any(k in q_without_limit for k in ["이상", "부터", ">="]):
                op = ">="
            elif any(k in q_without_limit for k in ["미만", "보다 작은", "적은", "낮은", "<"]):
                op = "<"
            elif any(k in q_without_limit for k in ["이하", "이내", "까지", "<="]):
                op = "<="
            else:
                op = "="
            where_clauses.append(f'"{target_col}" {op} {num_str}')

    sql = f"SELECT * FROM {table_name}"
    if where_clauses:
        sql += " WHERE " + " AND ".join(where_clauses)
    if limit_val is not None:
        sql += f" LIMIT {limit_val}"
    sql += ";"
    return sql


def generate_smart_assistant_question(
    target_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    user_filter_text: str,
    defect_threshold: float = DEFAULT_DEFECT_THRESHOLD,
) -> str:
    if len(target_df) == 0:
        return "현재 데이터의 핵심 현황, 이상·주의 항목, 원인 가설 및 실무 조치 제안을 정리해 줘."

    all_cols = target_df.columns.tolist()
    cols_lower = " ".join(c.lower() for c in all_cols)
    scope_parts = []
    if user_filter_text and user_filter_text.strip():
        scope_parts.append(f"검색조건('{user_filter_text.strip()}')")

    if "시트명" in all_cols:
        sheets = target_df["시트명"].unique().tolist()
        scope_parts.append(f"대상 시트 {len(sheets)}개({', '.join(map(str, sheets[:3]))})")

    stat_snippets = [f"총 {len(target_df):,}건"]
    scope_title = " · ".join(scope_parts) if scope_parts else "전체 조회 데이터"

    if "생산량" in all_cols and "불량수" in all_cols:
        p_sum = int(target_df["생산량"].sum())
        d_sum = int(target_df["불량수"].sum())
        rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
        stat_snippets.append(f"총 생산량 {p_sum:,}개, 총 불량수 {d_sum:,}개(불량률 {rate:.2f}%)")
        stat_str = ", ".join(stat_snippets)
        return (
            f"[AI 비서 자동 맞춤 질문] {scope_title} ({stat_str}) 데이터를 바탕으로 "
            f"① 라인·일자별 핵심 추세 및 불량 비중(%), ② 사내 기준({defect_threshold}%) 초과 이상 구간, "
            f"③ 공정·설비·작업 관점 원인 가설, ④ 담당 부서별 즉시 조치(24h) 및 근본 대책(1개월)을 정밀 분석해 줘."
        )

    if any(k in cols_lower for k in ["procedure", "pass/fail", "actual", "precondition", "pre-condition", "domain", "sub system", "tc"]):
        row_text = target_df.astype(str).agg(" ".join, axis=1)
        red_cnt = int(row_text.str.contains("빨강|불합격|FAIL|주의", case=False, regex=True).sum())
        if red_cnt > 0:
            stat_snippets.append(f"🔴색상강조/주의 표시 {red_cnt}건 감지")
        for c_cand in ["Domain", "Sub System", "App", "Automated"]:
            if c_cand in all_cols:
                vals = [str(v) for v in target_df[c_cand].unique().tolist() if str(v) != "-"]
                if vals:
                    stat_snippets.append(f"{c_cand}[{', '.join(vals[:3])}]")
        stat_str = ", ".join(stat_snippets)
        return (
            f"[AI 비서 자동 맞춤 질문] {scope_title} ({stat_str}) TC 검증 데이터를 바탕으로 "
            f"① 도메인(Domain)·서브시스템(Sub System/App/IP)별 분포와 비중(%), "
            f"② 주요 기능 시나리오(Title 키워드 묶음 및 수행 특성) 집중도와 대표 TC(ID·Title), "
            f"③ 판정 현황(정상/🔴빨강강조 주의/미기입 공란) 및 예외·경계값 커버리지 리스크, "
            f"④ QA·개발팀의 즉시 조치(24h) 및 검증 표준화 대책(1개월)을 마크다운 표 없이 깔끔한 불릿으로 분석해 줘."
        )

    stat_str = ", ".join(stat_snippets)
    return (
        f"[AI 비서 자동 맞춤 질문] {scope_title} ({stat_str}) 데이터를 바탕으로 "
        f"① 시트·분류별 핵심 현황 및 비중(%), ② 색상 강조(🎨주의) 또는 미기입 항목 점검, "
        f"③ 주요 리스크 및 원인 가설, ④ 실무 조치 제안(24h/1개월)을 정리해 줘."
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
    elif any(c in all_cols for c in ["Sub System", "Domain", "App"]):
        for c in ["Domain", "Sub System", "App"]:
            if c in all_cols:
                vals = [str(v) for v in target_df[c].astype(str).unique() if str(v) != "-"]
                keywords.extend(vals[:1])
        keywords.append("임베디드 SW 검증 테스트")
    else:
        cat_cols = target_df.select_dtypes(include=["object", "category"]).columns.tolist()
        if cat_cols:
            vals = [str(v) for v in target_df[cat_cols[0]].astype(str).unique() if str(v) != "-"]
            keywords.extend(vals[:2])

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


@st.dialog("📊 AI 모델별 상세 시각화 대시보드 & 인사이트 리포트", width="large")
def show_model_visual_popup(m_code: str, res_data: dict, target_df: pd.DataFrame, defect_threshold: float):
    st.markdown(f"## {MODEL_LABELS.get(m_code, m_code)}")
    st.caption(f"⏱️ 분석 소요 시간: **{res_data['elapsed_time']}초** | 분석 대상 데이터: **{len(target_df):,}건**")

    render_smart_charts(target_df, key_prefix=f"popup_{m_code}", defect_threshold=defect_threshold)
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
def render_visualization_fragment(
    target_df: pd.DataFrame,
    aux_sheets: dict[str, pd.DataFrame],
    defect_threshold: float,
):
    st.subheader("2️⃣ 조회된 데이터 다차원 시각화 & 한글 기초통계 집계")
    v_tab1, v_tab2, v_tab3 = st.tabs([
        "🎯 핵심 다차원 대시보드 (TOP 순위·시나리오 그룹핑·비중 자동 대응)",
        "🛠️ 자유 조합 커스텀 차트 (모든 엑셀/CSV/문서 범용)",
        "📋 한글 기초통계 & 시트통합·판정·배경색 문맥 집계표",
    ])

    with v_tab1:
        render_smart_charts(target_df, key_prefix="main_dash", defect_threshold=defect_threshold)

    with v_tab2:
        num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
        all_cols = target_df.columns.tolist()
        if len(target_df) > 0:
            c1, c2, c3, c4 = st.columns(4)
            default_x = next((i for i, c in enumerate(all_cols) if c in ["Sub System", "Domain", "App", "라인"]), 0)
            x_col = c1.selectbox("X축 (기준 항목)", options=all_cols, index=default_x, key="frag_x")
            y_options = ["(건수 집계 - COUNT)"] + num_cols
            y_sel = c2.selectbox("Y축 (집계 대상)", options=y_options, index=0 if not num_cols else min(1, len(y_options) - 1), key="frag_y")
            color_options = ["(색상 분류 없음)"] + all_cols
            color_sel = c3.selectbox("색상 구분 (세부 분류)", options=color_options, index=0, key="frag_color")
            agg_func = c4.selectbox(
                "집계 방식",
                options=["건수 (COUNT)", "합계 (SUM)", "평균 (MEAN)"] if num_cols else ["건수 (COUNT)"],
                key="frag_agg",
            )

            color_col = None if color_sel == "(색상 분류 없음)" else color_sel
            grp_cols = [x_col] + ([color_col] if color_col and color_col != x_col else [])

            if y_sel == "(건수 집계 - COUNT)" or "건수" in agg_func:
                custom_df = target_df.groupby(grp_cols, as_index=False).size()
                custom_df.columns = grp_cols + ["건수"]
                y_plot = "건수"
            elif "합계" in agg_func:
                custom_df = target_df.groupby(grp_cols, as_index=False)[y_sel].sum()
                y_plot = y_sel
            else:
                custom_df = target_df.groupby(grp_cols, as_index=False)[y_sel].mean().round(2)
                y_plot = y_sel

            custom_df = custom_df.sort_values(by=y_plot, ascending=False).head(30)
            fig_custom = px.bar(
                custom_df, x=x_col, y=y_plot, color=color_col or x_col, barmode="group",
                text_auto=True, color_discrete_sequence=DISTINCT_PALETTE,
                title=f"[{x_col}]별 {y_plot} 비교 (상위 30개)",
            )
            fig_custom.update_traces(textposition="outside", cliponaxis=False)
            st.plotly_chart(fig_custom, use_container_width=True, key="frag_custom_chart")

    with v_tab3:
        korean_tables = build_korean_stat_tables(target_df, aux_sheets=aux_sheets)
        for t_title, t_df in korean_tables.items():
            st.markdown(f"**■ {t_title}**")
            st.dataframe(style_dataframe_with_colors(t_df), use_container_width=True, hide_index=True)


def render_agent_analysis_section(
    target_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    selected_models: list[str],
    active_filter_text: str,
    aux_sheets: dict[str, pd.DataFrame],
    defect_threshold: float,
):
    st.subheader("3️⃣ AI 비서 맞춤 질문 & 에이전트 동시 비교 분석 (Groq vs Gemini vs Ollama 정밀 모드)")

    auto_assistant_q = generate_smart_assistant_question(
        target_df, raw_df, active_filter_text, defect_threshold=defect_threshold
    )

    q_mode = st.radio(
        "🤖 질문 생성 방식 선택",
        options=[
            "🤖 AI 비서 자동 맞춤 질문 (조회된 시트·컬럼·배경색/수치 자동 조합)",
            "✍️ 사용자 직접 질문 입력/수정",
        ],
        horizontal=True,
    )

    if q_mode.startswith("🤖"):
        st.info(f"**💡 AI 비서가 현재 조회된 데이터({len(target_df):,}건)에 맞춰 조합한 분석 질문입니다:**\n\n{auto_assistant_q}")
        question = auto_assistant_q
    else:
        question = st.text_input("💡 에이전트에게 지시할 분석 질문을 입력하세요:", value=auto_assistant_q)

    render_related_links_box(target_df)

    if st.button("🚀 에이전트 정밀 분석 실행 (선택 모델 동시 병렬 실행)", type="primary"):
        if not selected_models:
            st.warning("좌측 사이드바에서 최소 1개 이상의 모델을 선택해주세요.")
        else:
            agent_app = build_agent_graph()
            stat_summary_str = "\n\n".join(
                f"■ {k}\n{v.to_string(index=False)}"
                for k, v in build_korean_stat_tables(target_df, aux_sheets=aux_sheets).items()
            )
            sample_rows_df = target_df.head(50)
            data_info = (
                f"[조회 조건]: {active_filter_text or '전체 시트 데이터'}\n"
                f"[조회 데이터 건수]: {len(target_df)}건 (전체 원본 {len(raw_df)}건)\n"
                f"[사내 품질 기준값]: {defect_threshold}%\n"
                f"[컬럼 목록]: {list(target_df.columns)}\n\n"
                f"[파이썬 정밀 집계표]\n{stat_summary_str}\n\n"
                f"[상세 행 샘플 데이터 (상위 50건)]\n{sample_rows_df.to_string(index=False)}"
            )

            new_results = {}
            logs = [
                f"**[1~2단계 완료]** 총 `{len(target_df):,}건` 데이터 로드 및 파이썬 다차원 통계 집계 완료",
                f"**[3단계 가동]** 선택 모델 `{len(selected_models)}개` 병렬 추론 시작",
            ]
            status_box = st.status("⚡ 선택한 AI 모델(정밀 모드)들을 동시에 병렬 실행하고 있습니다...", expanded=True)

            def _run_single_model(m_code: str):
                res = agent_app.invoke({
                    "question": question,
                    "mode": m_code,
                    "data_summary": data_info,
                    "defect_threshold": defect_threshold,
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
                    msg = f"✅ **{MODEL_LABELS.get(m_done, m_done)}** 추론 · 양식 검증 · 엑셀 패키징 완료 (`{res_done['elapsed_time']}초`)"
                    status_box.write(msg)
                    logs.append(msg)

            logs.append("**[4단계 완료]** 리포트 수치 붙음 방지 · `<br>` 제거 · 4대 섹션 완결성 검증 완료")
            logs.append("**[5단계 완료]** 통합 비교 대시보드 및 다중 시트 엑셀(`.xlsx`) 패키징 완료")
            status_box.update(label="🎉 1~5단계 전 파이프라인 가동 및 정밀 분석 완료!", state="complete", expanded=False)

            st.session_state["saved_results"] = new_results
            st.session_state["last_analyzed_df"] = target_df
            st.session_state["last_aux_sheets"] = aux_sheets
            st.session_state["pipeline_logs"] = logs
            st.rerun()

    saved = st.session_state.get("saved_results", {})
    if saved:
        display_models = [m for m in selected_models if m in saved] or list(saved.keys())
        popup_df = st.session_state.get("last_analyzed_df") if st.session_state.get("last_analyzed_df") is not None else target_df
        saved_aux = st.session_state.get("last_aux_sheets", aux_sheets)

        st.markdown("##### ⚖️ AI 모델별 성능(응답 속도) 및 핵심 진단 한눈에 비교")
        matrix_rows = []
        for m_code in display_models:
            r_item = saved[m_code]
            first_line = next(
                (l.strip("#*- >") for l in r_item["report"].splitlines() if l.strip() and not l.strip().startswith("#")),
                "본문 참조",
            )
            matrix_rows.append({
                "AI 모델": MODEL_LABELS.get(m_code, m_code),
                "응답 소요 시간": f"{r_item['elapsed_time']} 초",
                "리포트 핵심 요약 (첫 줄)": first_line[:90],
            })
        st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True, hide_index=True)

        combined_excel_bytes = export_combined_analysis_to_excel(popup_df, saved, MODEL_LABELS, aux_sheets=saved_aux)
        top_c1, top_c2 = st.columns([2, 1])
        with top_c1:
            st.caption("👆 **각 모델의 제목 버튼을 클릭하면 상세 다차원 시각화 팝업창이 열립니다.**")
        with top_c2:
            st.download_button(
                label="📚 전체 모델 종합 비교 엑셀 추출 (줄바꿈·기초통계·차트 완비 .xlsx)",
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
                    show_model_visual_popup(m_code, res_item, popup_df, defect_threshold)

                st.info(f"⏱️ 소요 시간: **{res_item['elapsed_time']}초**")
                st.markdown(res_item["report"])

                st.download_button(
                    label=f"📥 {m_code} 개별 엑셀 추출 (기초통계·차트 포함 .xlsx)",
                    data=res_item["excel_bytes"],
                    file_name=f"report_{m_code}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_persist_{m_code}_{idx}",
                    use_container_width=True,
                )


# =====================================================================
# ⚙️ 사이드바 설정 (W47 사내 기준값 조절 기능 포함)
# =====================================================================
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
    st.subheader("🎚️ 사내 품질 기준값 설정 ")
    defect_threshold = st.slider(
        "목표 불량률 / 관리 임계치 (%)",
        min_value=0.1,
        max_value=10.0,
        value= float(DEFAULT_DEFECT_THRESHOLD),
        step=0.1,
        help="부서별 공정 기준에 맞춰 조절하면 차트 기준선과 AI 판정이 즉시 연동됩니다.",
    )

    st.divider()
    ollama_models = get_installed_ollama_models()
    with st.expander("🖥️ 로컬 Ollama 상세 설정 & 모델 안내 (클릭)", expanded=False):
        chosen_ollama = st.selectbox("구동할 로컬 모델 변경", options=ollama_models, index=0)
        os.environ["LOCAL_MODEL_NAME"] = chosen_ollama

    if st.session_state["saved_results"]:
        if st.button("🗑️ 이전 분석 결과 초기화", use_container_width=True):
            st.session_state["saved_results"] = {}
            st.session_state["pipeline_logs"] = []
            st.rerun()

# =====================================================================
# 📂 메인 화면: 파일 업로드 및 1~3단계 워크플로우
# =====================================================================
uploaded_file = st.file_uploader(
    "📂 사내 데이터·문서 파일(CSV, Excel, Word .docx, TXT) 업로드 (다중 시트 & 셀 배경색 자동 인식)",
    type=["csv", "xlsx", "docx", "txt"],
)

if uploaded_file is None:
    render_pipeline_status(has_file=False, has_query_df=False, saved_results={})
    st.info("👆 분석할 사내 파일(.csv, .xlsx, .docx, .txt)을 업로드하면 1~5단계 실시간 파이프라인이 가동됩니다.")
else:
    raw_df_all, sheet_names, sheet_counts, aux_sheets = load_and_sync_to_db(uploaded_file, selected_sheet="ALL")

    if len(sheet_names) > 1:
        sheet_summary = " | ".join(f"**{s}**: {cnt}건" for s, cnt in sheet_counts.items())
        st.info(f"📑 **다중 시트/표 감지 (총 {len(sheet_names)}개 시트, 메인 분석 대상 {len(raw_df_all):,}건):** {sheet_summary}")
        sheet_options = ["ALL"] + sheet_names
        chosen_sheet = st.radio(
            "조회할 시트 범위 선택:",
            options=sheet_options,
            index=0,
            format_func=lambda x: f"📚 메인 시트 통합 조회 (총 {len(raw_df_all):,}건)" if x == "ALL" else f"📄 {x} ({sheet_counts.get(x, 0):,}건)",
            horizontal=True,
        )
        if chosen_sheet == "ALL":
            raw_df = raw_df_all
        else:
            raw_df, _, _, _ = load_and_sync_to_db(uploaded_file, selected_sheet=chosen_sheet)
    else:
        raw_df = raw_df_all

    if aux_sheets:
        with st.expander(f"📎 보조 시트({', '.join(aux_sheets.keys())}) 원문 보기 (메인 집계 건수 오염 방지를 위해 자동 분리됨)", expanded=False):
            for a_name, a_df in aux_sheets.items():
                st.markdown(f"**■ 시트명: `{a_name}`**")
                st.dataframe(style_dataframe_with_colors(a_df), use_container_width=True, hide_index=True)

    # 실시간 파이프라인 1~5번 상태 배지
    render_pipeline_status(
        has_file=True,
        has_query_df=len(raw_df) > 0,
        saved_results=st.session_state.get("saved_results", {}),
    )

    # [W47 신규] 부서별 엑셀/문서 양식 자동 검증(Validation) 배너
    render_validation_banner(raw_df)

    st.subheader("1️⃣ 자연어 기반 데이터 수집 및 SQLite DB 자동 조회")
    tab_nl, tab_kw, tab_sql = st.tabs(["🤖 자연어 → SQL 자동 변환", "🔍 키워드 조합 필터", "🗄️ 수동 SQL 편집기 (.db / .sql 추출)"])

    active_filter_text = ""

    with tab_nl:
        st.caption(
            f"테이블(`{TABLE_NAME}`) 컬럼: `{list(raw_df.columns)}` | "
            "💡 **비워두거나 '전체조회' 입력 시 전체 건수(100% 전체 행)가 조회**되며, `'상위 50건'`, `'Linux만'`, `'빨강 배경만'` 등 조건 입력도 가능합니다."
        )
        nl_question = st.text_input("조회 조건을 자연어로 입력하세요 (비워두면 제한 없이 전체 데이터를 조회합니다):", value="")
        auto_sql = build_smart_sql(nl_question, raw_df, TABLE_NAME)
        st.code(auto_sql, language="sql")
        try:
            df = run_sql_query(auto_sql)
            st.success(f"✅ SQL 자동 실행 완료: 전체 **{len(raw_df):,}건** 중 **{len(df):,}건** 조회됨")
            st.dataframe(style_dataframe_with_colors(df), use_container_width=True, hide_index=True)
            if nl_question.strip():
                active_filter_text = nl_question.strip()
        except Exception as e:
            st.error(f"SQL 실행 오류: {e}")
            df = raw_df

    with tab_kw:
        c1, c2 = st.columns([3, 1])
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,` 구분, 예: `Linux, Android, Encode`)", value="")
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
                st.success(f"쿼리 실행 성공: {len(df):,}건 조회됨")
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

    render_visualization_fragment(target_df, aux_sheets, defect_threshold)
    st.divider()
    render_agent_analysis_section(
        target_df, raw_df, summary_df, selected_models, active_filter_text, aux_sheets, defect_threshold
    )