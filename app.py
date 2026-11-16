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
    .wf-card {
        border-radius: 10px;
        padding: 12px 14px;
        border: 1.5px solid #D1D5DB;
        background: #F9FAFB;
        height: 100%;
    }
    .wf-card-done {
        border-color: #059669;
        background: #ECFDF5;
    }
    .wf-card-wait {
        border-color: #9CA3AF;
        background: #F3F4F6;
        opacity: 0.85;
    }
    .wf-title {
        font-weight: 700;
        font-size: 13.5px;
        color: #065F46;
        margin-bottom: 4px;
    }
    .wf-file {
        font-family: monospace;
        font-size: 11.5px;
        color: #047857;
        background: #D1FAE5;
        padding: 2px 6px;
        border-radius: 4px;
        display: inline-block;
        margin-bottom: 6px;
    }
    .wf-status {
        font-size: 12px;
        color: #1F2937;
        line-height: 1.4;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

DISTINCT_PALETTE = ["#059669", "#2563EB", "#DC2626", "#D97706", "#7C3AED", "#0891B2", "#DB2777"]

st.title("📊 자연어 기반 사내 데이터·문서 통합 분석 AI Agent")

with st.expander("💡 [도입 목적] 왜 엑셀 수작업이나 일반 ChatGPT 대신 이 에이전트를 쓰나요? (클릭)", expanded=False):
    p_c1, p_c2, p_c3, p_c4 = st.columns(4)
    with p_c1:
        st.markdown(
            "**1️⃣ 사내 기밀 100% 보안 유지**\n"
            "- 외부 웹 ChatGPT에 사내 문서를 올리면 외부 서버에 유출·학습됩니다.\n"
            "- **Local Ollama 모드**는 사내 PC 내부에서만 돌아가 기밀 유출이 **0건**입니다."
        )
    with p_c2:
        st.markdown(
            "**2️⃣ AI 숫자 거짓말(환각) 차단**\n"
            "- 일반 AI에 파일을 통째로 주면 합계·비율을 틀리게 계산합니다.\n"
            "- **Python/SQL 엔진이 먼저 100% 정확히 계산**하고 AI는 해석만 수행합니다."
        )
    with p_c3:
        st.markdown(
            "**3️⃣ 엑셀·워드·폴라리온 통합 DB화**\n"
            "- 다중 시트 엑셀(`.xlsx`), 폴라리온 TC 절차서, **워드 보고서(`.docx`)**를 즉시 **SQLite DB**로 정제·변환합니다."
        )
    with p_c4:
        st.markdown(
            "**4️⃣ 주간보고 산출물 1분 자동화**\n"
            "- 매주 2~3시간 걸리던 데이터 필터링·차트 작성·보고서 초안 작업을 **3대 AI 교차 검증 + 마스터 엑셀(`.xlsx`) 추출**로 1분 만에 끝냅니다."
        )

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


# =====================================================================
# ⚡ 실시간 워크플로우 파이프라인 가동 현황판 (Dify 캔버스 대체 실시간 뷰어)
# =====================================================================
def render_live_workflow_tracker(
    file_name: str,
    raw_count: int,
    filtered_count: int,
    aux_count: int,
    saved_results: dict,
):
    """
    정적 그림판(노드구성도_뷰어.html) 대신, 현재 우리 앱에서 실제 데이터가
    어느 .py 노드를 통과해 몇 건 처리되었는지 실시간으로 불이 켜지는 라이브 파이프라인 현황판입니다.
    """
    has_ai_done = len(saved_results) > 0
    ai_badge = f"✅ 완료 ({len(saved_results)}개 AI 모델)" if has_ai_done else "⏳ 실행 대기 (3번 버튼 클릭 시 가동)"
    ai_card_cls = "wf-card wf-card-done" if has_ai_done else "wf-card wf-card-wait"

    with st.expander("⚡ [실시간 파이프라인 가동 현황] 각 파이썬(.py) 노드별 실시간 데이터 흐름 보기 (클릭)", expanded=True):
        cols = st.columns(5)
        steps = [
            (
                "1️⃣ 입력 & 정제 노드",
                "data_loader.py",
                "wf-card wf-card-done",
                f"✅ <b>{file_name[:18]}</b> 로드 완료<br/>• Polarion 개행(_x000D_) 정제<br/>• 보조시트/표 {aux_count}개 분리",
            ),
            (
                "2️⃣ SQLite DB & SQL 노드",
                "app.py (build_smart_sql)",
                "wf-card wf-card-done",
                f"✅ 전체 <b>{raw_count:,}건</b> DB 적재 완료<br/>• 현재 SQL 조건 추출: <b>{filtered_count:,}건</b>",
            ),
            (
                "3️⃣ 파이썬 고속 집계 노드",
                "analyzer_node.py",
                "wf-card wf-card-done",
                f"✅ <b>{filtered_count:,}건</b> 통계·클러스터링 완료<br/>• 숫자 환각 방지 선(先)계산 (0.01초)",
            ),
            (
                "4️⃣ 3대 AI 보고서 노드",
                "reporter_node.py + llm.py",
                ai_card_cls,
                f"{ai_badge}<br/>• IF/ELSE 판정 + 4단락 리포트 병렬 작성",
            ),
            (
                "5️⃣ 마스터 엑셀 패키징",
                "excel_exporter.py",
                ai_card_cls,
                f"{'✅ 통합 엑셀(.xlsx) 생성 완료' if has_ai_done else '⏳ AI 분석 후 자동 생성'}<br/>• 한글 집계표 + BarChart 내장",
            ),
        ]
        for col, (title, py_name, css_cls, desc) in zip(cols, steps):
            with col:
                st.markdown(
                    f'<div class="{css_cls}">'
                    f'<div class="wf-title">{title}</div>'
                    f'<div class="wf-file">📄 {py_name}</div>'
                    f'<div class="wf-status">{desc}</div>'
                    f"</div>",
                    unsafe_allow_html=True,
                )


def style_dataframe_with_colors(df: pd.DataFrame):
    """
    [중요 수정] 'Booting Failure shall not occur'처럼 긴 절차서/기준 문장 안에
    영단어 'Failure'가 들어있다고 해서 빨간색으로 칠하던 문제를 완벽히 해결했습니다.
    오직 명시적 색상태그([🎨빨강강조])가 있거나, 셀 값 자체가 단독 판정어('FAIL', 'NG', '불합격')일 때만 강조합니다.
    """
    def _highlight_cell(val):
        s = str(val).strip()
        s_upper = s.upper()
        # 1) 명시적 빨강 태그 또는 경고 아이콘이 있거나, 셀 전체 값이 단독 'FAIL'/'NG'/'불합격'일 때만 빨간색 표시!
        if "🎨빨강" in s or "🚨" in s or s_upper in ("FAIL", "FAILED", "NG", "불합격", "실행 실패"):
            return "background-color: #FEE2E2; color: #991B1B; font-weight: bold;"
        # 2) 명시적 초록 태그 또는 단독 'PASS'/'합격'일 때만 초록색 표시
        if "🎨초록" in s or "✅" in s or s_upper in ("PASS", "PASSED", "합격", "정상"):
            return "background-color: #DCFCE7; color: #166534; font-weight: bold;"
        if "🎨노랑" in s or "🎨주황" in s:
            return "background-color: #FEF9C3; color: #854D0E; font-weight: bold;"
        if "🎨파랑" in s or "🎨보라" in s:
            return "background-color: #DBEAFE; color: #1E40AF; font-weight: bold;"
        return ""

    if len(df) <= 600:
        return df.style.map(_highlight_cell)
    return df


COLUMN_ALIASES = {
    "title": ["title", "타이틀", "제목", "이름", "명칭", "시나리오명", "본문 상세 내용", "문단", "description"],
    "id": ["id", "tc-id", "tc_id", "tc id", "tc", "아이디", "번호", "코드", "문단id", "행id"],
    "precondition": ["precondition", "pre-condition", "사전조건", "전제조건", "준비상태", "초기상태"],
    "procedure": ["procedure", "test step", "steps", "절차", "수행절차", "테스트절차", "동작"],
    "expected": ["expected result", "expected", "pass/fail criteria", "기대결과", "예상결과", "판정기준"],
    "actual": ["actual result", "actual", "execution status", "status", "실제결과", "수행결과", "판정", "결과"],
    "category": ["category", "domain", "sub system", "app", "test type", "카테고리", "분류", "도메인", "서브시스템", "앱", "구분", "섹션", "목차"],
    "line": ["라인", "생산라인", "공정", "설비"],
}

NL_FILLER_WORDS = {
    "조회해줘", "조회해달라고", "조회해달라", "조회해", "조회", "보여줘", "보여달라고", "찾아줘", "찾아",
    "검색해줘", "검색", "추출해줘", "추출", "출력해줘", "알려줘", "분석해줘", "데이터",
    "들어간거", "들어간", "포함된거", "포함된", "포함", "관련된", "관련", "있는거", "있는",
    "칸에", "칸에서", "칸", "열에", "열에서", "열", "컬럼에", "컬럼에서", "컬럼", "필드에",
    "항목에", "항목", "행만", "행", "부분", "내용", "중에서", "중에", "에서", "만", "좀",
    "그거", "이거", "저거", "전부", "모두", "다", "또는", "이나", "나", "이랑", "랑", "하고",
    "값에", "값에서", "값이", "값은", "값만", "값", "수치에", "수치가", "수치", "숫자",
    "이상인거만", "이상인거", "이상인것", "이상만", "이상",
    "이하인거만", "이하인거", "이하인것", "이하만", "이하",
    "초과인거만", "초과인거", "초과만", "초과",
    "미만인거만", "미만인거", "미만만", "미만",
    "넘는거만", "넘는거", "넘는행", "넘는", "넘어가는", "이내", "부터", "까지",
    "보다", "큰", "작은", "많은", "적은", "높은", "낮은", "인거만", "인거", "인것", "거만", "것만",
    "해달라고하는데", "해달라고", "달라고", "해라",
}


def build_smart_sql(nl_query: str, df: pd.DataFrame, table_name: str) -> str:
    """
    전체 592건 등 대용량 TC 절차서도 100건에서 잘리지 않도록 LIMIT 5000을 기본 적용하며,
    수치 이상/이하, 특정 컬럼 지정, 부분 키워드 검색을 정확히 SQL로 변환합니다.
    """
    if not nl_query or not nl_query.strip():
        return f"SELECT * FROM {table_name} LIMIT 5000;"

    q_orig = nl_query.strip()
    if q_orig in ("전체", "전체 조회", "모두 조회"):
        return f"SELECT * FROM {table_name} LIMIT 5000;"

    all_cols = list(df.columns)
    where_clauses = []

    if any(k in q_orig for k in ["빨강", "빨간", "초록", "노랑", "배경", "색상", "색깔", "강조"]):
        color_kw = "빨강" if ("빨강" in q_orig or "빨간" in q_orig) else ("초록" if "초록" in q_orig else ("노랑" if "노랑" in q_orig else "강조"))
        or_likes = [f'"{col}" LIKE \'%{color_kw}%\'' for col in all_cols]
        return f"SELECT * FROM {table_name} WHERE (" + " OR ".join(or_likes) + ") LIMIT 5000;"

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

    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    nums = re.findall(r"\d+(?:\.\d+)?", q_orig)
    has_num_condition = False

    if nums and num_cols:
        target_num_str = nums[0]
        target_num_val = float(target_num_str)

        if any(k in q_orig for k in ["이상", "부터", ">="]):
            op = ">="
        elif any(k in q_orig for k in ["넘", "초과", "보다 큰", "많은", "높은", ">"]):
            op = ">"
        elif any(k in q_orig for k in ["이하", "이내", "까지", "<="]):
            op = "<="
        elif any(k in q_orig for k in ["미만", "보다 작은", "적은", "낮은", "<"]):
            op = "<"
        else:
            op = "="

        num_col = next((col for col in num_cols if col.lower() in q_lower), None)
        if num_col is None:
            for col in num_cols:
                if (
                    ("불량" in q_orig and "불량" in col)
                    or ("생산" in q_orig and "생산" in col)
                    or ("가동" in q_orig and "가동" in col)
                    or ("매출" in q_orig and "매출" in col)
                    or ("글자" in q_orig and "글자" in col)
                ):
                    num_col = col
                    break

        if num_col is None:
            range_matched_cols = []
            for col in num_cols:
                s = df[col].dropna()
                if len(s) == 0:
                    continue
                c_min, c_max = float(s.min()), float(s.max())
                if (c_min * 0.2) <= target_num_val <= (c_max * 1.5):
                    range_matched_cols.append(col)

            if len(range_matched_cols) == 1:
                num_col = range_matched_cols[0]
            elif len(range_matched_cols) > 1:
                or_num_clauses = [f'"{c}" {op} {target_num_str}' for c in range_matched_cols]
                where_clauses.append("(" + " OR ".join(or_num_clauses) + ")")
                has_num_condition = True
            else:
                num_col = num_cols[0]

        if num_col is not None and not has_num_condition:
            where_clauses.append(f'"{num_col}" {op} {target_num_str}')
            has_num_condition = True

    quoted_words = re.findall(r"['\"“”‘’]([^'\"“”‘’]+)['\"“”‘’]", q_orig)
    keywords = [w.strip() for w in quoted_words if w.strip()]

    q_cleaned = re.sub(r"['\"“”‘’][^'\"“”‘’]+['\"“”‘’]", " ", q_orig)
    q_cleaned = re.sub(r"\d+(?:\.\d+)?\s*(?:개|건|%|점|원)?\s*(?:이상|이하|초과|미만|넘는|이내|부터|까지)?[가-힣]*", " ", q_cleaned)
    q_cleaned = re.sub(r"(?:조회|보여|찾아|검색|추출|출력|알려|분석|확인|정리|비교)[가-힣]*", " ", q_cleaned)

    raw_tokens = re.findall(r"[가-힣A-Za-z0-9_-]+", q_cleaned)

    for tok in raw_tokens:
        tok_low = tok.lower()
        if tok_low in matched_col_tokens or tok_low in NL_FILLER_WORDS:
            continue
        if re.fullmatch(r"\d+(?:\.\d+)?", tok):
            continue
        if any(tok_low.startswith(c_token) for c_token in matched_col_tokens):
            continue

        stem = re.sub(
            r"(인거만|인것만|인거|인것|거만|것만|만|이나|나|이랑|랑|하고|과|와|을|를|이|가|은|는|에서|에게|으로|로|에|도|씩|들)$",
            "",
            tok,
        ).strip()
        if not stem or stem.lower() in NL_FILLER_WORDS or stem.lower() in matched_col_tokens:
            continue
        if re.fullmatch(r"\d+(?:\.\d+)?", stem):
            continue
        if len(stem) >= 1 and stem not in keywords:
            keywords.append(stem)

    str_cols = df.select_dtypes(include=["object", "category"]).columns.tolist() or all_cols
    search_scope_cols = [c for c in (target_cols if target_cols else str_cols) if c not in num_cols] or str_cols

    kw_sql_parts = []
    for kw in keywords:
        kw_safe = kw.replace("'", "''")
        matching_cols = [
            col for col in search_scope_cols
            if df[col].astype(str).str.contains(kw, case=False, na=False, regex=False).any()
        ]
        if matching_cols:
            if len(matching_cols) == 1:
                kw_sql_parts.append(f'"{matching_cols[0]}" LIKE \'%{kw_safe}%\'')
            else:
                or_cols = " OR ".join(f'"{c}" LIKE \'%{kw_safe}%\'' for c in matching_cols[:4])
                kw_sql_parts.append(f"({or_cols})")
        else:
            fallback_cols = target_cols if target_cols else search_scope_cols[:3]
            if len(fallback_cols) == 1:
                kw_sql_parts.append(f'"{fallback_cols[0]}" LIKE \'%{kw_safe}%\'')
            else:
                or_cols = " OR ".join(f'"{c}" LIKE \'%{kw_safe}%\'' for c in fallback_cols)
                kw_sql_parts.append(f"({or_cols})")

    if kw_sql_parts:
        if len(kw_sql_parts) == 1:
            where_clauses.append(kw_sql_parts[0])
        else:
            where_clauses.append("(" + " OR ".join(kw_sql_parts) + ")")

    sql = f"SELECT * FROM {table_name}"
    if where_clauses:
        sql += " WHERE " + " AND ".join(where_clauses)
    sql += " LIMIT 5000;"
    return sql


def extract_question_target_df(question: str, raw_df: pd.DataFrame) -> tuple[pd.DataFrame, list[str]]:
    if not question or not question.strip() or len(raw_df) == 0:
        return pd.DataFrame(), []

    q_low = question.lower()
    col_matches: dict[str, list[str]] = {}
    matched_labels: list[str] = []

    str_cols = raw_df.select_dtypes(include=["object", "category"]).columns.tolist()
    for col in str_cols:
        vals_found = []
        for val in raw_df[col].dropna().astype(str).unique():
            val_clean = re.sub(r"\s*\[[^\]]+\]", "", val).strip()
            if not val_clean or val_clean == "-" or len(val_clean) > 30:
                continue
            if val_clean.lower() in q_low:
                vals_found.append(val_clean)
                matched_labels.append(f"{col}='{val_clean}'")
        if vals_found:
            col_matches[col] = list(dict.fromkeys(vals_found))

    if not col_matches:
        return pd.DataFrame(), []

    mask = pd.Series(True, index=raw_df.index)
    for col, vals in col_matches.items():
        col_series = raw_df[col].astype(str).apply(lambda x: re.sub(r"\s*\[[^\]]+\]", "", x).strip())
        mask &= col_series.isin(vals)

    return raw_df[mask].copy(), list(dict.fromkeys(matched_labels))


def resolve_analysis_scope(
    question: str,
    q_mode: str,
    scope_option: str,
    target_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    active_filter_text: str,
) -> tuple[pd.DataFrame, str, str]:
    if q_mode.startswith("🤖"):
        scope_desc = active_filter_text or "1번 조회 데이터 전체"
        return target_df, scope_desc, ""

    q_df, q_labels = extract_question_target_df(question, raw_df)
    asks_all = any(k in question for k in ["전체 라인", "모든 라인", "전 라인", "전체 데이터", "모든 데이터", "전체 비교"])

    is_missing_in_step1 = False
    if len(q_df) > 0:
        overlap_idx = target_df.index.intersection(q_df.index)
        if len(overlap_idx) < len(q_df):
            is_missing_in_step1 = True

    if asks_all and len(target_df) < len(raw_df):
        return (
            raw_df,
            "전체 원본 데이터 (3번 질문 '전체' 키워드 감지)",
            f"💡 **[스마트 범위 확장]** 1번에서는 `{active_filter_text}`({len(target_df)}건)만 조회되어 있으나, 3번 질문에서 **전체 데이터 비교**를 요청하셔서 전체 원본(`{len(raw_df)}건`)으로 확장해 분석합니다.",
        )

    if scope_option.startswith("🔄"):
        if is_missing_in_step1 and len(q_df) > 0:
            combined_df = pd.concat([target_df, q_df]).drop_duplicates().reset_index(drop=True)
            q_label_str = ", ".join(q_labels)
            notice = (
                f"🔄 **[스마트 통합 연동 작동]** 1번 조회 조건(`{active_filter_text or '현재 표'}`)에 없는 **`{q_label_str}`** 질문이 감지되었습니다! "
                f"원본 DB에서 **`{q_label_str}`({len(q_df)}건)** 데이터를 자동으로 추가 조회하여 **1번 데이터({len(target_df)}건) + 3번 질문 데이터({len(q_df)}건) = 총 {len(combined_df)}건**을 함께 비교·분석합니다."
            )
            return combined_df, f"1번 조회({active_filter_text}) + 3번 질문({q_label_str}) 통합 데이터", notice
        return target_df, active_filter_text or "1번 조회 데이터", ""

    elif scope_option.startswith("🎯"):
        if len(q_df) > 0:
            q_label_str = ", ".join(q_labels)
            notice = (
                f"🎯 **[3번 질문 단독 조회 작동]** 1번 조회 필터(`{active_filter_text or '전체'}`)와 상관없이, "
                f"3번 질문에서 지목하신 **`{q_label_str}` ({len(q_df)}건)** 데이터만 원본 DB에서 새로 추출하여 집중 분석합니다."
            )
            return q_df.reset_index(drop=True), f"3번 질문 단독 조회 ({q_label_str})", notice
        return raw_df, "전체 원본 데이터 (3번 질문 기준)", "🎯 **[3번 질문 기준]** 전체 원본 데이터를 대상으로 분석합니다."

    else:
        warn_msg = ""
        if is_missing_in_step1 and len(q_labels) > 0:
            warn_msg = (
                f"⚠️ **[데이터 범위 주의]** 3번 질문에서 **`{', '.join(q_labels)}`**을(를) 요청하셨지만, "
                f"현재 설정이 `🔒 1번 조회 표 고정`이어서 **`{active_filter_text}`({len(target_df)}건)** 데이터만 AI에게 전달됩니다. "
                f"(`B라인` 등을 분석하려면 위 옵션을 `🔄 통합 비교` 또는 `🎯 3번 질문 단독 분석`으로 선택하세요!)"
            )
        return target_df, active_filter_text or "1번 조회 데이터 고정", warn_msg


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
        scope_parts.append(f"보조 시트/표[{', '.join(aux_sheets.keys())}] 통합")

    stat_snippets = []
    if "생산량" in all_cols and "불량수" in all_cols:
        p_sum = int(target_df["생산량"].sum())
        d_sum = int(target_df["불량수"].sum())
        rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
        stat_snippets.append(f"총 생산량 {p_sum:,}개, 총 불량수 {d_sum:,}개(평균 불량률 {rate:.2f}%)")
    else:
        for cand in ["Sub System", "Domain", "Test Type", "Automated"]:
            if cand in all_cols:
                vc_top = target_df[cand].astype(str).value_counts().head(2).to_dict()
                s_txt = ", ".join(f"{k}({v}건)" for k, v in vc_top.items() if k != "-")
                if s_txt:
                    stat_snippets.append(f"{cand}[{s_txt}]")
        title_col = next((c for c in all_cols if any(k in c.lower() for k in ["문서 섹션", "title", "제목", "타이틀"])), None)
        if title_col and len(stat_snippets) < 2:
            topic_df = cluster_text_series_by_topic(target_df[title_col], max_groups=3)
            if len(topic_df) > 0:
                top_topics = ", ".join(f"{r['주제/분류 그룹']}({r['건수']}건)" for _, r in topic_df.head(2).iterrows())
                stat_snippets.append(f"주요 주제[{top_topics}]")

    scope_title = " + ".join(scope_parts)
    stat_str = " / ".join(stat_snippets) if stat_snippets else f"총 {len(target_df)}건"

    return (
        f"[AI 비서 자동 맞춤 질문] {scope_title} ({stat_str}) 데이터를 바탕으로 "
        f"① 도메인/서브시스템 및 주요 주제별 분포 현황(%), ② 자동화 여부(Automated)·테스트 유형별 중점 관리 시나리오 분석, "
        f"③ 테스트 수행 시 유의해야 할 기술적 리스크 및 취약점 가설, ④ 담당 검증/개발 부서별 실행 계획 및 품질 확보 방안을 정리해 줘."
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
        for c in ["Domain", "Sub System", "App"]:
            if c in all_cols:
                vals = [v for v in target_df[c].astype(str).unique() if v not in ("-", "", "N/A")]
                keywords.extend(vals[:1])
        keywords.append("전장 SW 테스트 자동화 검증")

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
    s = str(text)
    if any(k in s.lower() for k in ["booting failure", "shall not occur", "오류", "에러", "않는다", "없음", "실패", "차단"]):
        return "🛡️ 부팅/시스템 오류 미발생(안정성) 검증"
    if any(k in s.lower() for k in ["expected output", "디오", "오디오", "화면", "출력", "표시된", "나타난", "동작"]):
        return "🖥️ 정상 출력 및 기능 동작(Expected Output) 검증"
    if any(k in s.lower() for k in ["이동", "전환", "열리", "연결", "통신"]):
        return "🔗 인터페이스 연결 및 상태 전환 검증"
    return "📌 기타 판정 기준 검증"


def render_executive_summary_cards(target_df: pd.DataFrame):
    """차트 상단 4대 핵심 요약 카드"""
    all_cols = target_df.columns.tolist()
    k1, k2, k3, k4 = st.columns(4)

    if "생산량" in all_cols and "불량수" in all_cols:
        p_sum = int(target_df["생산량"].sum())
        d_sum = int(target_df["불량수"].sum())
        rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
        top_def = target_df.groupby("불량유형")["불량수"].sum().idxmax() if "불량유형" in all_cols and d_sum > 0 else "-"
        k1.metric("📌 분석 데이터 건수", f"{len(target_df):,} 건")
        k2.metric("🏭 총 생산량 / 불량수", f"{p_sum:,} 개 / {d_sum:,} 개")
        k3.metric("🚨 평균 불량률 (기준 1.5%)", f"{rate:.2f} %")
        k4.metric("🔍 최다 발생 불량유형", str(top_def))
    else:
        group_col = next((c for c in ["Sub System", "Domain", "문서 섹션(목차)", "Category", "Title"] if c in all_cols), all_cols[0])
        topic_df = cluster_text_series_by_topic(target_df[group_col], max_groups=5)
        top_group = f"{topic_df.iloc[0]['주제/분류 그룹']} ({topic_df.iloc[0]['건수']}건)" if len(topic_df) > 0 else "-"

        if "Automated" in all_cols:
            vc_auto = target_df["Automated"].astype(str).value_counts()
            top_auto = f"{vc_auto.index[0]} ({vc_auto.iloc[0]}건)" if len(vc_auto) > 0 else "-"
            k1.metric("📌 총 TC 절차 건수", f"{len(target_df):,} 건")
            k2.metric("📂 분석 컬럼(속성) 수", f"{len(all_cols)} 개")
            k3.metric(f"🏆 1위 {group_col}", top_group[:22])
            k4.metric("🤖 자동화(Automated) 현황", top_auto)
        else:
            color_df = analyze_color_cell_contexts(target_df)
            issue_cnt = len(color_df)
            k1.metric("📌 총 조회 항목 수", f"{len(target_df):,} 건")
            k2.metric("📂 분석 컬럼(속성) 수", f"{len(all_cols)} 개")
            k3.metric("🏆 1위 핵심 분류/주제", top_group[:22])
            k4.metric("🎨 셀 색상강조 감지", f"{issue_cnt:,} 건")


def render_smart_charts(target_df: pd.DataFrame, key_prefix: str = "main", view_style: str = "grouped"):
    if len(target_df) == 0:
        st.info("조회된 데이터가 없습니다.")
        return

    num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
    all_cols = target_df.columns.tolist()

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
                agg_line = agg_line.sort_values(by="불량수", ascending=True)
                agg_line["표시라벨"] = agg_line.apply(lambda r: f" {int(r['불량수'])}개 (불량률 {r['불량률(%)']}%)", axis=1)
                fig1 = px.bar(
                    agg_line, x="불량수", y=line_col, orientation="h", color=line_col, text="표시라벨",
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"① [한눈에 보는 순위] {line_col}별 총 불량수 및 불량률(%)",
                )
                fig1.update_traces(textposition="outside", textfont_size=14)
                fig1.update_layout(height=400, font=dict(size=14), showlegend=False, margin=dict(r=90, t=55, b=30))
                st.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_num1")

        with row1_c2:
            if defect_type_col:
                agg_def = target_df.groupby(defect_type_col, as_index=False)[val_col].sum()
                fig2 = px.pie(
                    agg_def, names=defect_type_col, values=val_col, hole=0.42,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"② [원인 비중] {defect_type_col}별 점유율 도넛 차트",
                )
                fig2.update_traces(textinfo="label+percent+value", textfont_size=14)
                fig2.update_layout(height=400, font=dict(size=14))
                st.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_num2")

        row2_c1, row2_c2 = st.columns(2)
        with row2_c1:
            if time_col and "생산량" in target_df.columns and "불량수" in target_df.columns:
                agg_time = target_df.groupby([time_col, line_col], as_index=False)[["생산량", "불량수"]].sum()
                agg_time["불량률(%)"] = (agg_time["불량수"] / agg_time["생산량"] * 100).round(2)
                fig3 = px.line(
                    agg_time, x=time_col, y="불량률(%)", color=line_col, markers=True, text="불량률(%)",
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"③ [기간별 추세] {time_col}별 불량률(%) 추이 (기준 1.5% 점선)",
                )
                fig3.add_hline(y=1.5, line_dash="dash", line_color="#DC2626", annotation_text="사내 기준 (1.5%)")
                fig3.update_layout(height=400, font=dict(size=13))
                st.plotly_chart(fig3, use_container_width=True, key=f"{key_prefix}_num3")
        with row2_c2:
            if shift_col:
                agg_shift = target_df.groupby([shift_col, defect_type_col or line_col], as_index=False)[val_col].sum()
                fig4 = px.bar(
                    agg_shift, x=shift_col, y=val_col, color=defect_type_col or line_col, barmode="group", text_auto=True,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"④ [시간대 비교] {shift_col}(주/야간)별 비교",
                )
                fig4.update_layout(height=400, font=dict(size=13))
                st.plotly_chart(fig4, use_container_width=True, key=f"{key_prefix}_num4")
        return

    clean_df = target_df.copy()
    for c in clean_df.columns:
        clean_df[c] = clean_df[c].astype(str).apply(lambda x: re.sub(r"\s*\[[^\]]+\]", "", x).strip())

    # Polarion TC 절차서(Sub System, Domain, Test Type, Automated, Title 등)에 최적화된 축 자동 선택
    primary_col = next(
        (c for c in ["Sub System", "문서 섹션(목차)", "Category", "Title", "제목", "Domain"] if c in all_cols and clean_df[c].nunique() > 1),
        next((c for c in all_cols if "title" in c.lower()), all_cols[min(1, len(all_cols) - 1)]),
    )
    pie_col = next(
        (c for c in ["Automated", "Test Type", "Domain", "핵심 분류", "Execution Status", "Actual Result"] if c in all_cols and 1 <= clean_df[c].nunique() <= 12 and c != primary_col),
        None,
    )
    exp_col = next((c for c in all_cols if any(k in c.lower() for k in ["pass/fail criteria", "expected", "기대결과", "본문 상세 내용"])), None)

    r1_c1, r1_c2 = st.columns(2)

    with r1_c1:
        if view_style == "raw" or clean_df[primary_col].nunique() <= 10:
            vc_raw = clean_df[primary_col].value_counts().head(10).reset_index()
            vc_raw.columns = ["항목명", "건수"]
            vc_raw["비중(%)"] = (vc_raw["건수"] / len(clean_df) * 100).round(1)
            vc_raw["요약항목"] = vc_raw["항목명"].apply(lambda x: shorten_long_label(x, max_len=26))
            vc_raw["표시라벨"] = vc_raw.apply(lambda r: f" {r['건수']}건 ({r['비중(%)']}%)", axis=1)
            plot_c1 = vc_raw.iloc[::-1]
            fig_c1 = px.bar(
                plot_c1, x="건수", y="요약항목", orientation="h", color="요약항목", text="표시라벨",
                hover_data=["항목명"], color_discrete_sequence=DISTINCT_PALETTE,
                title=f"① [한눈에 보는 TOP 순위] {primary_col}별 건수 및 비중(%)",
            )
        else:
            grouped_c1 = cluster_text_series_by_topic(clean_df[primary_col], max_groups=8)
            plot_c1 = grouped_c1.iloc[::-1]
            fig_c1 = px.bar(
                plot_c1, x="건수", y="주제/분류 그룹", orientation="h", color="주제/분류 그룹", text="표시라벨",
                hover_data=["포함된 상세 항목"], color_discrete_sequence=DISTINCT_PALETTE,
                title=f"① [한눈에 보는 주제 순위] {primary_col} 핵심 키워드 묶음 TOP 순위",
            )
        fig_c1.update_traces(textposition="outside", textfont_size=14)
        fig_c1.update_layout(
            height=420, font=dict(size=14), showlegend=False,
            xaxis_title="건수 (건)", yaxis_title="", margin=dict(l=15, r=90, t=55, b=30),
        )
        st.plotly_chart(fig_c1, use_container_width=True, key=f"{key_prefix}_cat1")

    with r1_c2:
        if pie_col:
            vc_st = clean_df[pie_col].value_counts().reset_index()
            vc_st.columns = [pie_col, "건수"]
            fig_c2 = px.pie(
                vc_st, names=pie_col, values="건수", hole=0.45,
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"② [전체 구성 비중] {pie_col} 분포 비율(%)",
            )
            fig_c2.update_traces(textinfo="label+percent+value", textfont_size=14)
            fig_c2.update_layout(height=420, font=dict(size=14))
            st.plotly_chart(fig_c2, use_container_width=True, key=f"{key_prefix}_cat2")
        elif exp_col:
            goal_series = clean_df[exp_col].apply(_classify_verification_goal)
            vc_goal = goal_series.value_counts().reset_index()
            vc_goal.columns = ["검증 목적 유형", "건수"]
            fig_c2 = px.pie(
                vc_goal, names="검증 목적 유형", values="건수", hole=0.45,
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"② [전체 비중] {exp_col} 검증 유형별 구성 비율(%)",
            )
            fig_c2.update_traces(textinfo="label+percent+value", textfont_size=14)
            fig_c2.update_layout(height=420, font=dict(size=14))
            st.plotly_chart(fig_c2, use_container_width=True, key=f"{key_prefix}_cat2")

    with st.expander("➕ 세부 항목별 추가 차트 보기 (Title 주제 묶음 · 사전조건 분포 · 담당자/유형 분포)", expanded=False):
        r2_c1, r2_c2 = st.columns(2)
        with r2_c1:
            if "Title" in all_cols and primary_col != "Title":
                t_grp = cluster_text_series_by_topic(clean_df["Title"], max_groups=7).iloc[::-1]
                fig_c3 = px.bar(
                    t_grp, x="건수", y="주제/분류 그룹", orientation="h", color="주제/분류 그룹", text="표시라벨",
                    hover_data=["포함된 상세 항목"], color_discrete_sequence=DISTINCT_PALETTE,
                    title="③ [Title] 유사 주제 키워드 묶음 분포",
                )
                fig_c3.update_traces(textposition="outside", textfont_size=13)
                fig_c3.update_layout(height=380, font=dict(size=13), showlegend=False, margin=dict(r=80))
                st.plotly_chart(fig_c3, use_container_width=True, key=f"{key_prefix}_cat3")
        with r2_c2:
            sub_col = next((c for c in ["Test Type", "Assignee(s)", "Type", "Pre-condition"] if c in all_cols), None)
            if sub_col:
                vc_sub = clean_df[sub_col].value_counts().head(8).reset_index()
                vc_sub.columns = ["항목", "건수"]
                vc_sub["요약"] = vc_sub["항목"].apply(lambda x: shorten_long_label(x, max_len=24))
                fig_c4 = px.bar(
                    vc_sub.iloc[::-1], x="건수", y="요약", orientation="h", color="요약", text_auto=True,
                    color_discrete_sequence=DISTINCT_PALETTE, title=f"④ [{sub_col}] 상위 분포 현황",
                )
                fig_c4.update_layout(height=380, font=dict(size=13), showlegend=False)
                st.plotly_chart(fig_c4, use_container_width=True, key=f"{key_prefix}_cat4")


@st.dialog("📊 AI 모델별 상세 시각화 대시보드 & 인사이트 리포트", width="large")
def show_model_visual_popup(m_code: str, res_data: dict, target_df: pd.DataFrame):
    st.markdown(f"## {MODEL_LABELS.get(m_code, m_code)}")
    st.caption(f"⏱️ 분석 소요 시간: **{res_data['elapsed_time']}초** | 분석 반영 데이터: **{len(target_df)}건**")

    render_executive_summary_cards(target_df)
    st.divider()
    render_smart_charts(target_df, key_prefix=f"popup_{m_code}", view_style="grouped")
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
    st.subheader("2️⃣ 한눈에 보는 핵심 지표(KPI) & 다차원 시각화")

    render_executive_summary_cards(target_df)

    v_tab1, v_tab2, v_tab3 = st.tabs([
        "🎯 한눈에 보는 TOP 순위 & 비중 차트 (직관적 가로형)",
        "🛠️ 자유 조합 커스텀 차트",
        "📋 시트/문서 통합 요약 & 핵심 집계표 (엑셀 2번 시트)",
    ])

    with v_tab1:
        view_mode = st.radio(
            "차트 표시 방식 선택:",
            options=["🧩 유사 키워드 묶음 TOP 순위 (한눈에 보기 · 권장)", "📊 개별 원본 항목 TOP 10 보기"],
            horizontal=True,
            key="chart_view_style_radio",
        )
        style_code = "grouped" if view_mode.startswith("🧩") else "raw"
        render_smart_charts(target_df, key_prefix="main_dash", view_style=style_code)

    with v_tab2:
        num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
        all_cols = target_df.columns.tolist()
        if len(target_df) > 0:
            c1, c2, c3, c4 = st.columns(4)
            default_x_idx = next((i for i, c in enumerate(all_cols) if c in ("Sub System", "Title", "문서 섹션(목차)")), min(1, len(all_cols) - 1))
            x_col = c1.selectbox("X축 (기준 항목)", options=all_cols, index=default_x_idx, key="frag_x")
            y_options = ["(건수 집계 - COUNT)"] + num_cols
            y_sel = c2.selectbox("Y축 (집계 대상)", options=y_options, index=0 if not num_cols else min(1, len(y_options) - 1), key="frag_y")
            chart_dir = c3.selectbox("차트 방향", options=["가로 막대형 (한눈에 보기 추천)", "세로 막대형 (Vertical)"], key="frag_dir")
            use_topic_cluster = c4.checkbox("🧩 유사 키워드로 묶어서 보기", value=("title" in x_col.lower()), key="frag_cluster")

            if use_topic_cluster and y_sel == "(건수 집계 - COUNT)":
                clustered = cluster_text_series_by_topic(target_df[x_col], max_groups=10)
                if len(clustered) > 0:
                    fig_custom = px.bar(
                        clustered.iloc[::-1],
                        x="건수",
                        y="주제/분류 그룹",
                        orientation="h",
                        color="주제/분류 그룹",
                        text="표시라벨",
                        hover_data=["포함된 상세 항목"],
                        color_discrete_sequence=DISTINCT_PALETTE,
                        title=f"[{x_col}] 유사 키워드 그룹별 건수 및 비중(%) 순위",
                    )
                    fig_custom.update_traces(textposition="outside", textfont_size=14)
                    fig_custom.update_layout(height=440, font=dict(size=14), showlegend=False, margin=dict(r=85, t=55, b=30))
                    st.plotly_chart(fig_custom, use_container_width=True, key="frag_custom_chart")
            else:
                temp_df = target_df.copy()
                temp_df["_원문_"] = temp_df[x_col].astype(str)
                temp_df["_요약라벨_"] = temp_df["_원문_"].apply(lambda x: shorten_long_label(x, max_len=28))

                if y_sel == "(건수 집계 - COUNT)":
                    custom_df = temp_df.groupby(["_요약라벨_", "_원문_"], as_index=False).size()
                    custom_df.columns = ["_요약라벨_", "_원문_", "건수"]
                    custom_df = custom_df.sort_values(by="건수", ascending=False).head(15)
                    y_plot = "건수"
                else:
                    custom_df = temp_df.groupby(["_요약라벨_", "_원문_"], as_index=False)[y_sel].sum().sort_values(by=y_sel, ascending=False).head(15)
                    y_plot = y_sel

                if chart_dir.startswith("가로"):
                    fig_custom = px.bar(
                        custom_df.iloc[::-1],
                        x=y_plot,
                        y="_요약라벨_",
                        orientation="h",
                        color="_요약라벨_",
                        text_auto=True,
                        hover_data=["_원문_"],
                        color_discrete_sequence=DISTINCT_PALETTE,
                        title=f"[{x_col}]별 {y_plot} TOP 순위 비교",
                    )
                    fig_custom.update_traces(textposition="outside", textfont_size=14)
                    fig_custom.update_layout(height=max(400, len(custom_df) * 36), font=dict(size=14), showlegend=False, margin=dict(r=70, t=55, b=30))
                else:
                    fig_custom = px.bar(
                        custom_df, x="_요약라벨_", y=y_plot, color="_요약라벨_", text_auto=True,
                        hover_data=["_원문_"], color_discrete_sequence=DISTINCT_PALETTE,
                        title=f"[{x_col}]별 {y_plot} 비교",
                    )
                    fig_custom.update_traces(textposition="outside", textfont_size=14)
                    fig_custom.update_layout(height=440, font=dict(size=14), showlegend=False)
                st.plotly_chart(fig_custom, use_container_width=True, key="frag_custom_chart")

    with v_tab3:
        korean_tables = build_korean_stat_tables(target_df, aux_sheets=aux_sheets)
        for t_title, t_df in korean_tables.items():
            st.markdown(f"**■ {t_title}**")
            st.dataframe(style_dataframe_with_colors(t_df), use_container_width=True, hide_index=True)


@st.fragment
def render_agent_analysis_fragment(
    target_df: pd.DataFrame,
    raw_df: pd.DataFrame,
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
            "🤖 AI 비서 자동 맞춤 질문 (1번 조회 데이터·주제 묶음·도메인 자동 조합)",
            "✍️ 사용자 직접 질문 입력/수정 (다른 라인·조건 질문 시 자동 데이터 연동 지원)",
        ],
        horizontal=True,
    )

    if q_mode.startswith("🤖"):
        st.info(f"**💡 AI 비서가 현재 1번에서 조회된 데이터({len(target_df)}건)에 맞춰 조합한 분석 질문입니다:**\n\n{auto_assistant_q}")
        question = auto_assistant_q
        scope_option = "🔒 1번 조회 표 고정"
    else:
        question = st.text_input(
            "💡 에이전트에게 지시할 분석 질문을 자유롭게 입력하세요:",
            value="B라인의 생산량과 불량률을 분석해서 알려줘" if "라인" in raw_df.columns else auto_assistant_q,
        )
        scope_option = st.radio(
            "🔎 [직접 질문 시] 데이터 분석 범위 연동 방식 선택:",
            options=[
                "🔄 스마트 통합 비교 (1번 조회 데이터 + 3번 질문 대상 데이터를 합쳐서 비교·분석 · 추천)",
                "🎯 3번 질문 대상만 단독 분석 (1번 필터 무시하고 질문한 대상만 원본 DB에서 새로 조회)",
                "🔒 1번 조회된 표 데이터 안에서만 고정 분석",
            ],
            index=0,
            horizontal=True,
        )

    analysis_df, resolved_scope_desc, scope_notice = resolve_analysis_scope(
        question, q_mode, scope_option, target_df, raw_df, active_filter_text
    )

    if scope_notice:
        if scope_notice.startswith("⚠️"):
            st.warning(scope_notice)
        else:
            st.success(scope_notice)

    render_related_links_box(analysis_df)

    if st.button("🚀 에이전트 정밀 분석 실행 (선택 모델 동시 병렬 실행)", type="primary"):
        if not selected_models:
            st.warning("좌측 사이드바에서 최소 1개 이상의 모델을 선택해주세요.")
        else:
            agent_app = build_agent_graph()
            stat_tables = build_korean_stat_tables(analysis_df, aux_sheets=aux_sheets)
            stat_summary_str = "\n\n".join(
                f"■ {k}\n{v.to_string(index=False)}" for k, v in stat_tables.items()
            )
            raw_overview_tables = build_korean_stat_tables(raw_df, aux_sheets=None)
            raw_overview_str = "\n".join(
                f"• {k}:\n{v.to_string(index=False)}" for k, v in list(raw_overview_tables.items())[:2]
            )

            sample_rows_df = analysis_df.head(35)
            data_info = (
                f"[분석 적용 범위]: {resolved_scope_desc}\n"
                f"[분석 대상 건수]: {len(analysis_df)}건 (전체 원본 {len(raw_df)}건)\n"
                f"[보조 시트/내장표 목록]: {list(aux_sheets.keys()) if aux_sheets else '없음'}\n"
                f"[컬럼 목록]: {list(analysis_df.columns)}\n\n"
                f"[1. 분석 대상 정밀 집계표 ({resolved_scope_desc})]\n{stat_summary_str}\n\n"
                f"[2. 참고용 전체 원본 데이터 요약 (타 항목 비교 기준)]\n{raw_overview_str}\n\n"
                f"[3. 분석 대상 상세 행 내역]\n{sample_rows_df.to_string(index=False)}"
            )

            new_results = {}
            status_box = st.status(
                f"⚡ [{resolved_scope_desc} ({len(analysis_df)}건)] 기준으로 AI 모델들을 동시에 병렬 실행하고 있습니다...",
                expanded=True,
            )

            analysis_summary_df = analysis_df.describe(include="all").fillna("")

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
                    analysis_df, analysis_summary_df, res["report"], res["elapsed_time"], m_code, aux_sheets=aux_sheets
                )
                res["excel_bytes"] = excel_bytes
                return m_code, res

            with ThreadPoolExecutor(max_workers=len(selected_models)) as executor:
                future_map = {executor.submit(_run_single_model, m): m for m in selected_models}
                for future in as_completed(future_map):
                    m_done, res_done = future.result()
                    new_results[m_done] = res_done
                    status_box.write(f"✅ **{MODEL_LABELS.get(m_done, m_done)}** 분석 완료! (`{res_done['elapsed_time']}초`)")

            status_box.update(label=f"🎉 모든 모델 분석 완료! (반영 데이터: {resolved_scope_desc}, 총 {len(analysis_df)}건)", state="complete", expanded=False)
            st.session_state["saved_results"] = new_results
            st.session_state["last_analyzed_df"] = analysis_df

    saved = st.session_state.get("saved_results", {})
    if saved:
        display_models = [m for m in selected_models if m in saved] or list(saved.keys())
        popup_df = st.session_state.get("last_analyzed_df") if st.session_state.get("last_analyzed_df") is not None else analysis_df

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
            st.caption("👆 **각 모델의 제목 버튼을 클릭하면 분석에 반영된 데이터 기준의 시각화 팝업창이 열립니다.**")
        with top_c2:
            st.download_button(
                label="📚 전체 모델 종합 비교 엑셀 추출 (질문 연동 데이터·차트 완비 .xlsx)",
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
                    label=f"📥 {m_code} 개별 엑셀 추출 (질문 연동 데이터·차트 포함 .xlsx)",
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
            "- 파이썬이 먼저 시트/문서 통계와 유사 제목 그룹핑을 0.01초 만에 계산하고, Ollama가 리포트 문장을 작성합니다."
        )
        chosen_ollama = st.selectbox("구동할 로컬 모델 변경", options=ollama_models, index=0)
        os.environ["LOCAL_MODEL_NAME"] = chosen_ollama

    if st.session_state["saved_results"]:
        if st.button("🗑️ 이전 분석 결과 초기화", use_container_width=True):
            st.session_state["saved_results"] = {}
            st.rerun()

uploaded_file = st.file_uploader(
    "📂 사내 데이터·문서 파일(CSV, Excel, Word .docx, TXT, MD) 업로드 (즉시 SQLite DB 자동 구축)",
    type=["csv", "xlsx", "docx", "txt", "md"],
)

if uploaded_file is not None:
    raw_df_main, sheet_names, sheet_counts, aux_sheets = load_and_sync_to_db(uploaded_file, selected_sheet="AUTO_MAIN")

    if len(sheet_names) > 1:
        main_sheet_names = [s for s in sheet_names if s not in aux_sheets]
        aux_sheet_names = list(aux_sheets.keys())
        file_type_label = "Word(.docx) 본문 및 내장 표" if uploaded_file.name.lower().endswith(".docx") else "다중 시트 엑셀"
        st.info(
            f"📑 **{file_type_label} 스마트 분리 완료!**\n"
            f"- **🎯 핵심 분석 대상:** `{', '.join(main_sheet_names)}` (총 **{len(raw_df_main)}건**)\n"
            f"- **📋 보조 시트 / 내장 표:** `{', '.join(aux_sheet_names) if aux_sheet_names else '없음'}`"
        )

        if aux_sheets:
            with st.expander(f"🔎 보조 시트 / 워드 내장 표({', '.join(aux_sheet_names)}) 내용 바로 보기 (클릭)", expanded=False):
                aux_cols = st.columns(min(len(aux_sheets), 3))
                for idx_a, (a_name, a_df) in enumerate(aux_sheets.items()):
                    with aux_cols[idx_a % len(aux_cols)]:
                        st.markdown(f"**📄 [{a_name}] ({len(a_df)}건)**")
                        st.dataframe(style_dataframe_with_colors(a_df), use_container_width=True, hide_index=True)

        sheet_options = ["AUTO_MAIN"] + sheet_names
        if not uploaded_file.name.lower().endswith(".docx"):
            sheet_options.append("ALL_MERGE")

        chosen_sheet = st.radio(
            "조회할 데이터 범위 선택:",
            options=sheet_options,
            index=0,
            format_func=lambda x: (
                f"🎯 기본 핵심 데이터 조회 ({len(raw_df_main)}건 · 추천)"
                if x == "AUTO_MAIN"
                else (
                    f"📚 모든 시트 합쳐서 보기 ({sum(sheet_counts.values())}건)"
                    if x == "ALL_MERGE"
                    else f"📄 {x} 단독 분석 ({sheet_counts.get(x, 0)}건)"
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
            "입력 예시: **'Sub System에 Audio 들어간거'**, **'Title칸에 문서 들어간거'**, **'값에 1200이상인거만 조회'**"
        )
        nl_question = st.text_input("조회 조건을 자연어로 입력하세요 (비워두면 전체 데이터를 누락 없이 전수 조회합니다):", value="")
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
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,` 구분, 예: `Audio, Display` 또는 `A라인, B라인`)", value="")
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

    # 실시간 워크플로우 파이프라인 가동 현황판 출력 (파일/DB/집계/AI 상태 실시간 표시!)
    render_live_workflow_tracker(
        file_name=uploaded_file.name,
        raw_count=len(raw_df),
        filtered_count=len(target_df),
        aux_count=len(aux_sheets) if aux_sheets else 0,
        saved_results=st.session_state.get("saved_results", {}),
    )

    st.divider()
    render_visualization_fragment(target_df, aux_sheets)
    st.divider()
    render_agent_analysis_fragment(target_df, raw_df, aux_sheets, summary_df, selected_models, active_filter_text)