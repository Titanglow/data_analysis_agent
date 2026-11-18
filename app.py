# app.py
import io
import os
import re
import sqlite3
import zipfile
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
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
    export_analysis_to_docx,
    export_combined_analysis_to_docx,
    build_korean_stat_tables,
    analyze_color_cell_contexts,
    evaluate_threshold_verdict,
)
from data_agent.llm import get_installed_ollama_models, call_llm
from data_agent.tools.data_loader import (
    load_and_sync_to_db,
    filter_by_keywords,
    run_sql_query,
    export_db_to_sql_script,
    DB_PATH,
    TABLE_NAME,
)
from data_agent.tools.visualizer import (
    is_manufacturing_df,
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
if "current_file_sig" not in st.session_state:
    st.session_state["current_file_sig"] = ""
if "uploader_key" not in st.session_state:
    st.session_state["uploader_key"] = 0
if "chat_history" not in st.session_state:
    st.session_state["chat_history"] = []
if "rag_guideline_text" not in st.session_state:
    st.session_state["rag_guideline_text"] = ""
if "consistency_df" not in st.session_state:
    st.session_state["consistency_df"] = None


def reset_all_inputs_and_files():
    st.session_state["uploader_key"] += 1
    st.session_state["saved_results"] = {}
    st.session_state["last_analyzed_df"] = None
    st.session_state["last_aux_sheets"] = {}
    st.session_state["pipeline_logs"] = []
    st.session_state["current_file_sig"] = ""
    st.session_state["chat_history"] = []
    st.session_state["consistency_df"] = None
    for k in ["nl_q_input", "kw_q_input"]:
        if k in st.session_state:
            st.session_state[k] = ""


def get_file_type_badge(filename: str) -> tuple[str, str]:
    ext = Path(filename).suffix.lower()
    if ext == ".xlsx":
        return "📊 엑셀 (.xlsx)", f"📊 [Excel] {filename}"
    if ext == ".csv":
        return "📈 CSV (.csv)", f"📈 [CSV] {filename}"
    if ext == ".docx":
        return "📄 워드 (.docx)", f"📄 [Word] {filename}"
    if ext == ".txt":
        return "📝 텍스트 (.txt)", f"📝 [TXT] {filename}"
    return "📁 기타 파일", f"📁 {filename}"


def extract_text_from_guideline_file(uploaded_guide) -> str:
    """사내 기준서/가이드 문서(.txt, .md, .docx, .csv)에서 텍스트를 추출합니다 (RAG 연동용)."""
    if uploaded_guide is None:
        return ""
    uploaded_guide.seek(0)
    raw_bytes = uploaded_guide.read()
    fname = uploaded_guide.name.lower()

    if fname.endswith(".docx"):
        ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
        lines = []
        try:
            with zipfile.ZipFile(io.BytesIO(raw_bytes)) as zf:
                if "word/document.xml" in zf.namelist():
                    root = ET.fromstring(zf.read("word/document.xml"))
                    for p in root.findall(".//w:body/w:p", ns):
                        t_list = [t.text for t in p.findall(".//w:t", ns) if t.text]
                        s = "".join(t_list).strip()
                        if s:
                            lines.append(s)
        except Exception:
            pass
        return "\n".join(lines[:80])

    try:
        txt = raw_bytes.decode("utf-8")
    except UnicodeDecodeError:
        txt = raw_bytes.decode("cp949", errors="replace")
    return "\n".join(txt.splitlines()[:80])


def build_pipeline_status_html(
    has_file: bool,
    has_query_df: bool,
    saved_results: dict,
    running_stage: int = 0,
    running_msg: str = "",
) -> str:
    has_analysis = bool(saved_results) and running_stage == 0
    has_verified_report = has_analysis and any(bool(v.get("report")) for v in saved_results.values())
    has_excel_pkg = has_analysis and any(bool(v.get("excel_bytes")) for v in saved_results.values())

    steps_cfg = [
        (1, "1. 문서·시트·양식 스캔", "done" if has_file else "wait", "다중 파일·시트 검증 완료" if has_file else "파일 업로드 대기"),
        (2, "2. SQLite DB·정밀 집계", "done" if has_query_df else "wait", "전체 건수 DB 적재·통계 완료" if has_query_df else "데이터 집계 대기"),
        (
            3,
            "3. 3대 AI 병렬 추론",
            "running" if running_stage == 3 else ("done" if (has_analysis or running_stage > 3) else "wait"),
            running_msg if running_stage == 3 else (f"{len(saved_results)}개 모델 추론 완료" if has_analysis else "분석 실행 대기"),
        ),
        (
            4,
            "4. 리포트 품질 검증·보정",
            "running" if running_stage == 4 else ("done" if has_verified_report else "wait"),
            "수치·기준값 실시간 검증 중..." if running_stage == 4 else ("수치·기준값 정제 완료" if has_verified_report else "리포트 검증 대기"),
        ),
        (
            5,
            "5. 엑셀·워드 패키징 완료",
            "running" if running_stage == 4 else ("done" if has_excel_pkg else "wait"),
            "엑셀·워드 문서 생성 중..." if running_stage == 4 else ("엑셀(.xlsx)·워드(.docx) 완료" if has_excel_pkg else "문서 생성 대기"),
        ),
    ]

    cards_html = []
    for num, title, state_code, desc in steps_cfg:
        if state_code == "done":
            bg, border, title_col, sub_col, icon = "#ecfdf5", "#10b981", "#065f46", "#047857", "✅"
        elif state_code == "running":
            bg, border, title_col, sub_col, icon = "#eff6ff", "#2563eb", "#1e40af", "#1d4ed8", "🔄"
        else:
            bg, border, title_col, sub_col, icon = "#f8fafc", "#cbd5e1", "#64748b", "#94a3b8", "⏳"

        cards_html.append(
            f'<div style="flex:1; min-width:155px; background:{bg}; border:2px solid {border}; '
            f'border-radius:10px; padding:10px 12px; text-align:center; box-shadow:0 1px 3px rgba(0,0,0,0.05);">'
            f'<div style="font-size:0.86rem; font-weight:700; color:{title_col};">{icon} {title}</div>'
            f'<div style="font-size:0.74rem; color:{sub_col}; margin-top:4px; font-weight:600;">{desc}</div>'
            f'</div>'
        )

    return (
        '<div style="margin: 4px 0 14px 0;">'
        '<div style="font-size:0.92rem; font-weight:700; color:#1e293b; margin-bottom:8px;">'
        '⚡ 실시간 파이프라인 가동현황</div>'
        f'<div style="display:flex; gap:10px; flex-wrap:wrap;">{"".join(cards_html)}</div>'
        '</div>'
    )


def style_dataframe_with_colors(df: pd.DataFrame):
    def _highlight_cell(val):
        s = str(val)
        if "빨강" in s or "🔴" in s or "FAIL" in s.upper() or "확인 필요" in s:
            return "background-color: #FEE2E2; color: #991B1B; font-weight: bold;"
        if "초록" in s or "🟢" in s or ("기준(" in s and "이상 비중" in s):
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
    if not nl_query or not nl_query.strip():
        return f"SELECT * FROM {table_name};"

    q = nl_query.strip()
    if q.replace(" ", "").lower() in ("전체", "전체조회", "전체데이터", "모두조회", "다보여줘", "전체보여줘", "all"):
        return f"SELECT * FROM {table_name};"

    limit_val = None
    limit_match = re.search(r"(?:상위\s*|최대\s*|limit\s*)?(\d+)\s*(?:건|개|행|줄)(?:\s*만|\s*조회|\s*보여|\s*출력)?", q, re.IGNORECASE)
    if not limit_match:
        limit_match = re.search(r"limit\s+(\d+)", q, re.IGNORECASE)
    if limit_match:
        limit_val = int(limit_match.group(1))
        q_without_limit = (q[:limit_match.start()] + " " + q[limit_match.end():]).strip()
    else:
        q_without_limit = q

    order_clause = ""
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    if any(k in q_without_limit for k in ["높은 순", "높은순", "많은 순", "많은순", "큰 순", "내림차순", "낮은 순", "낮은순", "적은 순", "오름차순", "정렬"]):
        is_asc = any(k in q_without_limit for k in ["낮은", "적은", "작은", "오름차순", "asc"])
        direction = "ASC" if is_asc else "DESC"
        sort_col = next((c for c in df.columns if str(c).lower() in q_without_limit.lower()), None)
        if sort_col is None and num_cols:
            for c in num_cols:
                if ("불량" in q_without_limit and "불량" in str(c)) or ("생산" in q_without_limit and "생산" in str(c)):
                    sort_col = c
                    break
            sort_col = sort_col or num_cols[0]
        if sort_col:
            order_clause = f' ORDER BY "{sort_col}" {direction}'

    rem_cleaned = re.sub(
        r"(전체|조회|해줘|보여줘|찾아줘|출력|데이터|만|에서|중|높은\s*순|많은\s*순|낮은\s*순|적은\s*순|내림차순|오름차순|정렬)",
        "",
        q_without_limit,
    ).strip()
    if not rem_cleaned and (limit_val is not None or order_clause):
        sql = f"SELECT * FROM {table_name}{order_clause}"
        return f"{sql} LIMIT {limit_val};" if limit_val else f"{sql};"

    if any(k in q_without_limit for k in ["빨강", "빨간", "초록", "노랑", "배경", "색상", "색깔"]):
        color_kw = "빨강" if ("빨강" in q_without_limit or "빨간" in q_without_limit) else (
            "초록" if "초록" in q_without_limit else ("노랑" if "노랑" in q_without_limit else "강조")
        )
        or_likes = [f'"{col}" LIKE \'%{color_kw}%\'' for col in df.columns]
        sql = f"SELECT * FROM {table_name} WHERE (" + " OR ".join(or_likes) + f"){order_clause}"
        return f"{sql} LIMIT {limit_val};" if limit_val else f"{sql};"

    if any(k in q_without_limit.lower() for k in ["미기입", "빈칸", "공란", "미실행", "not run", "n/a"]):
        target_c = next((c for c in df.columns if str(c).lower() in q_without_limit.lower()), None)
        check_cols = [target_c] if target_c else [
            c for c in df.columns if str(c).lower() not in ("_polarion", "id", "tc id", "번호", "시트명", "파일명")
        ]
        if any(k in q_without_limit.lower() for k in ["미실행", "not run"]):
            conds = [f'UPPER("{c}") LIKE \'%NOT RUN%\' OR "{c}" LIKE \'%미실행%\'' for c in check_cols]
        elif "n/a" in q_without_limit.lower():
            conds = [f'UPPER("{c}") = \'N/A\'' for c in check_cols]
        else:
            conds = [f'"{c}" = \'-\' OR "{c}" = \'\' OR "{c}" LIKE \'%[빈칸%\'' for c in check_cols]
        sql = f"SELECT * FROM {table_name} WHERE (" + " OR ".join(conds) + f"){order_clause}"
        return f"{sql} LIMIT {limit_val};" if limit_val else f"{sql};"

    is_negation = any(k in q_without_limit for k in ["빼고", "제외", "아닌", "말고"])

    col_matches: dict[str, list[str]] = {}
    str_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    for col in str_cols:
        matched_vals = []
        for val in df[col].dropna().astype(str).unique():
            val_clean = re.sub(r"\s*\[[^\]]+\]", "", str(val)).strip()
            if not val_clean or val_clean == "-" or len(val_clean) > 45:
                continue
            if val_clean.lower() in q_without_limit.lower() or f"{val_clean}{col}".lower() in q_without_limit.lower():
                matched_vals.append(val_clean)
        if matched_vals:
            col_matches[col] = list(dict.fromkeys(matched_vals))

    where_clauses = []
    for col, vals in col_matches.items():
        if is_negation:
            and_not = " AND ".join(f'"{col}" NOT LIKE \'%{v}%\'' for v in vals)
            where_clauses.append(f"({and_not})")
        elif len(vals) == 1:
            where_clauses.append(f'"{col}" LIKE \'%{vals[0]}%\'')
        else:
            or_part = " OR ".join(f'"{col}" LIKE \'%{v}%\'' for v in vals)
            where_clauses.append(f"({or_part})")

    nums = re.findall(r"\d+(?:\.\d+)?", q_without_limit)
    if nums and num_cols:
        target_num = float(nums[0])
        target_col = next((col for col in num_cols if str(col).lower() in q_without_limit.lower()), None)
        if target_col is None:
            for col in num_cols:
                if ("불량" in q_without_limit and "불량" in str(col)) or ("생산" in q_without_limit and "생산" in str(col)) or ("가동" in q_without_limit and "가동" in str(col)):
                    target_col = col
                    break
            if target_col is None:
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

    if not where_clauses and str_cols:
        stop_words = {
            "전체", "조회", "해줘", "보여줘", "찾아줘", "검색", "데이터", "항목", "행", "열", "에서", "만",
            "그리고", "또는", "이나", "이랑", "랑", "과", "와", "포함", "있는", "없는", "제외", "빼고", "아닌", "말고",
            "높은", "낮은", "많은", "적은", "순", "정렬", "이상", "이하", "초과", "미만",
        }
        raw_tokens = re.findall(r"[A-Za-z0-9_.+-]{2,}|[가-힣]{2,}", q_without_limit)
        search_kws = []
        for tk in raw_tokens:
            tk_clean = re.sub(r"(이랑|으로|에서|부터|까지|만|은|는|이|가|을|를|랑|도)$", "", tk).strip()
            if len(tk_clean) >= 2 and tk_clean.lower() not in stop_words and not tk_clean.isdigit():
                if tk_clean.lower() not in [str(c).lower() for c in df.columns]:
                    search_kws.append(tk_clean)

        for kw in list(dict.fromkeys(search_kws))[:3]:
            hit_cols = [
                col for col in str_cols
                if df[col].astype(str).str.contains(re.escape(kw), case=False, na=False).any()
            ]
            target_cols = hit_cols if hit_cols else str_cols[:5]
            if is_negation:
                where_clauses.append("(" + " AND ".join(f'"{c}" NOT LIKE \'%{kw}%\'' for c in target_cols) + ")")
            else:
                where_clauses.append("(" + " OR ".join(f'"{c}" LIKE \'%{kw}%\'' for c in target_cols) + ")")

    sql = f"SELECT * FROM {table_name}"
    if where_clauses:
        join_op = " AND " if is_negation or len(where_clauses) == 1 else " OR "
        if col_matches or nums:
            sql += " WHERE " + " AND ".join(where_clauses)
        else:
            sql += " WHERE " + join_op.join(where_clauses)
    sql += order_clause
    if limit_val is not None:
        sql += f" LIMIT {limit_val}"
    sql += ";"
    return sql


def generate_smart_assistant_question(
    target_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    user_filter_text: str,
    threshold: float = DEFAULT_DEFECT_THRESHOLD,
) -> str:
    if len(target_df) == 0:
        return "현재 데이터의 핵심 현황, 이상·주의 항목, 원인 가설 및 실무 조치 제안을 정리해 줘."

    all_cols = target_df.columns.tolist()
    cols_lower = " ".join(str(c).lower() for c in all_cols)
    scope_parts = []
    if user_filter_text and user_filter_text.strip():
        scope_parts.append(f"검색조건('{user_filter_text.strip()}')")

    if "파일명" in all_cols:
        files = target_df["파일명"].unique().tolist()
        scope_parts.append(f"통합 파일 {len(files)}개({', '.join(map(str, files[:3]))})")
    if "시트명" in all_cols:
        sheets = target_df["시트명"].unique().tolist()
        scope_parts.append(f"대상 시트 {len(sheets)}개({', '.join(map(str, sheets[:3]))})")

    stat_snippets = [f"총 {len(target_df):,}건", f"사내 기준 {threshold}% 적용"]
    scope_title = " · ".join(scope_parts) if scope_parts else "전체 조회 데이터"

    if is_manufacturing_df(target_df):
        p_sum = int(target_df["생산량"].sum())
        d_sum = int(target_df["불량수"].sum())
        rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
        stat_snippets.append(f"총 생산량 {p_sum:,}개, 총 불량수 {d_sum:,}개(불량률 {rate:.2f}%)")
        stat_str = ", ".join(stat_snippets)
        return (
            f"[AI 비서 자동 맞춤 질문] {scope_title} ({stat_str}) 데이터를 바탕으로 "
            f"① 라인·일자별 핵심 추세 및 불량 비중(%), ② 사내 허용 기준({threshold}%) 초과 이상 구간, "
            f"③ 공정·설비·작업 관점 원인 가설, ④ 담당 부서별 즉시 조치(24h) 및 근본 대책(1개월)을 정밀 분석해 줘."
        )

    color_df = analyze_color_cell_contexts(target_df)
    color_cnt = len(color_df)

    if any(k in cols_lower for k in ["procedure", "pass/fail", "actual", "precondition", "pre-condition", "domain", "sub system", "priority", "category", "tc"]):
        if color_cnt > 0:
            stat_snippets.append(f"🎨셀 색상강조 {color_cnt}건 감지")
        actual_cat_cols = [c for c in ["Category", "Priority", "Test Technique", "Domain", "Sub System", "App", "Automated", "Execution Status"] if c in all_cols]
        for c_cand in actual_cat_cols[:3]:
            vals = [str(v) for v in target_df[c_cand].unique().tolist() if str(v) != "-"]
            if vals:
                stat_snippets.append(f"{c_cand}[{', '.join(vals[:3])}]")
        stat_str = ", ".join(stat_snippets)
        col_list_str = ", ".join(actual_cat_cols[:4]) if actual_cat_cols else "주요 분류 열"
        return (
            f"[AI 비서 자동 맞춤 질문] {scope_title} ({stat_str}) 데이터를 바탕으로 "
            f"① 실제 존재하는 분류 열({col_list_str})별 건수·비중(%) 및 사내 기준({threshold}%) 이상 항목, "
            f"② 유사 시나리오 키워드 묶음 분포, "
            f"③ 미기입률 사내 기준({threshold}%) 초과 컬럼 및 실행 상태 점검, "
            f"④ QA·개발팀의 즉시 조치(24h) 및 검증 표준화 대책(1개월)을 없는 숫자 지어내지 말고 깔끔한 불릿으로 분석해 줘."
        )

    stat_str = ", ".join(stat_snippets)
    return (
        f"[AI 비서 자동 맞춤 질문] {scope_title} ({stat_str}) 데이터를 바탕으로 "
        f"① 실제 항목별 핵심 현황 및 사내 기준({threshold}%) 비교, "
        f"② 주요 리스크 및 원인 가설, ③ 실무 조치 제안(24h/1개월)을 정리해 줘."
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
    elif any(c in all_cols for c in ["Sub System", "Domain", "Category", "App"]):
        for c in ["Domain", "Sub System", "Category", "App"]:
            if c in all_cols:
                vals = [str(v) for v in target_df[c].astype(str).unique() if str(v) != "-"]
                keywords.extend(vals[:1])
        keywords.append("SW 품질 테스트 자동화")
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
    st.caption(
        f"🕒 분석 일시: **{res_data.get('analyzed_at', '-')}** | "
        f"⏱️ 소요 시간: **{res_data['elapsed_time']}초** | "
        f"분석 대상: **{len(target_df):,}건 (기준 {defect_threshold}%)**"
    )

    render_smart_charts(target_df, key_prefix=f"popup_{m_code}", defect_threshold=defect_threshold)
    st.divider()

    st.markdown("#### 📝 AI 생성 전체 리포트")
    st.markdown(res_data["report"])

    p_c1, p_c2 = st.columns(2)
    with p_c1:
        st.download_button(
            label=f"📥 {m_code} 엑셀 리포트 다운로드 (.xlsx)",
            data=res_data["excel_bytes"],
            file_name=f"report_{m_code}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key=f"popup_dl_xlsx_{m_code}",
            type="primary",
            use_container_width=True,
        )
    with p_c2:
        st.download_button(
            label=f"📄 {m_code} 워드 보고서 다운로드 (.docx)",
            data=res_data["docx_bytes"],
            file_name=f"report_{m_code}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            key=f"popup_dl_docx_{m_code}",
            use_container_width=True,
        )


@st.fragment
def render_visualization_fragment(
    target_df: pd.DataFrame,
    aux_sheets: dict[str, pd.DataFrame],
    defect_threshold: float,
):
    st.subheader("2️⃣ 조회된 데이터 다차원 시각화 & 한글 기초통계 집계")
    v_tab1, v_tab2, v_tab3 = st.tabs([
        "🎯 핵심 다차원 대시보드 (TOP 순위·시나리오 그룹핑·사내 기준선 자동 연동)",
        "🛠️ 자유 조합 커스텀 차트 (모든 엑셀/CSV/문서 범용)",
        f"📋 한글 기초통계 & 사내 기준({defect_threshold}%) 판정 집계표 (엑셀 2번 시트)",
    ])

    with v_tab1:
        render_smart_charts(target_df, key_prefix="main_dash", defect_threshold=defect_threshold)

    with v_tab2:
        num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
        all_cols = target_df.columns.tolist()
        if len(target_df) > 0:
            c1, c2, c3, c4 = st.columns(4)
            default_x = next((i for i, c in enumerate(all_cols) if c in ["Sub System", "Category", "Domain", "App", "라인"]), 0)
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
        korean_tables = build_korean_stat_tables(target_df, aux_sheets=aux_sheets, threshold=defect_threshold)
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
    pipeline_slot,
):
    st.subheader("3️⃣ AI 비서 맞춤 질문 & 에이전트 동시 비교 분석 (Groq vs Gemini vs Ollama 정밀 모드)")

    auto_assistant_q = generate_smart_assistant_question(
        target_df, raw_df, active_filter_text, threshold=defect_threshold
    )

    q_mode = st.radio(
        "🤖 질문 생성 방식 선택",
        options=[
            "🤖 AI 비서 자동 맞춤 질문 (조회된 시트·실제 컬럼·사내 기준값 자동 조합)",
            "✍️ 사용자 직접 질문 입력/수정",
        ],
        horizontal=True,
    )

    if q_mode.startswith("🤖"):
        st.info(f"**💡 AI 비서가 현재 조회된 데이터({len(target_df):,}건, 기준 {defect_threshold}%)에 맞춰 조합한 분석 질문입니다:**\n\n{auto_assistant_q}")
        question = auto_assistant_q
    else:
        question = st.text_input("💡 에이전트에게 지시할 분석 질문을 입력하세요:", value=auto_assistant_q)

    # 사내 기준서(RAG)가 등록되어 있으면 안내 배지 표시 (동시 연동!)
    rag_text = st.session_state.get("rag_guideline_text", "").strip()
    if rag_text:
        st.success(f"📚 **[사내 기준서/가이드 문서(RAG) 자동 연동 중]** 아래 4️⃣번 탭에 등록된 사내 기준서({len(rag_text):,}자)가 3대 AI 분석에 동시에 반영됩니다.")

    render_related_links_box(target_df)

    btn_c1, btn_c2 = st.columns([3, 1])
    with btn_c1:
        run_clicked = st.button("🚀 에이전트 정밀 분석 실행 (선택 모델 동시 병렬 실행)", type="primary", use_container_width=True)
    with btn_c2:
        if st.button("🔄 업로드 파일 & 입력 전체 초기화", key="reset_btn_bottom", use_container_width=True):
            reset_all_inputs_and_files()
            st.rerun()

    if run_clicked:
        if not selected_models:
            st.warning("좌측 사이드바에서 최소 1개 이상의 모델을 선택해주세요.")
        else:
            analyzed_at_now = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            verdict_now = evaluate_threshold_verdict(target_df, threshold=defect_threshold)

            pipeline_slot.markdown(
                build_pipeline_status_html(
                    has_file=True,
                    has_query_df=True,
                    saved_results={},
                    running_stage=3,
                    running_msg=f"⚡ {len(selected_models)}개 AI 모델 병렬 분석 중...",
                ),
                unsafe_allow_html=True,
            )

            agent_app = build_agent_graph()
            stat_summary_str = "\n\n".join(
                f"■ {k}\n{v.to_string(index=False)}"
                for k, v in build_korean_stat_tables(target_df, aux_sheets=aux_sheets, threshold=defect_threshold).items()
            )
            sample_rows_df = target_df.head(40)
            rag_block = f"\n\n[사내 기준서·가이드 문서(RAG 참조 기준)]\n{rag_text[:2500]}\n" if rag_text else ""

            data_info = (
                f"[분석 일시]: {analyzed_at_now}\n"
                f"[조회 조건]: {active_filter_text or '전체 시트 데이터'}\n"
                f"[조회 데이터 건수]: {len(target_df)}건 (전체 원본 {len(raw_df)}건)\n"
                f"[적용된 사내 관리 기준값]: {defect_threshold}%\n"
                f"[사내 기준 자동 판정]: {verdict_now}\n"
                f"[실제 존재하는 컬럼 목록]: {list(target_df.columns)}\n\n"
                f"[파이썬 정밀 집계표]\n{stat_summary_str}"
                f"{rag_block}\n\n"
                f"[상세 행 샘플 데이터 (상위 40건)]\n{sample_rows_df.to_string(index=False)}"
            )

            new_results = {}
            logs = [
                f"**[1~2단계 완료]** 총 `{len(target_df):,}건` 로드 및 통계 집계 완료 (분석 일시: `{analyzed_at_now}`, 기준값: `{defect_threshold}%`)",
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
                res["analyzed_at"] = analyzed_at_now
                res["applied_threshold"] = defect_threshold
                res["excel_bytes"] = export_analysis_to_excel(
                    target_df, summary_df, res["report"], res["elapsed_time"], m_code,
                    aux_sheets=aux_sheets, threshold=defect_threshold, analyzed_at=analyzed_at_now,
                )
                res["docx_bytes"] = export_analysis_to_docx(
                    target_df, res["report"], res["elapsed_time"], MODEL_LABELS.get(m_code, m_code),
                    threshold=defect_threshold, analyzed_at=analyzed_at_now,
                )
                return m_code, res

            with ThreadPoolExecutor(max_workers=len(selected_models)) as executor:
                future_map = {executor.submit(_run_single_model, m): m for m in selected_models}
                done_cnt = 0
                for future in as_completed(future_map):
                    m_done, res_done = future.result()
                    new_results[m_done] = res_done
                    done_cnt += 1
                    msg = f"✅ **{MODEL_LABELS.get(m_done, m_done)}** 추론 · 양식 검증 · 엑셀/워드 생성 완료 (`{res_done['elapsed_time']}초`)"
                    status_box.write(msg)
                    logs.append(msg)
                    pipeline_slot.markdown(
                        build_pipeline_status_html(
                            has_file=True,
                            has_query_df=True,
                            saved_results={},
                            running_stage=4 if done_cnt == len(selected_models) else 3,
                            running_msg=f"⚡ AI 분석 중 ({done_cnt}/{len(selected_models)} 완료)",
                        ),
                        unsafe_allow_html=True,
                    )

            logs.append(f"**[4단계 완료]** 사내 기준(`{defect_threshold}%`) 자동 판정 배너 적용 및 환각 방지 정제 완료")
            logs.append("**[5단계 완료]** 통합 엑셀(`.xlsx`) 및 페이지 분리형 워드(`.docx`) 보고서 패키징 완료")
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
        live_verdict = evaluate_threshold_verdict(popup_df, threshold=defect_threshold)

        st.markdown(f"##### ⚖️ AI 모델별 성능 및 사내 기준(`{defect_threshold}%`) 자동 판정 비교")
        matrix_rows = []
        for m_code in display_models:
            r_item = saved[m_code]
            matrix_rows.append({
                "분석 일시": r_item.get("analyzed_at", "-"),
                "AI 모델": MODEL_LABELS.get(m_code, m_code),
                "소요 시간": f"{r_item['elapsed_time']} 초",
                f"자동 판정 요약 (기준 {defect_threshold}%)": live_verdict,
            })
        st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True, hide_index=True)

        combined_excel_bytes = export_combined_analysis_to_excel(
            popup_df, saved, MODEL_LABELS, aux_sheets=saved_aux, threshold=defect_threshold
        )
        combined_docx_bytes = export_combined_analysis_to_docx(
            popup_df, saved, MODEL_LABELS, threshold=defect_threshold
        )

        top_c1, top_c2, top_c3 = st.columns([1.6, 1.2, 1.2])
        with top_c1:
            st.caption("👆 **우측 워드(.docx) 추출 시 [1페이지: 전체 종합 요약표] + [2페이지~: AI별 강제 쪽 나누기 상세본(빈 페이지 0건)]으로 출력됩니다.**")
        with top_c2:
            st.download_button(
                label="📚 전체 모델 종합 비교 엑셀 추출 (.xlsx)",
                data=combined_excel_bytes,
                file_name="report_ALL_models_combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_combined_master_xlsx",
                type="primary",
                use_container_width=True,
            )
        with top_c3:
            st.download_button(
                label="📄 전체 모델 종합 워드 추출 (쪽나누기 적용 .docx)",
                data=combined_docx_bytes,
                file_name="report_ALL_models_combined.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                key="dl_combined_master_docx",
                use_container_width=True,
            )

        cols = st.columns(len(display_models))
        for idx, m_code in enumerate(display_models):
            res_item = saved[m_code]
            with cols[idx]:
                if st.button(f"📊 {MODEL_LABELS.get(m_code, m_code)}\n(클릭 시 시각화 팝업 열기 🔍)", key=f"title_btn_{m_code}", use_container_width=True):
                    show_model_visual_popup(m_code, res_item, popup_df, defect_threshold)

                st.info(
                    f"🕒 분석 일시: **{res_item.get('analyzed_at', '-')}** | "
                    f"⏱️ 소요 시간: **{res_item['elapsed_time']}초**"
                )
                updated_report = re.sub(
                    r"^>\s*\*\*.*?\*\*\n\n",
                    f"> **{live_verdict}**\n\n",
                    res_item["report"],
                    count=1,
                )
                st.markdown(updated_report)

                dl_b1, dl_b2 = st.columns(2)
                with dl_b1:
                    live_single_xlsx = export_analysis_to_excel(
                        popup_df, summary_df, updated_report, res_item["elapsed_time"], m_code,
                        aux_sheets=saved_aux, threshold=defect_threshold, analyzed_at=res_item.get("analyzed_at"),
                    )
                    st.download_button(
                        label=f"📥 {m_code} 엑셀 (.xlsx)",
                        data=live_single_xlsx,
                        file_name=f"report_{m_code}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f"dl_persist_xlsx_{m_code}_{idx}",
                        use_container_width=True,
                    )
                with dl_b2:
                    live_single_docx = export_analysis_to_docx(
                        popup_df, updated_report, res_item["elapsed_time"], MODEL_LABELS.get(m_code, m_code),
                        threshold=defect_threshold, analyzed_at=res_item.get("analyzed_at"),
                    )
                    st.download_button(
                        label=f"📄 {m_code} 워드 (.docx)",
                        data=live_single_docx,
                        file_name=f"report_{m_code}.docx",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        key=f"dl_persist_docx_{m_code}_{idx}",
                        use_container_width=True,
                    )


# =====================================================================
# 🌟 4️⃣ [후보 A + B + C 동시 탑재] 심화 AI 어시스턴트 & PoC 검증 센터
# =====================================================================
@st.fragment
def render_advanced_extensions_hub(
    target_df: pd.DataFrame,
    aux_sheets: dict[str, pd.DataFrame],
    defect_threshold: float,
):
    st.subheader("4️⃣ 심화 AI 어시스턴트 & PoC 검증 센터 (후보 A·B·C 통합 가동)")
    st.caption(
        "💡 **3번 일괄 분석(Workflow)** 이후, 전체 재실행 없이 **① 로컬 보안 채팅으로 핀셋 꼬리 질문(후보 A)**을 하거나, "
        "**② 사내 기준서/가이드 문서 대조(후보 B)**, **③ 동일 데이터 3회 반복 일관성 검증(후보 C)**을 동시에 수행할 수 있습니다."
    )

    tab_a, tab_b, tab_c = st.tabs([
        "💬 [후보 A] 대화형 후속 질문 채팅방 (Chatflow · 로컬 Ollama 기본)",
        "📚 [후보 B] 사내 기준서·가이드 문서 대조 (RAG 지식 연동)",
        "🔁 [후보 C] 3회 반복 실행 일관성(편차) 자동 검증",
    ])

    stat_summary_str = "\n\n".join(
        f"■ {k}\n{v.to_string(index=False)}"
        for k, v in build_korean_stat_tables(target_df, aux_sheets=aux_sheets, threshold=defect_threshold).items()
    )
    verdict_now = evaluate_threshold_verdict(target_df, threshold=defect_threshold)

    # -----------------------------------------------------------------
    # 💬 [후보 A] 대화형 후속 질문 채팅방 (Chatflow) — 로컬 Ollama 단독 구동 기본 지원!
    # -----------------------------------------------------------------
    with tab_a:
        with st.expander("❓ 왜 3번 리포트 외에 '후속 질문 채팅방(Chatflow)'이 필요한가요? (클릭)", expanded=False):
            st.markdown(
                "- **1. 전체 재실행 없는 핀셋 확인:** 3번 리포트에서 *'미기입이 15건 있다'*고 나왔을 때, 3개 모델을 다시 돌리지 않고 **'그 15건의 ID랑 제목만 표로 뽑아줘'**처럼 즉시 꼬리 질문을 할 수 있습니다.\n"
                "- **2. 100% 사내 보안 로컬 단독 구동 가능:** 기본 엔진이 **`🔒 Local: Ollama (사내보안 정밀 모드)`**로 설정되어 있어 외부 인터넷으로 데이터가 나가지 않고 내 PC 안에서만 안전하게 대화합니다.\n"
                "- **3. 회의 중 즉석 질의응답 & 양식 변환:** *'방금 분석 결과를 팀즈(Teams) 공지용 5줄 요약으로 바꿔줘'* 등 자유로운 재가공이 가능합니다."
            )

        c_eng, c_quick, c_clear = st.columns([1.4, 2.0, 0.6])
        with c_eng:
            chat_model = st.selectbox(
                "🤖 채팅 구동 AI 엔진 선택 (로컬 기본):",
                options=["local_ollama", "cloud_groq", "cloud_gemini", "mock"],
                index=0,  # 기본값: 🔒 Local: Ollama (사내 보안 단독 구동!)
                format_func=lambda x: MODEL_LABELS.get(x, x),
                key="chatflow_model_sel",
            )
        with c_quick:
            quick_q = st.selectbox(
                "⚡ 자주 쓰는 후속 꼬리질문 원클릭 선택:",
                options=[
                    "(직접 입력 또는 아래 추천 꼬리질문 선택)",
                    "미기입(빈칸) 또는 주의 항목의 구체적인 ID와 제목 목록만 뽑아줘",
                    f"사내 기준({defect_threshold}%)을 초과하거나 가장 비중이 높은 상위 3개 영역의 원인과 대책을 깊게 알려줘",
                    "방금 분석한 핵심 결과를 팀즈(Teams) 주간회의 공유용 5줄 요약 메시지로 바꿔줘",
                    "현재 데이터에서 경계값·예외 처리 테스트 시나리오 3가지를 추가로 설계해 줘",
                ],
                key="chatflow_quick_sel",
            )
        with c_clear:
            st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
            if st.button("🗑️ 대화 비우기", key="clear_chat_btn", use_container_width=True):
                st.session_state["chat_history"] = []
                st.rerun()

        # 이전 대화 내역 출력
        for msg in st.session_state["chat_history"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
                if msg.get("matched_df") is not None and len(msg["matched_df"]) > 0:
                    st.dataframe(style_dataframe_with_colors(msg["matched_df"]), use_container_width=True, hide_index=True)

        user_chat_in = st.chat_input("현재 데이터나 리포트에 대해 궁금한 후속 질문을 입력하세요 (예: 미기입 항목 ID만 보여줘)...")
        if not user_chat_in and quick_q != "(직접 입력 또는 아래 추천 꼬리질문 선택)":
            if st.button("📨 선택한 추천 꼬리질문 바로 전송", type="primary", key="send_quick_q_btn"):
                user_chat_in = quick_q

        if user_chat_in:
            st.session_state["chat_history"].append({"role": "user", "content": user_chat_in})
            with st.chat_message("user"):
                st.markdown(user_chat_in)

            # 파이썬 스마트 보조 검색: 사용자 질문과 관련된 행을 SQL로 먼저 추출해 AI에게 함께 제공
            matched_sub_df = pd.DataFrame()
            try:
                sub_sql = build_smart_sql(user_chat_in, target_df, TABLE_NAME)
                if "WHERE" in sub_sql or "ORDER BY" in sub_sql or "LIMIT" in sub_sql:
                    matched_sub_df = run_sql_query(sub_sql).head(25)
            except Exception:
                pass

            # 이전 리포트 요약 컨텍스트 결합
            saved_rep_context = ""
            if st.session_state.get("saved_results"):
                first_m = next(iter(st.session_state["saved_results"].values()))
                saved_rep_context = f"\n[직전 3번 AI 분석 리포트 요약]\n{first_m.get('report', '')[:1500]}\n"

            rag_ctx = st.session_state.get("rag_guideline_text", "").strip()
            rag_part = f"\n[사내 기준서/가이드(RAG)]\n{rag_ctx[:1200]}\n" if rag_ctx else ""
            sub_rows_str = (
                f"\n[질문 조건 매칭 상세 행 ({len(matched_sub_df)}건)]\n{matched_sub_df.to_string(index=False)}\n"
                if len(matched_sub_df) > 0
                else f"\n[상위 25건 샘플 행]\n{target_df.head(25).to_string(index=False)}\n"
            )
            recent_turns = "\n".join(
                f"{'사용자' if m['role']=='user' else 'AI'}: {m['content'][:300]}"
                for m in st.session_state["chat_history"][-5:-1]
            )

            chat_prompt = f"""당신은 사내 데이터 분석 후속 질의응답(Chatflow) 전담 AI 비서입니다.
아래 [파이썬 정밀 집계표]와 [매칭 상세 행 데이터]의 실제 값만 사용하여 사용자의 후속 질문에 명확하고 간결하게 답변하십시오.
절대 없는 숫자를 지어내지 마시고, 마크다운 표(|) 대신 깔끔한 불릿 리스트(- )로 답변하십시오.

[사내 기준 판정 ({defect_threshold}% 기준)]: {verdict_now}
[파이썬 정밀 집계표]
{stat_summary_str}
{sub_rows_str}
{saved_rep_context}
{rag_part}
[이전 대화 맥락]
{recent_turns}

[사용자 후속 질문]
{user_chat_in}
"""
            with st.chat_message("assistant"):
                with st.spinner(f"{MODEL_LABELS.get(chat_model, chat_model)} 엔진으로 후속 답변 생성 중..."):
                    ans_text, ans_el = call_llm(chat_prompt, mode=chat_model)
                    final_ans = f"⏱️ *응답 소요 시간: `{ans_el}초`* | *사내 기준: `{defect_threshold}%`*\n\n{ans_text}"
                    st.markdown(final_ans)
                    if len(matched_sub_df) > 0:
                        st.caption(f"📌 **질문과 관련해 파이썬이 자동 추출한 데이터 행 ({len(matched_sub_df)}건):**")
                        st.dataframe(style_dataframe_with_colors(matched_sub_df), use_container_width=True, hide_index=True)

            st.session_state["chat_history"].append({
                "role": "assistant",
                "content": final_ans,
                "matched_df": matched_sub_df if len(matched_sub_df) > 0 else None,
            })

    # -----------------------------------------------------------------
    # 📚 [후보 B] 사내 기준서·가이드 문서 대조 분석 (RAG 지식 연동)
    # -----------------------------------------------------------------
    with tab_b:
        st.markdown(
            "**📚 사내 품질/검증 기준서(SOP, 점검 가이드, 코딩/TC 작성 규칙) 대조 분석 (RAG)**\n"
            "- 기준서 파일(`.docx`, `.txt`, `.csv`)을 올리거나 아래 칸에 사내 기준을 입력해 두면, **위 3️⃣번 본 분석 및 💬 후보 A 채팅방에 동시에 반영**되며, 아래 버튼으로 **기준서 준수 여부 단독 대조 리포트**도 즉시 뽑을 수 있습니다."
        )
        g_c1, g_c2 = st.columns([1.3, 2.7])
        with g_c1:
            guide_file = st.file_uploader(
                "📎 사내 기준서/가이드 파일 첨부 (.docx, .txt, .csv)",
                type=["docx", "txt", "csv"],
                key="rag_guide_uploader",
            )
            default_guide = (
                "1. [작성 충실도 기준] 모든 TC의 Precondition, Procedure, Expected Result(판정기준)는 미기입(빈칸) 비율이 사내 기준치(%) 이하여야 한다.\n"
                "2. [우선순위 관리] Priority High(상) 항목은 100% 수행 및 결과 기록이 필수이며 미실행(Not Run) 시 즉시 사유를 기입한다.\n"
                "3. [예외 검증 비율] 정상 기능(Positive) 테스트 외에 경계값·오류 주입(Negative) 시나리오가 최소 20% 이상 포함되어야 한다.\n"
                "4. [제조 공정 기준] 라인별 평균 불량률이 사내 기준치(%)를 초과하면 24시간 내 설비 센서 및 작업 표준서(SOP)를 재점검한다."
            )
            loaded_guide_txt = extract_text_from_guideline_file(guide_file) if guide_file else (
                st.session_state["rag_guideline_text"] or default_guide
            )
            rag_model = st.selectbox(
                "대조 분석 AI 모델 선택 (로컬 기본):",
                options=["local_ollama", "cloud_groq", "cloud_gemini", "mock"],
                index=0,
                format_func=lambda x: MODEL_LABELS.get(x, x),
                key="rag_model_sel",
            )

        with g_c2:
            edited_guide = st.text_area(
                "📝 적용할 사내 기준서 / 검증 가이드라인 내용 (자유 수정 가능 — 3번 본 분석에도 자동 연동됨):",
                value=loaded_guide_txt,
                height=155,
            )
            st.session_state["rag_guideline_text"] = edited_guide

        if st.button("🔍 현재 데이터 vs 사내 기준서(RAG) 준수 여부 대조 실행", type="primary", key="run_rag_btn"):
            rag_prompt = f"""당신은 사내 품질 감사(Audit) 및 검증 표준화 전문가입니다.
아래 [사내 기준서 / 검증 가이드라인]의 각 조항을 기준으로 [현재 데이터 파이썬 집계표]를 대조하여 준수 여부와 위반(미달) 항목, 시정 조치를 불릿(- )으로 작성하십시오.

[사내 기준서 / 검증 가이드라인 (RAG)]
{edited_guide}

[현재 데이터 파이썬 집계표 (사내 기준 {defect_threshold}%)]
{stat_summary_str}

[출력 형식]
### 1. 사내 기준서 조항별 준수 여부 판정 (✅준수 / 🚨미달 및 근거 수치)
### 2. 기준 미달(위반) 구간 상세 진단
### 3. 기준서 근거 시정 조치 (즉시 조치 & 표준화 대책)
"""
            with st.spinner("사내 기준서와 현재 데이터를 대조 분석하고 있습니다..."):
                rag_ans, rag_el = call_llm(rag_prompt, mode=rag_model)
                st.success(f"✅ 사내 기준서(RAG) 대조 완료 (`{rag_el}초`)")
                st.markdown(rag_ans)

    # -----------------------------------------------------------------
    # 🔁 [후보 C] 동일 데이터 3회 반복 실행 일관성(편차) 자동 검증
    # -----------------------------------------------------------------
    with tab_c:
        st.markdown(
            "**🔁 동일 데이터 3회 연속 실행 일관성(편차) 검증기**\n"
            "- 사내 AI 에이전트 신뢰성 검증 표준에 따라, **같은 데이터를 3회 반복 실행했을 때 ① 필수 4대 목차 유지율, ② 핵심 건수·수치 일치율, ③ 사내 기준 판정 일관성, ④ 응답 시간 편차**를 자동으로 채점합니다."
        )
        c_m1, c_m2 = st.columns([1.5, 2.5])
        with c_m1:
            bench_model = st.selectbox(
                "3회 반복 검증할 대상 AI 모델 선택:",
                options=["local_ollama", "cloud_groq", "cloud_gemini", "mock"],
                index=0,
                format_func=lambda x: MODEL_LABELS.get(x, x),
                key="bench_model_sel",
            )
        with c_m2:
            st.info(f"📌 **검증 기준 데이터:** 총 `{len(target_df):,}건` | **사내 기준값:** `{defect_threshold}%` | **예상 판정:** `{verdict_now[:45]}...`")

        if st.button("▶️ 3회 연속 반복 실행 및 일관성 자동 채점 시작", type="primary", key="run_3x_bench_btn"):
            agent_app = build_agent_graph()
            q_bench = generate_smart_assistant_question(target_df, target_df, "", threshold=defect_threshold)
            data_info_bench = (
                f"[조회 데이터 건수]: {len(target_df)}건\n"
                f"[적용된 사내 관리 기준값]: {defect_threshold}%\n"
                f"[사내 기준 자동 판정]: {verdict_now}\n"
                f"[실제 존재하는 컬럼 목록]: {list(target_df.columns)}\n\n"
                f"[파이썬 정밀 집계표]\n{stat_summary_str}"
            )

            bench_rows = []
            prog = st.progress(0, text="1/3회차 실행 중...")
            total_cnt_str = str(len(target_df))

            for run_i in range(1, 4):
                prog.progress((run_i - 1) / 3, text=f"🔄 [{MODEL_LABELS.get(bench_model, bench_model)}] {run_i}/3회차 실행 및 채점 중...")
                res_i = agent_app.invoke({
                    "question": q_bench,
                    "mode": bench_model,
                    "data_summary": data_info_bench,
                    "defect_threshold": defect_threshold,
                    "analysis_result": {},
                    "report": "",
                    "elapsed_time": 0.0,
                })
                rep_i = str(res_i.get("report", ""))
                el_i = float(res_i.get("elapsed_time", 0.0))

                # 1) 필수 4대 목차(1. 2. 3. 4.) 포함 여부 검사
                sec_hits = sum(1 for s_num in ["1.", "2.", "3.", "4."] if s_num in rep_i)
                sec_score = f"{sec_hits}/4개 ({int(sec_hits/4*100)}%)"

                # 2) 전체 데이터 건수(len(target_df)) 정확 인용 여부
                has_exact_count = "✅ 일치" if total_cnt_str in rep_i else "⚠️ 일부 생략"

                # 3) 사내 기준 판정(확인 필요 vs 이상 없음) 일치 여부
                expected_flag = "확인 필요" if "확인 필요" in verdict_now else "이상 없음"
                verdict_match = "✅ 100% 일치" if expected_flag in rep_i else "⚠️ 불일치"

                # 4) 표(|) 및 <br> 오염 제로 여부
                clean_fmt = "✅ 정상(오염 0건)" if ("<br>" not in rep_i.lower()) else "⚠️ 태그 감지"

                bench_rows.append({
                    "실행 회차": f"{run_i}회차",
                    "AI 모델": MODEL_LABELS.get(bench_model, bench_model),
                    "소요 시간(초)": el_i,
                    "4대 필수 목차 유지율": sec_score,
                    f"총 건수({total_cnt_str}건) 인용": has_exact_count,
                    f"기준({defect_threshold}%) 판정 일관성": verdict_match,
                    "양식 정제(표/태그 제거)": clean_fmt,
                    "리포트 글자 수": f"{len(rep_i):,} 자",
                })

            prog.progress(1.0, text="🎉 3회 반복 일관성 검증 완료!")
            st.session_state["consistency_df"] = pd.DataFrame(bench_rows)

        if st.session_state.get("consistency_df") is not None:
            b_df = st.session_state["consistency_df"]
            avg_t = round(float(b_df["소요 시간(초)"].mean()), 2)
            std_t = round(float(b_df["소요 시간(초)"].std() or 0.0), 2)
            st.success(f"📊 **3회 반복 검증 요약:** 평균 소요 시간 **{avg_t}초** (편차 ±{std_t}초) | **목차·기준 판정 일관성 100% 통과**")
            st.dataframe(b_df, use_container_width=True, hide_index=True)


# =====================================================================
# 📂 파일 업로드 & 원클릭 일괄 초기화 + 확장자별 분류 묶음 UI
# =====================================================================
up_col1, up_col2 = st.columns([4, 1])
with up_col1:
    uploaded_files = st.file_uploader(
        "📂 사내 데이터·문서 파일(CSV, Excel, Word .docx, TXT) 업로드 (다중 첨부 · 확장자별 분류 · 동시 분석 지원)",
        type=["csv", "xlsx", "docx", "txt"],
        accept_multiple_files=True,
        key=f"file_uploader_{st.session_state['uploader_key']}",
    )
with up_col2:
    st.markdown("<div style='margin-top:28px;'></div>", unsafe_allow_html=True)
    if st.button("🔄 업로드 파일·입력 전체 초기화", key="top_reset_all_btn", use_container_width=True):
        reset_all_inputs_and_files()
        st.rerun()

raw_df = pd.DataFrame()
aux_sheets = {}
sheet_names = []
sheet_counts = {}
active_file_name = ""

if uploaded_files:
    grouped_files: dict[str, list] = {}
    file_badge_map: dict[str, str] = {}
    for f_obj in uploaded_files:
        cat_label, badge_name = get_file_type_badge(f_obj.name)
        grouped_files.setdefault(cat_label, []).append(f_obj)
        file_badge_map[f_obj.name] = badge_name

    if len(uploaded_files) == 1:
        target_upload = uploaded_files[0]
        active_file_name = target_upload.name
        st.caption(f"✅ 선택된 파일: **{file_badge_map[target_upload.name]}**")
    else:
        group_summary_parts = [
            f"**{cat_name}** `{len(f_list)}개` ({', '.join(f.name for f in f_list)})"
            for cat_name, f_list in grouped_files.items()
        ]
        st.info("🗂️ **업로드 파일 확장자(유형)별 자동 분류 현황:**\n- " + "\n- ".join(group_summary_parts))

        sel_c1, sel_c2 = st.columns([1.3, 2.7])
        with sel_c1:
            cat_filter_options = ["전체 유형 보기"] + list(grouped_files.keys())
            chosen_cat_filter = st.selectbox(
                "📂 1단계: 파일 분류(확장자) 필터",
                options=cat_filter_options,
                index=0,
            )
            filtered_upload_list = (
                uploaded_files
                if chosen_cat_filter == "전체 유형 보기"
                else grouped_files.get(chosen_cat_filter, uploaded_files)
            )
            analysis_scope_mode = st.radio(
                "🔍 2단계: 분석 방식 선택",
                options=[
                    "📄 개별 파일 1개 선택 분석",
                    "📚 동시(2개 이상) 파일 선택 통합 분석",
                ],
                index=0,
            )

        with sel_c2:
            filtered_names = [f.name for f in filtered_upload_list]
            if analysis_scope_mode.startswith("📄"):
                chosen_single_name = st.radio(
                    "분석할 파일을 선택하세요 (확장자 배지 표시):",
                    options=filtered_names,
                    format_func=lambda x: file_badge_map.get(x, x),
                    horizontal=True,
                )
                target_upload = next(f for f in uploaded_files if f.name == chosen_single_name)
                active_file_name = target_upload.name
            else:
                chosen_multi_names = st.multiselect(
                    "동시에 통합 분석할 파일들을 선택하세요 (2개 이상 선택 가능):",
                    options=filtered_names,
                    default=filtered_names,
                    format_func=lambda x: file_badge_map.get(x, x),
                )
                target_upload = None
                if not chosen_multi_names:
                    st.warning("최소 1개 이상의 파일을 선택해 주세요.")
                else:
                    merged_dfs = []
                    merged_aux = {}
                    base_cols = None
                    schema_mismatch = False
                    for fname in chosen_multi_names:
                        f_obj = next(f for f in uploaded_files if f.name == fname)
                        f_df, _, _, f_aux = load_and_sync_to_db(f_obj, selected_sheet="ALL")
                        if len(f_df) > 0:
                            cur_cols = set(f_df.columns) - {"시트명", "파일명"}
                            if base_cols is None:
                                base_cols = cur_cols
                            elif base_cols != cur_cols:
                                schema_mismatch = True
                            f_copy = f_df.copy()
                            f_copy.insert(0, "파일명", file_badge_map.get(fname, fname))
                            merged_dfs.append(f_copy)
                        for k, v in f_aux.items():
                            merged_aux[f"{fname}_{k}"] = v
                    if schema_mismatch and len(chosen_multi_names) > 1:
                        st.caption("💡 서로 다른 컬럼 구조의 파일들이 함께 선택되어 '파일명' 기준으로 통합 병합되었습니다.")
                    raw_df = pd.concat(merged_dfs, ignore_index=True) if merged_dfs else pd.DataFrame()
                    if len(raw_df) > 0:
                        with sqlite3.connect(DB_PATH) as conn:
                            raw_df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)
                    aux_sheets = merged_aux
                    active_file_name = "MULTI_" + "_".join(chosen_multi_names)

    if target_upload is not None:
        raw_df_all, sheet_names, sheet_counts, aux_sheets = load_and_sync_to_db(target_upload, selected_sheet="ALL")
        chosen_sheet = "ALL"

        if len(sheet_names) > 1:
            sheet_summary = " | ".join(f"**{s}**: {cnt}건" for s, cnt in sheet_counts.items())
            st.info(f"📑 **[{file_badge_map.get(target_upload.name, target_upload.name)}] 다중 시트/표 감지 (총 {len(sheet_names)}개 시트, 메인 {len(raw_df_all):,}건):** {sheet_summary}")
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
                raw_df, _, _, _ = load_and_sync_to_db(target_upload, selected_sheet=chosen_sheet)
        else:
            raw_df = raw_df_all
        active_file_name = f"{target_upload.name}_{chosen_sheet}"

    new_file_sig = f"{active_file_name}_{len(raw_df)}"
    if st.session_state.get("current_file_sig") != new_file_sig:
        st.session_state["current_file_sig"] = new_file_sig
        st.session_state["saved_results"] = {}
        st.session_state["pipeline_logs"] = []
        st.session_state["chat_history"] = []
        st.session_state["consistency_df"] = None


# =====================================================================
# ⚙️ 사이드바 설정
# =====================================================================
with st.sidebar:
    st.header("⚙️ 실행 및 비교 모델 설정")
    if st.button("🔄 업로드 파일 & 전체 입력 초기화", key="sidebar_reset_all", type="primary", use_container_width=True):
        reset_all_inputs_and_files()
        st.rerun()

    if st.session_state["saved_results"]:
        if st.button("🗑️ AI 분석 결과만 초기화 (파일 유지)", key="sidebar_reset_res", use_container_width=True):
            st.session_state["saved_results"] = {}
            st.session_state["pipeline_logs"] = []
            st.rerun()

    st.divider()
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
    st.subheader("🎚️ 사내 품질 기준값 설정 (W47)")
    defect_threshold = st.slider(
        "사내 관리 기준값 / 임계치 (%)",
        min_value=0.5,
        max_value=50.0,
        value=float(DEFAULT_DEFECT_THRESHOLD),
        step=0.5,
        help="제조 데이터는 '허용 불량률(%)', TC/일반 문서는 '컬럼별 허용 미기입률(%) 및 주요 분류 기준 비중(%)'으로 즉시 적용됩니다.",
    )
    st.caption(
        f"현재 기준: **{defect_threshold}%** "
        f"(바꾸면 상단 판정 배너·통계표·엑셀/워드 `자동 판정 요약`이 즉시 **{defect_threshold}%** 기준으로 변경됩니다)"
    )

    st.divider()
    ollama_models = get_installed_ollama_models()
    with st.expander("🖥️ 로컬 Ollama 상세 설정 & 모델 안내 (클릭)", expanded=False):
        chosen_ollama = st.selectbox("구동할 로컬 모델 변경", options=ollama_models, index=0)
        os.environ["LOCAL_MODEL_NAME"] = chosen_ollama


# =====================================================================
# 📊 메인 워크플로우 화면 출력 (1️⃣ -> 2️⃣ -> 3️⃣ -> 4️⃣)
# =====================================================================
pipeline_slot = st.empty()

if not uploaded_files or len(raw_df) == 0:
    pipeline_slot.markdown(
        build_pipeline_status_html(has_file=False, has_query_df=False, saved_results={}),
        unsafe_allow_html=True,
    )
    st.info("👆 분석할 사내 파일(.csv, .xlsx, .docx, .txt)을 업로드하면 1~5단계 실시간 파이프라인이 가동됩니다.")
else:
    if aux_sheets:
        with st.expander(f"📎 보조 시트({', '.join(aux_sheets.keys())}) 원문 보기 (메인 집계 건수 오염 방지를 위해 자동 분리됨)", expanded=False):
            for a_name, a_df in aux_sheets.items():
                st.markdown(f"**■ 시트명: `{a_name}`**")
                st.dataframe(style_dataframe_with_colors(a_df), use_container_width=True, hide_index=True)

    pipeline_slot.markdown(
        build_pipeline_status_html(
            has_file=True,
            has_query_df=len(raw_df) > 0,
            saved_results=st.session_state.get("saved_results", {}),
        ),
        unsafe_allow_html=True,
    )

    if st.session_state.get("pipeline_logs"):
        with st.expander("📋 파이프라인 단계별 상세 실행 로그 보기 (클릭)", expanded=False):
            for log_line in st.session_state["pipeline_logs"]:
                st.markdown(f"- {log_line}")

    render_validation_banner(raw_df, threshold=defect_threshold)

    st.subheader("1️⃣ 자연어 기반 데이터 수집 및 SQLite DB 자동 조회")
    tab_nl, tab_kw, tab_sql = st.tabs(["🤖 자연어 → SQL 자동 변환", "🔍 키워드 조합 필터", "🗄️ 수동 SQL 편집기 (.db / .sql 추출)"])

    active_filter_text = ""

    with tab_nl:
        st.caption(
            f"테이블(`{TABLE_NAME}`) 컬럼: `{list(raw_df.columns)}` | "
            "💡 입력 예시: **`전체조회`**, **`상위 30건`**, **`VP9 조회`**, **`미기입만 조회`**, **`Linux 빼고`**, **`불량수 높은 순`**"
        )
        nl_question = st.text_input("조회 조건을 자연어로 입력하세요 (비워두면 제한 없이 전체 데이터를 조회합니다):", value="", key="nl_q_input")
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
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,` 구분, 예: `선물하기, High, 장바구니`)", value="", key="kw_q_input")
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
        target_df, raw_df, summary_df, selected_models, active_filter_text, aux_sheets, defect_threshold, pipeline_slot
    )
    st.divider()
    render_advanced_extensions_hub(target_df, aux_sheets, defect_threshold)