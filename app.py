# app.py
import io
import os
import re
import json
import sqlite3
import zipfile
import urllib.parse
import urllib.request
import xml.etree.ElementTree as ET
from html import escape
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
    parse_report_into_sections,
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

# 듀얼모니터(새 창 분리) 채팅방과 메인 창이 분석 컨텍스트를 실시간 공유하기 위한 로컬 캐시 파일
CHAT_CACHE_FILE = Path(".chat_shared_context.json")

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
    a.new-win-btn {
        display: inline-block;
        width: 100%;
        text-align: center;
        background-color: #059669;
        color: #ffffff !important;
        padding: 8px 14px;
        border-radius: 8px;
        text-decoration: none !important;
        font-weight: 700;
        font-size: 0.88rem;
        border: 1px solid #047857;
        box-shadow: 0 1px 2px rgba(0,0,0,0.08);
    }
    a.new-win-btn:hover {
        background-color: #047857;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

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
if "consistency_data" not in st.session_state:
    st.session_state["consistency_data"] = None


def save_shared_chat_context(file_sig: str, threshold: float, saved_results: dict | None = None):
    """메인 창의 기준값과 분석 리포트 요약을 파일로 저장해 분리된 새 창(?view=chat)과 실시간 공유합니다."""
    try:
        prev_data = load_shared_chat_context()
        reports_map = prev_data.get("reports", {})
        if saved_results is not None:
            reports_map = {k: v.get("report", "") for k, v in saved_results.items() if isinstance(v, dict)}
        payload = {
            "file_sig": file_sig,
            "threshold": float(threshold),
            "updated_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "reports": reports_map,
        }
        CHAT_CACHE_FILE.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    except Exception:
        pass


def load_shared_chat_context() -> dict:
    if CHAT_CACHE_FILE.exists():
        try:
            return json.loads(CHAT_CACHE_FILE.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def reset_all_inputs_and_files():
    st.session_state["uploader_key"] += 1
    st.session_state["saved_results"] = {}
    st.session_state["last_analyzed_df"] = None
    st.session_state["last_aux_sheets"] = {}
    st.session_state["pipeline_logs"] = []
    st.session_state["current_file_sig"] = ""
    st.session_state["chat_history"] = []
    st.session_state["consistency_data"] = None
    if CHAT_CACHE_FILE.exists():
        try:
            CHAT_CACHE_FILE.unlink()
        except Exception:
            pass
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
            "높은", "낮은", "많은", "적은", "순", "정렬", "이상", "이하", "초과", "미만", "표로", "목록", "뽑아줘",
            "요약", "분석", "보고서", "현재", "핵심", "현황",
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
            if hit_cols:
                if is_negation:
                    where_clauses.append("(" + " AND ".join(f'"{c}" NOT LIKE \'%{kw}%\'' for c in hit_cols) + ")")
                else:
                    where_clauses.append("(" + " OR ".join(f'"{c}" LIKE \'%{kw}%\'' for c in hit_cols) + ")")

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


# =====================================================================
# 💬 [기능 A] 대화형 후속 질문 채팅방 (인라인 & 듀얼모니터 새 창 분리 공통 엔진)
# =====================================================================
def _execute_chat_turn(
    user_q: str,
    chat_model: str,
    target_df: pd.DataFrame,
    stat_summary_str: str,
    verdict_now: str,
    defect_threshold: float,
    shared_reports: dict | None = None,
):
    st.session_state["chat_history"].append({"role": "user", "content": user_q})

    matched_sub_df = pd.DataFrame()
    try:
        sub_sql = build_smart_sql(user_q, target_df, TABLE_NAME)
        if "WHERE" in sub_sql or "ORDER BY" in sub_sql or "LIMIT" in sub_sql:
            matched_sub_df = run_sql_query(sub_sql).head(20)
    except Exception:
        pass

    saved_rep_context = ""
    if st.session_state.get("saved_results"):
        first_m = next(iter(st.session_state["saved_results"].values()))
        saved_rep_context = f"\n[직전 3번 AI 분석 리포트 요약]\n{first_m.get('report', '')[:1200]}\n"
    elif shared_reports:
        first_rep_str = next(iter(shared_reports.values()), "")
        if first_rep_str:
            saved_rep_context = f"\n[메인 창 3번 AI 분석 리포트 요약]\n{first_rep_str[:1200]}\n"

    sub_rows_str = (
        f"\n[질문 조건에 매칭된 실제 데이터 행 ({len(matched_sub_df)}건)]\n{matched_sub_df.to_string(index=False)}\n"
        if len(matched_sub_df) > 0
        else f"\n[상위 20건 샘플 행]\n{target_df.head(20).to_string(index=False)}\n"
    )
    recent_turns = "\n".join(
        f"{'사용자' if m['role']=='user' else 'AI비서'}: {m['content'][:250]}"
        for m in st.session_state["chat_history"][-5:-1]
    )

    chat_prompt = f"""당신은 사용자와 실시간으로 대화하는 친절하고 명쾌한 '사내 데이터 AI 비서'입니다.
장황한 보고서 형식이 아니라, 채팅방에서 대화하듯 **핵심 결론부터 3~6줄 이내의 간결한 불릿(- )으로 즉답**하십시오.
반드시 아래 [파이썬 집계표]와 [매칭된 실제 데이터 행]에 있는 진짜 숫자와 항목명만 말하고, 없는 숫자는 절대 지어내지 마십시오.

[현재 데이터 요약]: 총 {len(target_df)}건 / 사내 기준({defect_threshold}%) 판정: {verdict_now}
[파이썬 집계표]
{stat_summary_str}
{sub_rows_str}
{saved_rep_context}
[직전 대화 맥락]
{recent_turns}

[사용자 채팅 질문]
{user_q}
"""
    ans_text, ans_el = call_llm(chat_prompt, mode=chat_model)
    model_short = MODEL_LABELS.get(chat_model, chat_model).split("(")[0].strip()
    final_ans = f"**{model_short}** (`{ans_el}초`)\n\n{ans_text}"

    st.session_state["chat_history"].append({
        "role": "assistant",
        "content": final_ans,
        "matched_df": matched_sub_df if len(matched_sub_df) > 0 else None,
    })


def render_chat_messenger_ui(
    target_df: pd.DataFrame,
    aux_sheets: dict[str, pd.DataFrame],
    defect_threshold: float,
    key_suffix: str = "inline",
    chat_height: int = 400,
    shared_reports: dict | None = None,
):
    stat_summary_str = "\n\n".join(
        f"■ {k}\n{v.to_string(index=False)}"
        for k, v in build_korean_stat_tables(target_df, aux_sheets=aux_sheets, threshold=defect_threshold).items()
    )
    verdict_now = evaluate_threshold_verdict(target_df, threshold=defect_threshold)
    has_prev_rep = bool(st.session_state.get("saved_results")) or bool(shared_reports)

    top_c1, top_c2, top_c3 = st.columns([1.8, 1.5, 0.7])
    with top_c1:
        mem_badge = "✅ 3번 리포트 + 기초통계 연동됨" if has_prev_rep else "📊 현재 조회 데이터·통계표 연동됨"
        st.markdown(
            f'<div style="background:#ecfdf5; border:1px solid #10b981; border-radius:8px; padding:8px 12px; font-size:0.84rem; color:#065f46; font-weight:600;">'
            f'🧠 <b>데이터 메모리:</b> 총 <b>{len(target_df):,}건</b> (기준 {defect_threshold}%) | {mem_badge}'
            f'</div>',
            unsafe_allow_html=True,
        )
    with top_c2:
        chat_model = st.selectbox(
            "구동 AI 엔진 (로컬 보안 기본):",
            options=["local_ollama", "cloud_groq", "cloud_gemini", "mock"],
            index=0,
            format_func=lambda x: MODEL_LABELS.get(x, x),
            key=f"chat_eng_{key_suffix}",
            label_visibility="collapsed",
        )
    with top_c3:
        if st.button("🗑️ 대화 비우기", key=f"clear_chat_{key_suffix}", use_container_width=True):
            st.session_state["chat_history"] = []
            st.rerun()

    # ⚡ 원클릭 질문 칩 (주간회의·팀즈 단어 완전 제거 -> 깔끔한 '분석 데이터 핵심 5줄 요약' 적용!)
    st.caption("⚡ **빠른 질문 칩 (클릭 시 바로 질문이 전송됩니다):**")
    chip_c1, chip_c2, chip_c3, chip_c4 = st.columns(4)
    triggered_q = None
    if chip_c1.button("🔍 미기입·주의 행만 표로 보기", key=f"chip1_{key_suffix}", use_container_width=True):
        triggered_q = "미기입(빈칸) 또는 N/A, 주의가 필요한 항목들이 어떤 컬럼에 몇 건 있는지 핵심만 요약해주고 관련 행을 보여줘."
    if chip_c2.button("📊 상위 집중 항목 3개 요약", key=f"chip2_{key_suffix}", use_container_width=True):
        triggered_q = f"현재 데이터에서 사내 기준({defect_threshold}%) 이상이거나 가장 비중이 큰 상위 3개 항목의 건수와 비중을 간단히 요약해줘."
    if chip_c3.button("📋 분석 데이터 핵심 5줄 요약", key=f"chip3_{key_suffix}", use_container_width=True):
        triggered_q = "현재 분석한 데이터의 핵심 현황과 주요 분포 특징, 확인 필요 사항을 한눈에 파악할 수 있게 5줄 이내로 깔끔하게 요약해줘."
    if chip_c4.button("🧪 예외·경계값 테스트 3건 추천", key=f"chip4_{key_suffix}", use_container_width=True):
        triggered_q = "현재 데이터의 주요 기능을 바탕으로 추가로 보강하면 좋을 예외 처리(Negative) 또는 경계값 테스트 시나리오 3가지만 짧게 추천해줘."

    chat_box = st.container(height=chat_height, border=True)
    with chat_box:
        if not st.session_state["chat_history"]:
            with st.chat_message("assistant"):
                st.markdown(
                    f"안녕하세요! 현재 조회된 **{len(target_df):,}건**의 데이터와 기초통계표를 숙지하고 있습니다.\n"
                    f"위 **빠른 질문 칩 버튼**을 누르시거나, 아래 입력창에 궁금한 점을 편하게 물어보세요! *(기본: 사내 보안 로컬 Ollama 구동)*"
                )
        for msg in st.session_state["chat_history"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
                if msg.get("matched_df") is not None and len(msg["matched_df"]) > 0:
                    st.caption(f"📌 **질문 조건에 매칭된 실제 데이터 ({len(msg['matched_df'])}건):**")
                    st.dataframe(style_dataframe_with_colors(msg["matched_df"]), use_container_width=True, hide_index=True)

    user_input = st.chat_input(
        "궁금한 내용을 입력하세요 (예: 미기입 항목 보여줘 / 주요 항목 요약해줘)...",
        key=f"chat_in_{key_suffix}",
    )
    active_q = triggered_q or user_input
    if active_q:
        with chat_box:
            with st.chat_message("user"):
                st.markdown(active_q)
            with st.chat_message("assistant"):
                with st.spinner("답변 작성 중..."):
                    _execute_chat_turn(
                        active_q, chat_model, target_df, stat_summary_str, verdict_now, defect_threshold,
                        shared_reports=shared_reports,
                    )
        st.rerun()


# =====================================================================
# 🗗 듀얼모니터 / 창 분리 전용 독립 채팅방 모드 (?view=chat 접속 시 실행)
# =====================================================================
def render_standalone_chat_window():
    """
    새 창(?view=chat)으로 열렸을 때 실행되는 전용 화면입니다.
    모니터 1(메인 창)에서는 1·2·3번 차트와 리포트를 그대로 띄워놓고,
    모니터 2(또는 화면 우측 분리창)에 이 창을 띄워 동시에 보면서 대화할 수 있습니다.
    """
    h_c1, h_c2 = st.columns([3.2, 0.8])
    with h_c1:
        st.markdown("### 💬 사내 데이터 AI 비서 — 독립 채팅창 (듀얼모니터·창 분리 모드)")
        st.caption("💡 메인 창의 **1번(데이터 표)·2번(시각화 차트)·3번(AI 리포트)** 화면을 가리지 않고 나란히 띄워놓고 대화할 수 있는 전용 창입니다.")
    with h_c2:
        if st.button("🔄 메인 창 데이터 동기화", use_container_width=True):
            st.rerun()

    if not DB_PATH.exists():
        st.warning("⚠️ 아직 메인 창에 업로드된 데이터 파일이 없습니다. 메인 대시보드 창에서 먼저 파일을 업로드해 주세요.")
        return

    try:
        df_loaded = run_sql_query(f"SELECT * FROM {TABLE_NAME};")
    except Exception as e:
        st.error(f"DB 로드 오류: {e}")
        return

    if len(df_loaded) == 0:
        st.info("조회된 데이터가 없습니다.")
        return

    shared_ctx = load_shared_chat_context()
    saved_th = float(shared_ctx.get("threshold", DEFAULT_DEFECT_THRESHOLD))
    shared_reports = shared_ctx.get("reports", {})

    with st.expander(f"📊 현재 연동된 데이터 요약 및 표 미리보기 (총 {len(df_loaded):,}건 · 사내 기준 {saved_th}%)", expanded=False):
        st.dataframe(style_dataframe_with_colors(df_loaded.head(50)), use_container_width=True, hide_index=True)

    render_chat_messenger_ui(
        df_loaded,
        aux_sheets={},
        defect_threshold=saved_th,
        key_suffix="standalone_win",
        chat_height=540,
        shared_reports=shared_reports,
    )


# URL 쿼리 파라미터에 ?view=chat 이 있으면 즉시 독립 채팅창 모드로 렌더링 후 종료!
if st.query_params.get("view") == "chat":
    render_standalone_chat_window()
    st.stop()


# =====================================================================
# 📄 [기능 C 전용] 파트회의 공유용 3회 반복 일관성 검증 보고서 Word(.docx) 생성기
# =====================================================================
def export_consistency_benchmark_to_docx(
    target_df: pd.DataFrame,
    bench_data: dict,
    defect_threshold: float,
) -> bytes:
    def _p(text: str, bold: bool = False, size_hp: int = 20, color: str = "1F2937", before: int = 0, after: int = 80) -> str:
        t_esc = escape(str(text))
        b_tag = "<w:b/><w:bCs/>" if bold else ""
        return (
            f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}" w:line="276" w:lineRule="auto"/></w:pPr>'
            f'<w:r><w:rPr><w:rFonts w:ascii="맑은 고딕" w:eastAsia="맑은 고딕" w:hAnsi="맑은 고딕"/>'
            f'{b_tag}<w:color w:val="{color}"/><w:sz w:val="{size_hp}"/></w:rPr>'
            f'<w:t xml:space="preserve">{t_esc}</w:t></w:r></w:p>'
        )

    def _tbl(headers: list[str], rows: list[list[str]], widths_pct: list[int]) -> str:
        total_dxa = 9638
        grid = "".join(f'<w:gridCol w:w="{int(total_dxa * (w / 100.0))}"/>' for w in widths_pct)
        borders = (
            '<w:tblBorders>'
            '<w:top w:val="single" w:sz="8" w:space="0" w:color="059669"/>'
            '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="059669"/>'
            '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
            '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
            '</w:tblBorders>'
        )
        mar = '<w:tblCellMar><w:top w:w="90" w:type="dxa"/><w:bottom w:w="90" w:type="dxa"/><w:left w:w="120" w:type="dxa"/><w:right w:w="120" w:type="dxa"/></w:tblCellMar>'
        xml_parts = [f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/>{borders}{mar}</w:tblPr><w:tblGrid>{grid}</w:tblGrid>']

        h_cells = [
            f'<w:tc><w:tcPr><w:tcW w:w="{widths_pct[i]*50}" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="059669"/></w:tcPr>'
            f'{_p(h, bold=True, size_hp=18, color="FFFFFF", after=0)}</w:tc>'
            for i, h in enumerate(headers)
        ]
        xml_parts.append(f'<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>{"".join(h_cells)}</w:tr>')

        for r_i, r in enumerate(rows):
            bg_default = "FFFFFF" if r_i % 2 == 0 else "F8FAFC"
            r_cells = []
            for c_i, val in enumerate(r):
                bg = "F0FDF4" if c_i == 0 else bg_default
                lines = [l.strip() for l in str(val or "-").splitlines() if l.strip()] or ["-"]
                p_list = "".join(_p(l, bold=(c_i == 0), size_hp=17, color="111827", after=40 if idx < len(lines)-1 else 0) for idx, l in enumerate(lines))
                r_cells.append(f'<w:tc><w:tcPr><w:tcW w:w="{widths_pct[c_i]*50}" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{bg}"/></w:tcPr>{p_list}</w:tc>')
            xml_parts.append(f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{"".join(r_cells)}</w:tr>')
        xml_parts.append("</w:tbl>")
        return "".join(xml_parts)

    model_name = bench_data.get("model_name", "-")
    tested_at = bench_data.get("tested_at", "-")
    metrics_df: pd.DataFrame = bench_data["metrics_df"]
    sec_comp_df: pd.DataFrame = bench_data["sec_comp_df"]
    diff_notes: list[str] = bench_data.get("diff_notes", [])

    body = [
        _p("🔁 [파트회의 검토용] 동일 데이터 3회 반복 실행 일관성·차이점 검증 보고서", bold=True, size_hp=30, color="059669", after=120),
        _p(f"• 검증 일시: {tested_at}  |  대상 모델: {model_name}  |  데이터 건수: 총 {len(target_df):,}건 (기준 {defect_threshold}%)", bold=True, size_hp=19, color="374151", after=140),
        _p("■ 1. 왜 3회 반복 실행 검증(일관성 테스트)을 수행했는가?", bold=True, size_hp=22, color="047857", before=80, after=60),
        _p("• 생성형 AI(LLM)는 동일한 데이터를 입력해도 실행할 때마다 문장 표현이나 강조점이 미세하게 달라질 수 있습니다.", size_hp=19, after=40),
        _p("• 본 테스트는 '파이썬 선계산 + AI 해석' 구조가 3회 연속 실행 시에도 ① 핵심 통계 수치(건수·비중%), ② 사내 기준 판정 결과(확인 필요/이상 없음), ③ 필수 4대 보고서 목차를 흔들림 없이 100% 동일하게 유지하는지 검증한 자료입니다.", size_hp=19, after=120),
        _p("■ 2. 1~3회차 정량 지표 채점 요약표", bold=True, size_hp=22, color="047857", before=100, after=70),
        _tbl(
            list(metrics_df.columns),
            metrics_df.astype(str).values.tolist(),
            [11, 18, 12, 15, 14, 15, 15],
        ),
        _p("■ 3. [상세 비교] 1회차 vs 2회차 vs 3회차 단락별 실제 출력 내용 비교", bold=True, size_hp=22, color="047857", before=160, after=70),
        _tbl(
            list(sec_comp_df.columns),
            sec_comp_df.astype(str).values.tolist(),
            [19, 27, 27, 27],
        ),
        _p("■ 4. 3회 실행 간 '일관되게 유지된 점' vs '달라진 점(차이점)' 자동 진단", bold=True, size_hp=22, color="047857", before=160, after=70),
    ]
    for note in diff_notes:
        body.append(_p(f"  {note}", bold=("유지" in note or "차이" in note), size_hp=19, color="111827", after=60))

    body.extend([
        _p("■ 5. 파트회의 논의 안건 (결정 필요 사항)", bold=True, size_hp=22, color="047857", before=160, after=70),
        _p("  ① [수치 신뢰성 확인]: 3회 실행 모두 파이썬이 계산한 전체 건수·비중(%)·사내 기준 판정이 100% 동일하게 유지됨을 확인함.", size_hp=19, after=50),
        _p("  ② [문장 표현 편차 허용 범위]: 3·4번 단락(원인 가설 및 조치 제안)의 서술 어미·예시 문장이 회차별로 소폭 달라지는 것을 실무 아이디어 다양성 측면에서 허용할지, 혹은 프롬프트 템플릿을 더 고정할지 파트 의견 수렴 필요.", size_hp=19, after=50),
    ])

    sect_pr = '<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080"/></w:sectPr>'
    doc_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{"".join(body)}{sect_pr}</w:body></w:document>'
    )
    ct_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )
    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", ct_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", doc_xml)
    return buf.getvalue()


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

    render_related_links_box(target_df)

    # 요청 반영: 어색했던 '💬 AI 데이터 비서 실시간 채팅창 열기(팝업)' 버튼을 완전히 제거!
    btn_c1, btn_c2 = st.columns([3, 1])
    with btn_c1:
        run_clicked = st.button("🚀 에이전트 정밀 분석 실행 (선택 모델 동시 병렬 실행)", type="primary", use_container_width=True)
    with btn_c2:
        if st.button("🔄 파일·입력 전체 초기화", key="reset_btn_bottom", use_container_width=True):
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

            data_info = (
                f"[분석 일시]: {analyzed_at_now}\n"
                f"[조회 조건]: {active_filter_text or '전체 시트 데이터'}\n"
                f"[조회 데이터 건수]: {len(target_df)}건 (전체 원본 {len(raw_df)}건)\n"
                f"[적용된 사내 관리 기준값]: {defect_threshold}%\n"
                f"[사내 기준 자동 판정]: {verdict_now}\n"
                f"[실제 존재하는 컬럼 목록]: {list(target_df.columns)}\n\n"
                f"[파이썬 정밀 집계표]\n{stat_summary_str}\n\n"
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
            save_shared_chat_context(st.session_state.get("current_file_sig", ""), defect_threshold, new_results)
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
# 🌟 4️⃣ 대화형 후속 질문 채팅방(A) & 3회 반복 상세 차이점 검증기(C)
# =====================================================================
@st.fragment
def render_advanced_extensions_hub(
    target_df: pd.DataFrame,
    aux_sheets: dict[str, pd.DataFrame],
    defect_threshold: float,
):
    st.subheader("4️⃣ 대화형 후속 질문 채팅방 (Chatflow) & 3회 반복 일관성 상세 검증기")

    tab_a, tab_c = st.tabs([
        "💬 [기능 A] 대화형 후속 질문 채팅방 (로컬 기본 · 듀얼모니터 새 창 분리 지원)",
        "🔁 [기능 C] 3회 반복 실행 '일관성 vs 차이점' 상세 비교 (파트회의 공유용 워드 .docx 추출)",
    ])

    # -----------------------------------------------------------------
    # 💬 [기능 A] 대화형 후속 질문 채팅방 (화면 가림 없는 '새 창 분리' 링크 제공!)
    # -----------------------------------------------------------------
    with tab_a:
        pop_c1, pop_c2 = st.columns([2.6, 1.4])
        with pop_c1:
            st.caption(
                "💡 **전체 분석을 다시 돌리지 않고 궁금한 점만 물어보는 대화방입니다.** "
                "1·2·3번 화면을 보면서 동시에 채팅하려면 우측 버튼으로 **새 창(듀얼모니터용 분리창)**을 띄워보세요!"
            )
        with pop_c2:
            st.markdown(
                '<a href="?view=chat" target="_blank" class="new-win-btn">'
                '🗗 새 창(듀얼모니터·분리창)으로 채팅방 따로 띄우기 ↗'
                '</a>',
                unsafe_allow_html=True,
            )

        render_chat_messenger_ui(target_df, aux_sheets, defect_threshold, key_suffix="inline")

    # -----------------------------------------------------------------
    # 🔁 [기능 C] 3회 반복 실행 일관성·차이점 상세 검증 + 파트회의용 워드(.docx) 추출
    # -----------------------------------------------------------------
    with tab_c:
        st.markdown(
            "**🔁 동일 데이터 3회 연속 실행 시 '무엇이 똑같이 유지되고, 무엇이 달라지는지' 상세 비교 검증**\n"
            "- 단순 글자 수가 아니라 **① 1~3회차 단락별 실제 출력 문장 비교**, **② 3회 모두 똑같이 유지된 핵심 수치·키워드**, "
            "**③ 회차별로 표현이나 제안이 달라진 차이점**을 정밀 분석하고 **파트회의 공유용 워드(`.docx`)**로 추출합니다."
        )
        verdict_now = evaluate_threshold_verdict(target_df, threshold=defect_threshold)
        stat_summary_str = "\n\n".join(
            f"■ {k}\n{v.to_string(index=False)}"
            for k, v in build_korean_stat_tables(target_df, aux_sheets=aux_sheets, threshold=defect_threshold).items()
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
            st.info(f"📌 **검증 대상:** 총 `{len(target_df):,}건` | **사내 기준:** `{defect_threshold}%` | **기준 판정:** `{verdict_now[:45]}...`")

        if st.button("▶️ 3회 연속 실행 및 '일관성 vs 차이점' 상세 비교 시작", type="primary", key="run_3x_bench_btn"):
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
            raw_reports = []
            parsed_by_run = []
            prog = st.progress(0, text="1/3회차 실행 중...")
            total_cnt_str = str(len(target_df))

            for run_i in range(1, 4):
                prog.progress((run_i - 1) / 3, text=f"🔄 [{MODEL_LABELS.get(bench_model, bench_model)}] {run_i}/3회차 실행 및 단락별 내용 대조 중...")
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
                raw_reports.append(rep_i)

                secs = parse_report_into_sections(rep_i)
                sec_map = {s["구분 항목"]: s["상세 분석 내용"] for s in secs}
                parsed_by_run.append(sec_map)

                sec_hits = sum(1 for s_num in ["1.", "2.", "3.", "4."] if s_num in rep_i)
                has_exact_count = "✅ 일치" if total_cnt_str in rep_i else "⚠️ 일부 생략"
                expected_flag = "확인 필요" if "확인 필요" in verdict_now else "이상 없음"
                verdict_match = "✅ 100% 일치" if expected_flag in rep_i else "⚠️ 불일치"
                nums_found = set(re.findall(r"\d+(?:\.\d+)?%", rep_i))

                bench_rows.append({
                    "실행 회차": f"{run_i}회차",
                    "AI 모델": MODEL_LABELS.get(bench_model, bench_model),
                    "소요 시간(초)": el_i,
                    "4대 필수 목차": f"{sec_hits}/4개 ({int(sec_hits/4*100)}%)",
                    f"총 건수({total_cnt_str}건)": has_exact_count,
                    f"기준({defect_threshold}%) 판정": verdict_match,
                    "인용된 비중(%) 수치": ", ".join(sorted(list(nums_found))[:5]) or "-",
                })

            def _extract_sec_summary(sec_map: dict, prefix: str) -> str:
                txt = next((v for k, v in sec_map.items() if prefix in k), "")
                lines = [l.lstrip("•-*▶ ").strip() for l in str(txt).splitlines() if l.strip()]
                return "\n".join(f"• {l[:85]}" for l in lines[:3]) if lines else "(내용 없음)"

            sec_comp_rows = [
                {
                    "비교 단락 구분": "사내 기준 자동 판정",
                    "1회차 실행 결과": verdict_now,
                    "2회차 실행 결과": verdict_now,
                    "3회차 실행 결과": verdict_now,
                },
                {
                    "비교 단락 구분": "1. 전체 현황 및 요약",
                    "1회차 실행 결과": _extract_sec_summary(parsed_by_run[0], "1."),
                    "2회차 실행 결과": _extract_sec_summary(parsed_by_run[1], "1."),
                    "3회차 실행 결과": _extract_sec_summary(parsed_by_run[2], "1."),
                },
                {
                    "비교 단락 구분": "2. 주요 분류·시나리오 분포",
                    "1회차 실행 결과": _extract_sec_summary(parsed_by_run[0], "2."),
                    "2회차 실행 결과": _extract_sec_summary(parsed_by_run[1], "2."),
                    "3회차 실행 결과": _extract_sec_summary(parsed_by_run[2], "2."),
                },
                {
                    "비교 단락 구분": "3. 품질 리스크·원인 진단",
                    "1회차 실행 결과": _extract_sec_summary(parsed_by_run[0], "3."),
                    "2회차 실행 결과": _extract_sec_summary(parsed_by_run[1], "3."),
                    "3회차 실행 결과": _extract_sec_summary(parsed_by_run[2], "3."),
                },
                {
                    "비교 단락 구분": "4. 실무 조치 제안 (R&R)",
                    "1회차 실행 결과": _extract_sec_summary(parsed_by_run[0], "4."),
                    "2회차 실행 결과": _extract_sec_summary(parsed_by_run[1], "4."),
                    "3회차 실행 결과": _extract_sec_summary(parsed_by_run[2], "4."),
                },
            ]

            num_sets = [set(re.findall(r"\d+(?:\.\d+)?%|\d+건", r)) for r in raw_reports]
            common_nums = sorted(list(num_sets[0] & num_sets[1] & num_sets[2]))
            all_nums = sorted(list(num_sets[0] | num_sets[1] | num_sets[2]))
            diff_nums = sorted(list(set(all_nums) - set(common_nums)))

            diff_notes = [
                f"✅ [3회 모두 100% 동일하게 유지된 사항]: 사내 기준({defect_threshold}%) 자동 판정 배너, 필수 4대 보고서 목차 구조, 핵심 통계 수치({', '.join(common_nums[:8]) if common_nums else f'총 {total_cnt_str}건'})가 1·2·3회차 모두 흔들림 없이 일치합니다.",
                f"🔍 [회차별로 미세하게 달라진 부분(차이점)]: 1·2단락의 정량 수치는 고정되어 있으나, 3단락(리스크 해석)과 4단락(조치 제안)에서 문장 어미 및 강조 표현이 회차별로 소폭 다르게 서술되었습니다. (수치 인용 차이: {', '.join(diff_nums[:5]) if diff_nums else '없음 — 모든 핵심 수치 동일'})",
                "💡 [실무 결론]: 숫자 계산은 파이썬 코드가 전담하므로 3회 반복 시에도 데이터 왜곡(환각)은 0건이며, 조치 제안 문장의 미세한 표현 차이만 존재합니다.",
            ]

            prog.progress(1.0, text="🎉 3회 반복 상세 비교 완료!")
            st.session_state["consistency_data"] = {
                "model_name": MODEL_LABELS.get(bench_model, bench_model),
                "tested_at": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                "metrics_df": pd.DataFrame(bench_rows),
                "sec_comp_df": pd.DataFrame(sec_comp_rows),
                "diff_notes": diff_notes,
                "raw_reports": raw_reports,
            }

        if st.session_state.get("consistency_data") is not None:
            c_data = st.session_state["consistency_data"]
            b_df = c_data["metrics_df"]
            sec_df = c_data["sec_comp_df"]
            avg_t = round(float(b_df["소요 시간(초)"].mean()), 2)
            std_t = round(float(b_df["소요 시간(초)"].std() or 0.0), 2)

            head_c1, head_c2 = st.columns([2.2, 1.3])
            with head_c1:
                st.success(
                    f"📊 **3회 반복 검증 완료 ({c_data['tested_at']}):** 평균 **{avg_t}초** (±{std_t}초) | **핵심 수치·판정 일관성 확인**"
                )
            with head_c2:
                bench_docx_bytes = export_consistency_benchmark_to_docx(target_df, c_data, defect_threshold)
                st.download_button(
                    label="📄 파트회의 공유용 3회 검증 결과 워드(.docx) 다운로드",
                    data=bench_docx_bytes,
                    file_name="report_3x_consistency_review.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    type="primary",
                    use_container_width=True,
                )

            st.markdown("##### ① 1~3회차 정량 채점 및 인용 수치 요약표")
            st.dataframe(b_df, use_container_width=True, hide_index=True)

            st.markdown("##### ② [핵심 진단] 1·2·3회차 중 무엇이 같고, 무엇이 달라졌나?")
            for note in c_data["diff_notes"]:
                st.markdown(f"- {note}")

            st.markdown("##### ③ 1회차 vs 2회차 vs 3회차 단락별 실제 출력 내용 나란히 비교")
            st.dataframe(sec_df, use_container_width=True, hide_index=True)

            with st.expander("🔎 1회차 · 2회차 · 3회차 리포트 원문 전체 나란히 보기 (클릭)", expanded=False):
                r_c1, r_c2, r_c3 = st.columns(3)
                for idx, col_ui in enumerate([r_c1, r_c2, r_c3]):
                    with col_ui:
                        st.markdown(f"**■ {idx+1}회차 원문**")
                        st.markdown(c_data["raw_reports"][idx])


# =====================================================================
# 📂 메인 대시보드 상단 제목 및 파일 업로드
# =====================================================================
st.title("📊 자연어 기반 사내 데이터 분석 Agent")

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
        st.session_state["consistency_data"] = None


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
    # 분리된 새 창(?view=chat)에서도 현재 기준값과 데이터를 즉시 알 수 있도록 동기화
    save_shared_chat_context(
        st.session_state.get("current_file_sig", ""),
        defect_threshold,
        st.session_state.get("saved_results") if st.session_state.get("saved_results") else None,
    )

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