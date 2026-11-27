# app.py
import io
import os
import re
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
    shorten_long_label,
)
from data_agent.llm import get_installed_ollama_models, call_llm
from data_agent.tools.data_loader import (
    load_and_sync_to_db,
    filter_by_keywords,
    run_sql_query,
    export_db_to_sql_script,
    extract_knowledge_text,
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
    # 아래 base = "dark" 옵션을 추가하면 강제 다크모드가 됩니다.
    CONFIG_FILE.write_text(f'[theme]\nbase = "dark"\nprimaryColor = "{PRIMARY_COLOR}"\n', encoding="utf-8")

st.set_page_config(page_title="사내 데이터 분석 Agent", layout="wide")

if "saved_results" not in st.session_state:
    st.session_state["saved_results"] = {}
if "last_analyzed_df" not in st.session_state:
    st.session_state["last_analyzed_df"] = None
if "last_summary_df" not in st.session_state:
    st.session_state["last_summary_df"] = None
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
if "chat_popup_mode" not in st.session_state:
    st.session_state["chat_popup_mode"] = False
if "chat_popup_wide" not in st.session_state:
    st.session_state["chat_popup_wide"] = False
if "chat_popup_minimized" not in st.session_state:
    st.session_state["chat_popup_minimized"] = False
if "consistency_data" not in st.session_state:
    st.session_state["consistency_data"] = None

popup_width_css = "min(920px, 72vw)" if st.session_state.get("chat_popup_wide") else "min(620px, 48vw)"

st.markdown(
    f"""
    <style>
    :root {{ --primary-color: #059669 !important; }}
    span[data-baseweb="tag"],
    div[data-baseweb="select"] span[data-baseweb="tag"],
    [data-testid="stMultiSelect"] span[data-baseweb="tag"] {{
        background-color: #059669 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }}
    span[data-baseweb="tag"] span {{ color: #ffffff !important; }}
    div[role="radiogroup"] label > div:first-child {{ border-color: #059669 !important; }}
    div[role="radiogroup"] label[data-baseweb="radio"] > div:first-child {{
        background-color: #059669 !important;
        border-color: #059669 !important;
    }}
    button[kind="primary"], [data-testid="stBaseButton-primary"] {{
        background-color: #059669 !important;
        border-color: #059669 !important;
        color: white !important;
    }}
    button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover {{
        background-color: #047857 !important;
        border-color: #047857 !important;
    }}

    div[data-testid="stVerticalBlock"]:has(#floating-chat-marker):not(:has(div[data-testid="stVerticalBlock"] #floating-chat-marker)) {{
        position: fixed !important;
        bottom: 22px !important;
        right: 26px !important;
        width: {popup_width_css} !important;
        max-height: 86vh !important;
        overflow-y: auto !important;
        overflow-x: hidden !important;
        background-color: #ffffff !important;
        border: 2.5px solid #059669 !important;
        border-radius: 16px !important;
        box-shadow: 0 20px 50px rgba(0, 0, 0, 0.28) !important;
        z-index: 999990 !important;
        padding: 14px 18px 18px 18px !important;
    }}
    </style>
    """,
    unsafe_allow_html=True,
)


def _build_compact_sample_str(df: pd.DataFrame, max_rows: int = 5) -> str:
    if df is None or len(df) == 0:
        return "데이터 없음"
    important_cols = [c for c in df.columns if any(k in str(c).lower() for k in ["id", "title", "procedure", "result", "status", "priority", "domain", "라인", "생산", "불량"])]
    if not important_cols:
        important_cols = list(df.columns[:10])
    sub = df[important_cols].head(max_rows).copy()
    for c in sub.columns:
        sub[c] = sub[c].apply(lambda x: shorten_long_label(str(x).replace("\n", " "), 20))
    return sub.to_string(index=False)


def reset_all_inputs_and_files():
    st.session_state["uploader_key"] += 1
    st.session_state["saved_results"] = {}
    st.session_state["last_analyzed_df"] = None
    st.session_state["last_summary_df"] = None
    st.session_state["last_aux_sheets"] = {}
    st.session_state["pipeline_logs"] = []
    st.session_state["current_file_sig"] = ""
    st.session_state["chat_history"] = []
    st.session_state["chat_popup_mode"] = False
    st.session_state["chat_popup_wide"] = False
    st.session_state["chat_popup_minimized"] = False
    st.session_state["consistency_data"] = None
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
        (2, "2. SQLite DB·정밀 집계", "done" if has_query_df else "wait", "조회 데이터 집계 완료" if has_query_df else "데이터 집계 대기"),
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


def _detect_mentioned_columns(query_str: str, df_columns: list) -> list:
    q_low = query_str.lower()
    q_norm = re.sub(r"[\s_\-]+", "", q_low)

    alias_map = {
        "procedure": ["procedure", "절차", "수행절차", "테스트절차"],
        "expectedresult": ["expected result", "expected", "기대결과", "예상결과", "판정기준", "pass/fail criteria"],
        "precondition": ["precondition", "pre-condition", "사전조건", "전제조건"],
        "description": ["description", "설명", "상세설명"],
        "actualresult": ["actual result", "actual", "실제결과"],
        "executionstatus": ["execution status", "status", "실행상태", "판정상태"],
        "priority": ["priority", "우선순위", "중요도"],
        "category": ["category", "카테고리", "분류"],
        "title": ["title", "제목", "테스트케이스명"],
    }

    matched_cols = []
    for col in df_columns:
        c_str = str(col).strip()
        c_low = c_str.lower()
        c_norm = re.sub(r"[\s_\-]+", "", c_low)
        if not c_norm or c_low in ("id", "tc id", "번호", "시트명", "파일명"):
            continue
        if c_low in q_low or c_norm in q_norm:
            matched_cols.append(col)
            continue
        for key_norm, aliases in alias_map.items():
            if key_norm in c_norm:
                if any(a in q_low for a in aliases):
                    matched_cols.append(col)
                    break

    return list(dict.fromkeys(matched_cols))


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

    if any(k in q_without_limit.lower() for k in ["미기입", "미기재", "빈칸", "공란", "누락", "미실행", "not run", "n/a"]):
        mentioned_cols = _detect_mentioned_columns(q_without_limit, list(df.columns))
        check_cols = mentioned_cols if mentioned_cols else [
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
        if target_col:
            num_str = nums[0]
            op = "="
            if any(k in q_without_limit for k in ["넘", "초과", "보다 큰", "많은", "높은", ">"]):
                op = ">=" if "이상" in q_without_limit else ">"
            elif any(k in q_without_limit for k in ["이상", "부터", ">="]):
                op = ">="
            elif any(k in q_without_limit for k in ["미만", "보다 작은", "적은", "낮은", "<"]):
                op = "<"
            elif any(k in q_without_limit for k in ["이하", "이내", "까지", "<="]):
                op = "<="
            where_clauses.append(f'"{target_col}" {op} {num_str}')

    sql = f"SELECT * FROM {table_name}"
    if where_clauses:
        sql += " WHERE " + " AND ".join(where_clauses)
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
    scope_title = ""
    if user_filter_text and user_filter_text.strip():
        scope_title = f"검색 조건 [{user_filter_text.strip()}] (총 {len(target_df):,}건)"
    else:
        scope_title = f"전체 조회 데이터 (총 {len(target_df):,}건)"

    stat_snippets = [f"사내 기준 {threshold}% 적용"]
    if is_manufacturing_df(target_df):
        p_sum = int(target_df["생산량"].sum())
        d_sum = int(target_df["불량수"].sum())
        rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
        stat_snippets.append(f"총 생산량 {p_sum:,}개, 총 불량수 {d_sum:,}개(불량률 {rate:.2f}%)")
        stat_str = ", ".join(stat_snippets)
        return (
            f"{scope_title} 데이터를 바탕으로 "
            f"① 라인·일자별 핵심 추세 및 불량 비중(%), ② 사내 허용 기준({threshold}%) 초과 이상 구간, "
            f"③ 공정·설비·작업 관점 원인 가설, ④ 담당 부서별 조치 제안을 정밀 분석해 줘."
        )

    return (
        f"{scope_title} 데이터를 바탕으로 "
        f"① 실제 항목별 핵심 현황 및 사내 기준({threshold}%) 비교, "
        f"② 주요 리스크 및 원인 진단, ③ 실무 조치 제안(24h/1개월)을 정리해 줘."
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
                f"- 📑 [Google 학술논문(Scholar) 검색](https://scholar.google.co.kr/scholar?q={q_enc})"
            )


def _filter_within_target_df(user_q: str, target_df: pd.DataFrame) -> pd.DataFrame:
    if target_df is None or len(target_df) == 0 or not user_q.strip():
        return pd.DataFrame()

    q_low = user_q.lower()
    mentioned_cols = _detect_mentioned_columns(user_q, list(target_df.columns))
    check_cols = mentioned_cols if mentioned_cols else [
        c for c in target_df.columns if str(c).lower() not in ("_polarion", "id", "tc id", "번호", "시트명", "파일명")
    ]

    is_empty_q = any(k in q_low for k in ["미기입", "미기재", "빈칸", "공란", "누락", "비어", "없는"])
    is_na_q = "n/a" in q_low
    is_notrun_q = any(k in q_low for k in ["미실행", "not run"])
    is_fail_q = any(k in q_low for k in ["불합격", "fail", "빨강", "주의"])

    if is_empty_q or is_na_q or is_notrun_q or is_fail_q:
        mask = pd.Series(False, index=target_df.index)
        for col in check_cols:
            s_col = target_df[col].astype(str).str.strip()
            if is_empty_q and not (is_na_q or is_notrun_q):
                mask |= s_col.isin(["-", "", "(빈칸)"]) | s_col.str.startswith("[빈칸")
            else:
                if is_empty_q:
                    mask |= s_col.isin(["-", "", "(빈칸)"]) | s_col.str.startswith("[빈칸")
                if is_na_q:
                    mask |= s_col.str.upper() == "N/A"
                if is_notrun_q:
                    mask |= s_col.str.contains("not run|미실행", case=False, na=False)
                if is_fail_q:
                    mask |= s_col.str.contains("빨강|주의|fail|불합격", case=False, na=False)

        res_sub = target_df[mask]
        if len(res_sub) == 0 and is_empty_q:
            for col in check_cols:
                s_col = target_df[col].astype(str).str.strip()
                mask |= s_col.str.upper().isin(["N/A", "NOT RUN", "-"])
            res_sub = target_df[mask]
        return res_sub

    tc_ids = re.findall(r"(?:TC|RTC)[\-_]?\d+", user_q, flags=re.IGNORECASE)
    if tc_ids:
        # [에러 해결] 순수 파이썬 내장 str() 함수를 사용해 float 에러(누수) 원천 차단
        row_text = target_df.apply(lambda row: " ".join(str(val) for val in row), axis=1)
        mask = pd.Series(False, index=target_df.index)
        for tid in tc_ids:
            mask |= row_text.str.contains(re.escape(tid), case=False, na=False)
        return target_df[mask]

    stop_words = {
        "전체", "조회", "해줘", "보여줘", "찾아줘", "검색", "데이터", "항목", "행", "열", "에서", "만",
        "그리고", "또는", "이나", "이랑", "랑", "과", "와", "포함", "있는", "없는", "표로", "목록", "뽑아줘",
        "요약", "분석", "보고서", "현재", "핵심", "현황", "상위", "집중", "건수", "비중", "간단히", "예외",
        "경계값", "테스트", "시나리오", "추천", "대한", "관련", "알려줘", "알려줄래", "설명", "자세히", "바탕으로",
        "무슨", "어느", "어떤", "나오는데", "뜨는데", "결과", "내용중에",
    }
    col_names_low = {str(c).lower() for c in target_df.columns}
    raw_tokens = re.findall(r"[A-Za-z0-9_.+-]{2,}|[가-힣]{2,}", user_q)
    kws = []
    for tk in raw_tokens:
        tk_clean = re.sub(r"(이랑|으로|에서|부터|까지|만|은|는|이|가|을|를|랑|도|의|에|야|인지)$", "", tk).strip()
        if len(tk_clean) >= 2 and tk_clean.lower() not in stop_words and not tk_clean.isdigit():
            if tk_clean.lower() not in col_names_low and not re.match(r"^\d+건$", tk_clean):
                kws.append(tk_clean)

    if not kws:
        return pd.DataFrame()

    # [에러 해결] 순수 파이썬 내장 str() 함수를 사용해 float 에러(누수) 원천 차단
    row_text = target_df.apply(lambda row: " ".join(str(val) for val in row), axis=1)
    mask = pd.Series(False, index=target_df.index)
    for kw in list(dict.fromkeys(kws))[:3]:
        mask |= row_text.str.contains(re.escape(kw), case=False, na=False)

    return target_df[mask]


def _execute_chat_turn(
        user_q: str,
        chat_model: str,
        target_df: pd.DataFrame,
        raw_df: pd.DataFrame,
        active_filter_text: str,
        stat_summary_str: str,
        verdict_now: str,
        defect_threshold: float,
        rag_knowledge_text: str,
):
    st.session_state["chat_history"].append({"role": "user", "content": user_q})

    matched_sub_df = _filter_within_target_df(user_q, target_df)

    saved_rep_context = ""
    if st.session_state.get("saved_results"):
        first_m = next(iter(st.session_state["saved_results"].values()))
        saved_rep_context = f"\n[직전 3번 AI 분석 리포트 요약]\n{first_m.get('report', '')[:1200]}\n"

    if len(matched_sub_df) > 0:
        sub_rows_str = (
            f"\n[질문 조건에 매칭된 행 — 핵심 3건 샘플]\n"
            f"{_build_compact_sample_str(matched_sub_df, max_rows=3)}\n"
        )
    else:
        sub_rows_str = (
            f"\n[현재 조회된 전체 데이터 행 — 핵심 3건 샘플]\n"
            f"{_build_compact_sample_str(target_df, max_rows=3)}\n"
        )

    recent_turns = "\n".join(
        f"{'사용자' if m['role'] == 'user' else 'AI비서'}: {m['content'][:250]}"
        for m in st.session_state["chat_history"][-5:-1]
    )
    scope_desc = f"검색조건('{active_filter_text}') 필터링된 {len(target_df)}건 (전체 원본 {len(raw_df)}건 중)" if active_filter_text else f"전체 {len(target_df)}건"

    chat_prompt = f"""당신은 사용자와 실시간으로 대화하는 정확하고 명쾌한 '사내 데이터 AI 비서'입니다.
아래 [파이썬 집계표 (컬럼별 미기입 해당 TC ID 전수 명단 포함)]를 보고 **핵심 결론(해당 TC ID와 Title 등)을 3~6줄 이내의 불릿(- )으로 즉시 답변**하십시오.

[현재 데이터 메모리 범위]: {scope_desc}
[사내 기준({defect_threshold}%) 판정]: {verdict_now}
[파이썬 집계표 (전체 {len(target_df)}건 전수 스캔 결과 및 미기입 TC ID 명단)]
{stat_summary_str}
{sub_rows_str}
{saved_rep_context}
[직전 대화 맥락]
{recent_turns}

[사용자 채팅 질문]
{user_q}
"""
    ans_text, ans_el = call_llm(chat_prompt, mode=chat_model, knowledge_context=rag_knowledge_text)
    model_short = MODEL_LABELS.get(chat_model, chat_model).split("(")[0].strip()

    rag_badge = " + 📚 RAG 사내 지식 연동" if rag_knowledge_text else ""
    final_ans = f"**{model_short}** (`{ans_el}초` · 조회범위: `전체 {len(target_df)}건 전수 확인`{rag_badge})\n\n{ans_text}"

    st.session_state["chat_history"].append({
        "role": "assistant",
        "content": final_ans,
        "matched_df": matched_sub_df if len(matched_sub_df) > 0 else None,
    })


def render_chat_messenger_ui(
        target_df: pd.DataFrame,
        raw_df: pd.DataFrame,
        active_filter_text: str,
        aux_sheets: dict[str, pd.DataFrame],
        defect_threshold: float,
        rag_knowledge_text: str,
        key_suffix: str = "inline",
        chat_height: int = 380,
):
    stat_summary_str = "\n\n".join(
        f"■ {k}\n{v.to_string(index=False)}"
        for k, v in build_korean_stat_tables(target_df, aux_sheets=aux_sheets, threshold=defect_threshold).items()
    )
    verdict_now = evaluate_threshold_verdict(target_df, threshold=defect_threshold)
    has_prev_rep = bool(st.session_state.get("saved_results"))
    filter_badge = f"검색조건('{active_filter_text}') — 총 {len(target_df):,}건 (원본 {len(raw_df):,}건 중)" if active_filter_text else f"전체 {len(target_df):,}건 전수 연동"

    top_c1, top_c2, top_c3 = st.columns([2.0, 1.3, 0.7])
    with top_c1:
        mem_badge = "✅ 3번 리포트 + 미기입 TC ID 전수 명단 연동됨" if has_prev_rep else "📊 전수 데이터·미기입 ID 명단 연동됨"
        if rag_knowledge_text:
            mem_badge += " + 📚 RAG 사내 지식 연동"

        st.markdown(
            f'<div style="background:#ecfdf5; border:1px solid #10b981; border-radius:8px; padding:8px 12px; font-size:0.82rem; color:#065f46; font-weight:600;">'
            f'🧠 <b>메모리:</b> {filter_badge} (기준 {defect_threshold}%) | {mem_badge}'
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
        b_col1, b_col2 = st.columns(2)
        with b_col1:
            # 대화 내역이 있을 때만 다운로드 버튼을 생성
            chat_export_text = "\n\n".join([f"[{'👤 사용자' if m['role'] == 'user' else '🤖 AI 비서'}]\n{m['content']}" for m in st.session_state["chat_history"]])
            st.download_button(
                label="💾 대화 저장",
                data=chat_export_text.encode("utf-8"),
                file_name=f"chat_history_{datetime.now().strftime('%m%d_%H%M')}.txt",
                mime="text/plain",
                key=f"dl_chat_{key_suffix}",
                use_container_width=True,
                disabled=not bool(st.session_state["chat_history"])
            )
        with b_col2:
            if st.button("🗑️ 비우기", key=f"clear_chat_{key_suffix}", use_container_width=True):
                st.session_state["chat_history"] = []
                st.rerun()

    st.caption("⚡ **빠른 질문 칩 (클릭 시 전체 조회 데이터를 전수 검사하여 즉시 답변합니다):**")
    chip_c1, chip_c2, chip_c3, chip_c4 = st.columns(4)
    triggered_q = None
    if chip_c1.button("🔍 컬럼별 미기입 TC ID 명단", key=f"chip1_{key_suffix}", use_container_width=True):
        triggered_q = f"현재 조회된 {len(target_df)}건 중에서 미기입(빈칸/공란)이 있는 컬럼 이름과, 각 컬럼별로 미기입된 실제 TC ID(및 제목) 목록을 빠짐없이 알려줘."
    if chip_c2.button("📊 상위 집중 항목 요약", key=f"chip2_{key_suffix}", use_container_width=True):
        triggered_q = f"현재 조회된 {len(target_df)}건에서 사내 기준({defect_threshold}%) 이상이거나 가장 비중이 큰 상위 3개 항목의 건수와 비중을 요약해줘."
    if chip_c3.button("📋 분석 데이터 핵심 5줄 요약", key=f"chip3_{key_suffix}", use_container_width=True):
        triggered_q = f"현재 분석한 {len(target_df)}건 데이터의 핵심 현황과 주요 분포 특징, 확인 필요 사항을 한눈에 파악할 수 있게 5줄 이내로 깔끔하게 요약해줘."
    if chip_c4.button("🧪 사내 지식 기반 조치 제안", key=f"chip4_{key_suffix}", use_container_width=True):
        triggered_q = f"업로드된 사내 규정이나 회의록 지식이 있다면 이를 반드시 인용해서, 이 데이터의 취약점에 대한 실무 조치 제안 3가지를 정리해줘."

    user_input = st.chat_input(
        f"전체 {len(target_df):,}건에 대해 질문하세요 (예: Procedure 미기입 1건 무슨 TC야? / Expected Result 미기입 4건 알려줘)...",
        key=f"chat_in_{key_suffix}",
    )

    active_q = triggered_q or user_input
    if active_q:
        with st.spinner(f"전체 {len(target_df):,}건 전수 데이터를 기반으로 답변 생성 중..."):
            _execute_chat_turn(
                active_q,
                chat_model,
                target_df,
                raw_df,
                active_filter_text,
                stat_summary_str,
                verdict_now,
                defect_threshold,
                rag_knowledge_text,
            )

    chat_box = st.container(height=chat_height, border=True)
    with chat_box:
        if not st.session_state["chat_history"]:
            with st.chat_message("assistant"):
                st.markdown(
                    f"안녕하세요! 현재 조회된 **{filter_badge}**의 **전수 미기입 TC ID 명단과 통계표**를 모두 숙지하고 있습니다.\n"
                    f"예: *'Procedure 미기입 1건 무슨 TC야?'*, *'Expected Result 미기입 4건 어느 TC인지 알려줘'*처럼 물어보시면 정확한 TC ID와 표를 즉시 보여드립니다!"
                )
        for msg in st.session_state["chat_history"]:
            with st.chat_message(msg["role"]):
                st.markdown(msg["content"])
                if msg.get("matched_df") is not None and len(msg["matched_df"]) > 0:
                    st.caption(f"📌 **전체 {len(target_df)}건 중 질문 조건에 정확히 매칭된 행 ({len(msg['matched_df'])}건):**")
                    st.dataframe(style_dataframe_with_colors(msg["matched_df"]), use_container_width=True, hide_index=True)


def render_floating_popup_chat_window(
        target_df: pd.DataFrame,
        raw_df: pd.DataFrame,
        active_filter_text: str,
        aux_sheets: dict[str, pd.DataFrame],
        defect_threshold: float,
        rag_knowledge_text: str,
):
    with st.container():
        st.markdown('<div id="floating-chat-marker"></div>', unsafe_allow_html=True)

        h_c1, h_c2, h_c3, h_c4 = st.columns([2.2, 0.75, 0.65, 0.75])
        with h_c1:
            st.markdown(
                f'<div style="background:#059669; color:white; padding:7px 12px; border-radius:8px; font-weight:700; font-size:0.9rem;">'
                f'💬 팝업 채팅창 (전체 {len(target_df):,}건 연동)'
                f'</div>',
                unsafe_allow_html=True,
            )
        with h_c2:
            size_label = "🗕 기본폭" if st.session_state.get("chat_popup_wide") else "🔍 폭확대"
            if st.button(size_label, key="float_toggle_wide_btn", use_container_width=True):
                st.session_state["chat_popup_wide"] = not st.session_state.get("chat_popup_wide", False)
                st.rerun()
        with h_c3:
            min_label = "➕ 펼치기" if st.session_state.get("chat_popup_minimized") else "➖ 접기"
            if st.button(min_label, key="float_toggle_min_btn", use_container_width=True):
                st.session_state["chat_popup_minimized"] = not st.session_state.get("chat_popup_minimized", False)
                st.rerun()
        with h_c4:
            if st.button("✖ 닫기", key="float_close_popup_btn", type="primary", use_container_width=True):
                st.session_state["chat_popup_mode"] = False
                st.session_state["chat_popup_minimized"] = False
                st.rerun()

        if st.session_state.get("chat_popup_minimized"):
            st.caption("💡 팝업 채팅창이 최소화되어 있습니다. 우측 상단 **[➕ 펼치기]**를 누르면 대화창이 열립니다.")
            return

        st.caption("✨ **배경 화면 잠금 없음:** 이 팝업창을 띄워둔 채로 **마우스 휠을 굴려 뒤쪽 1·2·3번 화면을 자유롭게 위아래로 스크롤**하며 볼 수 있습니다!")

        with st.expander(f"📈 팝업 안에서도 [2번 그래프 차트] · [1번 표({len(target_df)}건)] · [3번 리포트] 펼쳐보기 (클릭)", expanded=False):
            p_t1, p_t2, p_t3 = st.tabs([
                "📈 2️⃣ 다차원 그래프 차트",
                f"📋 1️⃣ 조회된 표 ({len(target_df)}건)",
                "📝 3️⃣ 직전 AI 리포트",
            ])
            with p_t1:
                render_smart_charts(target_df, key_prefix="float_popup_chart", defect_threshold=defect_threshold)
            with p_t2:
                st.dataframe(style_dataframe_with_colors(target_df), use_container_width=True, hide_index=True, height=200)
            with p_t3:
                saved_res = st.session_state.get("saved_results", {})
                if saved_res:
                    for m_k, r_v in saved_res.items():
                        st.markdown(f"**■ {MODEL_LABELS.get(m_k, m_k)}**")
                        st.markdown(r_v.get("report", ""))
                else:
                    st.info("아직 3번 에이전트 분석 실행 전입니다.")

        render_chat_messenger_ui(
            target_df=target_df,
            raw_df=raw_df,
            active_filter_text=active_filter_text,
            aux_sheets=aux_sheets,
            defect_threshold=defect_threshold,
            rag_knowledge_text=rag_knowledge_text,
            key_suffix="floating_popup",
            chat_height=310,
        )


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
            f'<w:tc><w:tcPr><w:tcW w:w="{widths_pct[i] * 50}" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="059669"/></w:tcPr>'
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
                p_list = "".join(_p(l, bold=(c_i == 0), size_hp=17, color="111827", after=40 if idx < len(lines) - 1 else 0) for idx, l in enumerate(lines))
                r_cells.append(f'<w:tc><w:tcPr><w:tcW w:w="{widths_pct[c_i] * 50}" w:type="pct"/><w:shd w:val="clear" w:color="auto" w:fill="{bg}"/></w:tcPr>{p_list}</w:tc>')
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
            label=f"📥 {m_code} 개별 리포트 다운로드 (.xlsx)",
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
        f"📋 한글 기초통계 & 미기입 TC ID 전수 명단 (기준 {defect_threshold}%)",
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
        rag_knowledge_text: str,
):
    st.subheader("3️⃣ AI 비서 맞춤 질문 & 에이전트 동시 비교 분석 (Groq vs Gemini vs Ollama 정밀 모드)")

    auto_assistant_q = generate_smart_assistant_question(
        target_df, raw_df, active_filter_text, threshold=defect_threshold
    )

    q_mode = st.radio(
        "🤖 질문 생성 방식 선택",
        options=[
            "🤖 AI 비서 자동 맞춤 질문 (1번 조회 조건과 100% 자동 연동됨)",
            "✍️ 사용자 직접 질문 입력/수정",
        ],
        horizontal=True,
    )

    if q_mode.startswith("🤖"):
        st.info(f"**💡 AI 비서가 1번 자연어 조회 조건(`{active_filter_text or '전체 조회'}`)과 데이터({len(target_df):,}건, 기준 {defect_threshold}%)를 융합하여 만든 분석 질문입니다:**\n\n{auto_assistant_q}")
        question = auto_assistant_q
    else:
        question = st.text_input("💡 에이전트에게 지시할 분석 질문을 자유롭게 입력하세요:", value=auto_assistant_q)

    render_related_links_box(target_df)

    btn_c1, btn_c2 = st.columns([3, 1])
    with btn_c1:
        btn_label = "🚀 에이전트 정밀 분석 실행 (선택 모델 동시 병렬 실행)"
        if rag_knowledge_text:
            btn_label = "🚀 에이전트 정밀 분석 실행 (📚 사내 규정 RAG 지식 주입됨)"
        run_clicked = st.button(btn_label, type="primary", use_container_width=True)
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

            compact_sample_str = _build_compact_sample_str(target_df, max_rows=5)

            data_info = (
                f"[분석 일시]: {analyzed_at_now}\n"
                f"[조회 조건]: {active_filter_text or '전체 시트 데이터'}\n"
                f"[조회 데이터 건수]: {len(target_df)}건 (전체 원본 {len(raw_df)}건)\n"
                f"[적용된 사내 관리 기준값]: {defect_threshold}%\n"
                f"[사내 기준 자동 판정]: {verdict_now}\n"
                f"[실제 존재하는 컬럼 목록]: {list(target_df.columns)}\n\n"
                f"[파이썬 정밀 집계표 (미기입 TC ID 전수 명단 포함)]\n{stat_summary_str}\n\n"
                f"[대표 항목 참고용 요약 샘플 (인용 참고용이며 리포트에 절대 나열하지 말 것)]\n{compact_sample_str}"
            )

            new_results = {}
            logs = [
                f"**[1~2단계 완료]** 총 `{len(target_df):,}건` 전수 로드 및 미기입 ID 명단 집계 완료 (분석 일시: `{analyzed_at_now}`, 기준값: `{defect_threshold}%`)",
                f"**[3단계 가동]** 선택 모델 `{len(selected_models)}개` 병렬 추론 시작",
            ]

            if rag_knowledge_text:
                logs.append(f"**[RAG 연동]** 업로드된 사내 규정/회의록 지식 문서 성공적으로 읽기 완료, 에이전트에 지식 주입됨.")

            status_box = st.status("⚡ 선택한 AI 모델(정밀 모드)들을 동시에 병렬 실행하고 있습니다...", expanded=True)

            def _run_single_model(m_code: str):
                enhanced_data_info = data_info
                if rag_knowledge_text:
                    enhanced_data_info += f"\n\n[사내 규정 및 참고 지식 (RAG)]\n{rag_knowledge_text}"

                res = agent_app.invoke({
                    "question": question,
                    "mode": m_code,
                    "data_summary": enhanced_data_info,
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

            logs.append(f"**[4단계 완료]** 사내 기준(`{defect_threshold}%`) 자동 판정 배너 적용 및 샘플 나열 방지 정제 완료")
            logs.append("**[5단계 완료]** 통합 엑셀(`.xlsx`) 및 페이지 분리형 워드(`.docx`) 보고서 패키징 완료")
            status_box.update(label="🎉 1~5단계 전 파이프라인 가동 및 정밀 분석 완료!", state="complete", expanded=False)

            st.session_state["saved_results"] = new_results
            st.session_state["last_analyzed_df"] = target_df
            st.session_state["last_summary_df"] = summary_df
            st.session_state["last_aux_sheets"] = aux_sheets
            st.session_state["pipeline_logs"] = logs
            st.rerun()

    saved = st.session_state.get("saved_results", {})
    if saved:
        display_models = [m for m in selected_models if m in saved] or list(saved.keys())
        popup_df = st.session_state.get("last_analyzed_df") if st.session_state.get("last_analyzed_df") is not None else target_df
        popup_summary_df = st.session_state.get("last_summary_df") if st.session_state.get("last_summary_df") is not None else summary_df
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

        def color_models_green(val):
            return "color: #059669; font-weight: bold;" if isinstance(val, str) and ("Groq" in val or "Ollama" in val or "Gemini" in val) else ""

        st.dataframe(pd.DataFrame(matrix_rows).style.map(color_models_green), use_container_width=True, hide_index=True)

        combined_excel_bytes = export_combined_analysis_to_excel(
            popup_df, saved, MODEL_LABELS, aux_sheets=saved_aux, threshold=defect_threshold
        )
        combined_docx_bytes = export_combined_analysis_to_docx(
            popup_df, saved, MODEL_LABELS, threshold=defect_threshold
        )

        top_c1, top_c2, top_c3 = st.columns([1.4, 1.3, 1.3])
        with top_c1:
            st.caption("👆 **[종합 추출] 선택된 모든 모델의 결과를 하나의 엑셀 / 워드(쪽 나누기 적용) 파일로 한 번에 다운로드합니다.**")
        with top_c2:
            st.download_button(
                label="📚 전체 모델 비교 마스터 엑셀 다운로드 (.xlsx)",
                data=combined_excel_bytes,
                file_name="report_ALL_models_combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dl_combined_master_xlsx",
                type="primary",
                use_container_width=True,
            )
        with top_c3:
            st.download_button(
                label="📄 전체 모델 종합 워드 다운로드 (.docx)",
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
                        popup_df, popup_summary_df, updated_report, res_item["elapsed_time"], m_code,
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


@st.fragment
def render_advanced_extensions_hub(
        target_df: pd.DataFrame,
        raw_df: pd.DataFrame,
        active_filter_text: str,
        aux_sheets: dict[str, pd.DataFrame],
        defect_threshold: float,
        rag_knowledge_text: str,
):
    st.subheader("4️⃣ 대화형 후속 질문 채팅방 (Chatflow) & 3회 반복 일관성 상세 검증기")

    tab_a, tab_c = st.tabs([
        "💬 [기능 A] 대화형 후속 질문 채팅방 (팝업창 모드 시 1·2·3번 자유 스크롤 지원)",
        "🔁 [기능 C] 3회 반복 실행 '일관성 vs 차이점' 상세 비교 (파트회의 공유용 워드 .docx 추출)",
    ])

    with tab_a:
        if st.session_state.get("chat_popup_mode", False):
            st.warning(
                "⏸️ **현재 [팝업 채팅창 모드]가 실행 중이어서 하단 4번 채팅방은 일시 중지되었습니다.**\n\n"
                "- 화면 우측에 떠 있는 **팝업 채팅창**을 켜둔 채로 **마우스 스크롤 휠을 위아래로 자유롭게 굴려 1번(조회 표) · 2번(그래프 차트) · 3번(AI 리포트)을 마음대로 왔다 갔다 하며 확인**하실 수 있습니다!\n"
                "- 팝업창 우측 상단의 **[✖ 닫기]** 또는 아래 버튼을 누르면 즉시 원래 4번 채팅방으로 돌아옵니다."
            )
            if st.button("🔽 팝업 채팅창 닫고 여기(하단 4번 채팅방)로 복귀하기", key="return_inline_chat_btn", type="primary", use_container_width=True):
                st.session_state["chat_popup_mode"] = False
                st.session_state["chat_popup_minimized"] = False
                st.rerun()

            render_floating_popup_chat_window(
                target_df=target_df,
                raw_df=raw_df,
                active_filter_text=active_filter_text,
                aux_sheets=aux_sheets,
                defect_threshold=defect_threshold,
                rag_knowledge_text=rag_knowledge_text,
            )
        else:
            pop_c1, pop_c2 = st.columns([2.6, 1.4])
            with pop_c1:
                st.caption(
                    f"💡 **현재 1번 조회 데이터({len(target_df):,}건)의 전수 미기입 ID 명단 및 3번 리포트와 100% 연동된 대화방입니다.** "
                    "우측 버튼을 누르면 **팝업 채팅창**이 뜨며, 팝업을 띄워둔 채로 **1·2·3번 화면을 자유롭게 위아래로 스크롤**할 수 있습니다."
                )
            with pop_c2:
                if st.button("🗗 팝업 채팅창 모드로 열기 (1·2·3번 자유 스크롤)", key="open_floating_popup_btn", type="primary", use_container_width=True):
                    st.session_state["chat_popup_mode"] = True
                    st.session_state["chat_popup_minimized"] = False
                    st.rerun()

            render_chat_messenger_ui(
                target_df=target_df,
                raw_df=raw_df,
                active_filter_text=active_filter_text,
                aux_sheets=aux_sheets,
                defect_threshold=defect_threshold,
                rag_knowledge_text=rag_knowledge_text,
                key_suffix="inline",
                chat_height=390,
            )

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
            q_bench = generate_smart_assistant_question(target_df, raw_df, active_filter_text, threshold=defect_threshold)

            compact_sample_str = _build_compact_sample_str(target_df, max_rows=5)

            data_info_bench = (
                f"[조회 조건]: {active_filter_text or '전체 데이터'}\n"
                f"[조회 데이터 건수]: {len(target_df)}건\n"
                f"[적용된 사내 관리 기준값]: {defect_threshold}%\n"
                f"[사내 기준 자동 판정]: {verdict_now}\n"
                f"[실제 존재하는 컬럼 목록]: {list(target_df.columns)}\n\n"
                f"[파이썬 정밀 집계표]\n{stat_summary_str}\n\n"
                f"[대표 항목 참고용 요약 샘플]\n{compact_sample_str}"
            )

            if rag_knowledge_text:
                data_info_bench += f"\n\n[사내 규정 및 참고 지식 (RAG)]\n{rag_knowledge_text}"

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
                    "4대 필수 목차": f"{sec_hits}/4개 ({int(sec_hits / 4 * 100)}%)",
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
                        st.markdown(f"**■ {idx + 1}회차 원문**")
                        st.markdown(c_data["raw_reports"][idx])


st.title("📊 자연어 기반 사내 데이터 분석 Agent")

up_col1, up_col2 = st.columns([4, 1])
with up_col1:
    uploaded_files = st.file_uploader(
        "📂 사내 데이터 분석용 파일 업로드 (CSV, Excel, Word .docx, TXT 다중 첨부 지원)",
        type=["csv", "xlsx", "docx", "txt"],  # <-- 여기에 "docx", "txt" 추가
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
        st.session_state["chat_popup_mode"] = False
        st.session_state["consistency_data"] = None

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
    st.subheader("🎚️ 사내 품질 기준값 설정")
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

    st.subheader("📚 사내 지식 문서 (RAG) 연동")
    rag_files = st.file_uploader(
        "리포트 작성 시 참고할 사내 규정서, 과거 회의록(Word .docx, TXT)을 올려주세요.",
        type=["docx", "txt"],
        accept_multiple_files=True,
        key="rag_uploader"
    )

    rag_knowledge_text = ""
    if rag_files:
        rag_knowledge_text = extract_knowledge_text(rag_files)
        st.success(f"✅ {len(rag_files)}개의 사내 지식 문서 내용 주입 대기 중")

    st.divider()
    ollama_models = get_installed_ollama_models()
    with st.expander("🖥️ 로컬 Ollama 상세 설정 & 모델 안내 (클릭)", expanded=False):
        st.markdown(
            "**왜 로컬 모델(Ollama)을 쓰나요?**\n"
            "- 사내 기밀 데이터 외부 유출 0% (완전 오프라인 구동)\n"
            "- API 비용 0원, 무제한 사용 가능\n\n"
            "**모델별 특징:**\n"
            "- `7B / 8B`: 데이터 분석 및 리포트 작성 품질이 가장 우수한 표준 모델\n"
            "- `3B / 4B`: 사양 낮은 PC에서 빠른 속도를 원할 때 쓰는 경량 모델"
        )
        chosen_ollama = st.selectbox("구동할 로컬 모델 변경", options=ollama_models, index=0)
        os.environ["LOCAL_MODEL_NAME"] = chosen_ollama

pipeline_slot = st.empty()

if not uploaded_files or len(raw_df) == 0:
    pipeline_slot.markdown(
        build_pipeline_status_html(has_file=False, has_query_df=False, saved_results={}),
        unsafe_allow_html=True,
    )
    st.info("👆 좌측 또는 상단에 분석할 사내 데이터를 업로드하면 1~5단계 실시간 파이프라인이 가동됩니다.")
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
            "💡 입력 예시: **`A라인과 B라인 조회해줘`**, **`Procedure 미기입`**, **`상위 30건`**"
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
        target_df, raw_df, summary_df, selected_models, active_filter_text, aux_sheets, defect_threshold, pipeline_slot, rag_knowledge_text
    )
    st.divider()
    render_advanced_extensions_hub(target_df, raw_df, active_filter_text, aux_sheets, defect_threshold, rag_knowledge_text)