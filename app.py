# app.py
import os
import re
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
        if "초록" in s or "🟢" in s or "기준(" in s and "이상 비중" in s:
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
    if q.replace(" ", "") in ("전체", "전체조회", "전체데이터", "모두조회", "다보여줘", "전체보여줘", "all"):
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

    rem_cleaned = re.sub(r"(전체|조회|해줘|보여줘|출력|데이터|만|에서|중)", "", q_without_limit).strip()
    if not rem_cleaned and limit_val is not None:
        return f"SELECT * FROM {table_name} LIMIT {limit_val};"

    if any(k in q_without_limit for k in ["빨강", "빨간", "초록", "노랑", "배경", "색상", "색깔"]):
        color_kw = "빨강" if ("빨강" in q_without_limit or "빨간" in q_without_limit) else (
            "초록" if "초록" in q_without_limit else ("노랑" if "노랑" in q_without_limit else "강조")
        )
        or_likes = [f'"{col}" LIKE \'%{color_kw}%\'' for col in df.columns]
        sql = f"SELECT * FROM {table_name} WHERE " + " OR ".join(or_likes)
        return f"{sql} LIMIT {limit_val};" if limit_val else f"{sql};"

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
        if len(vals) == 1:
            where_clauses.append(f'"{col}" LIKE \'%{vals[0]}%\'')
        else:
            or_part = " OR ".join(f'"{col}" LIKE \'%{v}%\'' for v in vals)
            where_clauses.append(f"({or_part})")

    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
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
    threshold: float = DEFAULT_DEFECT_THRESHOLD,
) -> str:
    if len(target_df) == 0:
        return "현재 데이터의 핵심 현황, 이상·주의 항목, 원인 가설 및 실무 조치 제안을 정리해 줘."

    all_cols = target_df.columns.tolist()
    cols_lower = " ".join(str(c).lower() for c in all_cols)
    scope_parts = []
    if user_filter_text and user_filter_text.strip():
        scope_parts.append(f"검색조건('{user_filter_text.strip()}')")

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

    render_related_links_box(target_df)

    if st.button("🚀 에이전트 정밀 분석 실행 (선택 모델 동시 병렬 실행)", type="primary"):
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
            logs.append("**[5단계 완료]** 통합 엑셀(`.xlsx`) 및 워드(`.docx`) 보고서 패키징 완료")
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
            st.caption("👆 **각 모델의 제목 버튼을 클릭하면 상세 시각화 팝업창이 열리며, 우측에서 엑셀(.xlsx)과 워드(.docx)를 모두 내려받을 수 있습니다.**")
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
                label="📄 전체 모델 종합 비교 워드 추출 (.docx)",
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
                # 슬라이더를 변경했을 때 상단 배너도 최신 threshold 기준으로 즉시 표시
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
# 📂 파일 업로드 (단일 및 다중 파일 동시 첨부 + 서로 다른 파일 오염 방지 지원)
# =====================================================================
uploaded_files = st.file_uploader(
    "📂 사내 데이터·문서 파일(CSV, Excel, Word .docx, TXT) 업로드 (여러 파일 첨부 및 개별/통합 선택 가능)",
    type=["csv", "xlsx", "docx", "txt"],
    accept_multiple_files=True,
)

raw_df = pd.DataFrame()
aux_sheets = {}
sheet_names = []
sheet_counts = {}
active_file_name = ""

if uploaded_files:
    if len(uploaded_files) == 1:
        target_upload = uploaded_files[0]
        active_file_name = target_upload.name
    else:
        file_options = [f.name for f in uploaded_files]
        st.info(f"🗂️ **총 {len(uploaded_files)}개 파일이 첨부되었습니다.** 서로 다른 양식의 파일이 섞여 분석이 왜곡되지 않도록 분석할 대상 파일을 선택하거나 동일 양식일 때만 통합 조회하세요.")
        chosen_file_mode = st.radio(
            "분석할 첨부 파일 선택:",
            options=file_options + ["📚 동일 양식 파일 전체 병합 조회"],
            index=0,
            horizontal=True,
        )
        if chosen_file_mode == "📚 동일 양식 파일 전체 병합 조회":
            merged_dfs = []
            merged_aux = {}
            base_cols = None
            schema_mismatch = False
            for f_obj in uploaded_files:
                f_df, _, _, f_aux = load_and_sync_to_db(f_obj, selected_sheet="ALL")
                if len(f_df) > 0:
                    cur_cols = set(f_df.columns) - {"시트명", "파일명"}
                    if base_cols is None:
                        base_cols = cur_cols
                    elif base_cols != cur_cols:
                        schema_mismatch = True
                    f_copy = f_df.copy()
                    f_copy.insert(0, "파일명", f_obj.name)
                    merged_dfs.append(f_copy)
                for k, v in f_aux.items():
                    merged_aux[f"{f_obj.name}_{k}"] = v
            if schema_mismatch:
                st.warning("⚠️ 첨부된 파일들의 컬럼 구조가 서로 다릅니다(관련 없는 파일 혼합). 정확한 통계를 위해 위 라디오 버튼에서 개별 파일을 선택해 분석하는 것을 권장합니다.")
            raw_df = pd.concat(merged_dfs, ignore_index=True) if merged_dfs else pd.DataFrame()
            aux_sheets = merged_aux
            active_file_name = "ALL_FILES"
            target_upload = None
        else:
            target_upload = next(f for f in uploaded_files if f.name == chosen_file_mode)
            active_file_name = target_upload.name

    if target_upload is not None:
        raw_df_all, sheet_names, sheet_counts, aux_sheets = load_and_sync_to_db(target_upload, selected_sheet="ALL")
        chosen_sheet = "ALL"

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
                raw_df, _, _, _ = load_and_sync_to_db(target_upload, selected_sheet=chosen_sheet)
        else:
            raw_df = raw_df_all
        active_file_name = f"{target_upload.name}_{chosen_sheet}"

    new_file_sig = f"{active_file_name}_{len(raw_df)}"
    if st.session_state.get("current_file_sig") != new_file_sig:
        st.session_state["current_file_sig"] = new_file_sig
        st.session_state["saved_results"] = {}
        st.session_state["pipeline_logs"] = []


# =====================================================================
# ⚙️ 사이드바 설정 (1.5% vs 10.0% 차이가 즉시 반영되는 사내 품질 기준값 슬라이더)
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
        f"(바꾸면 상단 판정 배너·통계표·엑셀 `자동 판정 요약`이 즉시 **{defect_threshold}%** 기준으로 변경됩니다)"
    )

    with st.expander("❓ [발표 Q&A] 1.5%와 10.0%가 어떻게 다른가요?", expanded=False):
        st.markdown(
            "- **제조 데이터:** 불량률이 **1.5%**를 넘는지 vs **10.0%**를 넘는지 비교해 `🚨 [확인 필요]` 라인을 판정합니다.\n"
            "- **TC/문서 데이터:** 각 컬럼의 **미기입(빈칸/N/A) 비율**이 기준값을 초과하는지, 그리고 각 분류 항목의 비중이 기준값 이상인지 판정합니다.\n"
            "- **엑셀·워드 연동:** 설정한 기준값 숫자(`1.5%`, `10.0%` 등)가 엑셀 `자동 판정 요약` 칸에 그대로 기록됩니다."
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
# 📊 메인 워크플로우 화면 출력
# =====================================================================
pipeline_slot = st.empty()

if not uploaded_files:
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
            "💡 **비워두거나 '전체조회' 입력 시 전체 건수(100% 전체 행)가 조회**되며, `'상위 50건'`, `'High만'` 등 조건 입력도 가능합니다."
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
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,` 구분, 예: `선물하기, High, 장바구니`)", value="")
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