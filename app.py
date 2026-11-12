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

# 빠른 모드를 제외하고 완성도 높은 '로컬 정밀 모드'로 일원화
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


def build_smart_sql(nl_query: str, df: pd.DataFrame, table_name: str) -> str:
    """동일 컬럼 복수 키워드는 IN (...) 조건으로, 서로 다른 컬럼 조건은 AND로 결합합니다."""
    if not nl_query or not nl_query.strip():
        return f"SELECT * FROM {table_name} LIMIT 50;"

    q = nl_query.strip()
    col_matches: dict[str, list[str]] = {}

    str_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    for col in str_cols:
        matched_vals = []
        for val in df[col].dropna().astype(str).unique():
            val_clean = val.strip()
            if not val_clean:
                continue
            if val_clean.lower() in q.lower() or f"{val_clean}{col}".lower() in q.lower():
                matched_vals.append(val_clean)
        if matched_vals:
            col_matches[col] = list(dict.fromkeys(matched_vals))

    where_clauses = []
    for col, vals in col_matches.items():
        if len(vals) == 1:
            where_clauses.append(f"{col} = '{vals[0]}'")
        else:
            joined_vals = ", ".join(f"'{v}'" for v in vals)
            where_clauses.append(f"{col} IN ({joined_vals})")

    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    nums = re.findall(r"\d+(?:\.\d+)?", q)
    if nums and num_cols:
        target_num = nums[0]
        target_col = next((col for col in num_cols if col.lower() in q.lower()), None)
        if target_col is None:
            for col in num_cols:
                if ("불량" in q and "불량" in col) or ("생산" in q and "생산" in col) or ("가동" in q and "가동" in col):
                    target_col = col
                    break
        if target_col:
            if any(k in q for k in ["넘", "초과", "보다 큰", "많은", "높은", ">"]):
                op = ">=" if "이상" in q else ">"
            elif any(k in q for k in ["이상", "부터", ">="]):
                op = ">="
            elif any(k in q for k in ["미만", "보다 작은", "적은", "낮은", "<"]):
                op = "<"
            elif any(k in q for k in ["이하", "이내", "까지", "<="]):
                op = "<="
            else:
                op = "="
            where_clauses.append(f"{target_col} {op} {target_num}")

    sql = f"SELECT * FROM {table_name}"
    if where_clauses:
        sql += " WHERE " + " AND ".join(where_clauses)
    sql += " LIMIT 50;"
    return sql


def generate_smart_assistant_question(target_df: pd.DataFrame, raw_df: pd.DataFrame, user_filter_text: str) -> str:
    """조회된 데이터의 실제 통계(라인, 생산량, 불량수, 불량률, 최다 불량유형)를 조합해 맞춤 질문을 생성합니다."""
    if len(target_df) == 0:
        return "현재 데이터의 핵심 추세, 이상 징후, 원인 가설 및 조치 제안을 정리해 줘."

    line_col = "라인" if "라인" in target_df.columns else ("생산라인" if "생산라인" in target_df.columns else None)
    scope_parts = []

    if user_filter_text and user_filter_text.strip():
        scope_parts.append(f"검색조건('{user_filter_text.strip()}')")

    if line_col:
        lines = target_df[line_col].dropna().astype(str).unique().tolist()
        if len(lines) < len(raw_df[line_col].dropna().unique()):
            scope_parts.append(f"대상 라인[{', '.join(lines)}]")
        else:
            scope_parts.append(f"전체 라인[{', '.join(lines)}]")

    stat_snippets = [f"조회 {len(target_df)}건"]
    if "생산량" in target_df.columns and "불량수" in target_df.columns:
        p_sum = int(target_df["생산량"].sum())
        d_sum = int(target_df["불량수"].sum())
        rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
        stat_snippets.append(f"총 생산량 {p_sum:,}개")
        stat_snippets.append(f"총 불량수 {d_sum:,}개(평균 불량률 {rate:.2f}%)")

    if "불량유형" in target_df.columns and "불량수" in target_df.columns:
        top_def = target_df.groupby("불량유형")["불량수"].sum()
        if not top_def.empty:
            stat_snippets.append(f"최다 불량유형 '{top_def.idxmax()}({int(top_def.max())}개)'")

    scope_title = " · ".join(scope_parts) if scope_parts else "현재 조회 데이터"
    stat_str = ", ".join(stat_snippets)

    return (
        f"[AI 비서 자동 맞춤 질문] {scope_title} ({stat_str}) 데이터를 바탕으로 "
        f"① 주차·시간대별 불량률 추세(%p), ② 주요 불량유형 및 기준(1.5%) 초과 이상 항목, "
        f"③ 해당 데이터 특성에 맞는 5M1E 원인 가설, ④ 담당 부서별 즉시 조치(24h)·근본 대책(1개월) 및 목표 KPI를 정밀 분석해 줘."
    )


@st.cache_data(ttl=600, show_spinner=False)
def fetch_related_news_links(query_keyword: str, max_items: int = 4) -> list[dict[str, str]]:
    """데이터에서 추출된 핵심 키워드로 Google 뉴스 RSS에서 최신 뉴스 제목과 링크를 가져옵니다."""
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
    """분석 중인 데이터와 연관된 실시간 뉴스 및 기술 표준 검색 링크를 제공합니다."""
    keywords = []
    if "불량유형" in target_df.columns and "불량수" in target_df.columns:
        top_defects = target_df.groupby("불량유형")["불량수"].sum().sort_values(ascending=False).index.tolist()
        keywords.extend(str(d) for d in top_defects[:2])
    elif "불량유형" in target_df.columns:
        keywords.extend(target_df["불량유형"].dropna().astype(str).unique()[:2].tolist())

    if not keywords:
        cat_cols = target_df.select_dtypes(include=["object", "category"]).columns.tolist()
        if cat_cols:
            keywords.extend(target_df[cat_cols[0]].dropna().astype(str).unique()[:2].tolist())

    search_query = f"제조 공정 {' '.join(keywords[:2])} 품질 관리" if keywords else "스마트팩토리 제조 품질 불량 분석"

    with st.expander(f"🌐 분석 데이터 연관 뉴스 & 외부 기술 자료 링크 자동 추천 (키워드: `{search_query}`)", expanded=True):
        custom_kw = st.text_input("🔎 연관 뉴스/기술자료 검색 키워드 (데이터 기반으로 자동 추출됨, 직접 수정 가능):", value=search_query)
        news_items = fetch_related_news_links(custom_kw, max_items=4)

        c_news, c_ref = st.columns([2, 1])
        with c_news:
            st.markdown("**📰 실시간 연관 뉴스 헤드라인 (클릭 시 이동)**")
            if news_items:
                for item in news_items:
                    st.markdown(f"- 🔗 [{item['title']}]({item['link']}) `{item['date']}`")
            else:
                st.caption("사내망 방화벽 환경이거나 실시간 RSS 응답이 지연 중입니다. 우측 바로가기 링크를 이용해 주세요.")

        with c_ref:
            q_enc = urllib.parse.quote(custom_kw)
            st.markdown("**📚 전문 기술 자료 & 포털 원클릭 검색**")
            st.markdown(
                f"- 🟢 [네이버 뉴스: '{custom_kw}' 검색 결과](https://search.naver.com/search.naver?where=news&query={q_enc})\n"
                f"- 🔵 [구글 뉴스: '{custom_kw}' 최신 동향](https://news.google.com/search?q={q_enc}&hl=ko&gl=KR)\n"
                f"- 📑 [Google 학술/기술논문(Scholar) 원인 검색](https://scholar.google.co.kr/scholar?q={q_enc})\n"
                f"- 🏭 [국가기술표준원(KATS) 품질 표준 검색](https://www.kats.go.kr/)"
            )


def render_smart_charts(target_df: pd.DataFrame, key_prefix: str = "main"):
    """사전 집계(Groupby) 후 정확한 숫자와 선명한 색상으로 다차원 차트를 그립니다."""
    num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
    cat_cols = target_df.select_dtypes(include=["object", "category"]).columns.tolist()
    all_cols = target_df.columns.tolist()

    if not num_cols or len(target_df) == 0:
        st.info("시각화할 수치형 데이터가 없습니다.")
        return

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
            fig1.update_traces(textposition="outside")
            fig1.update_layout(showlegend=True, yaxis_title="총 불량수 (개)", margin=dict(t=50, b=20))
            st.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_fig1")
        else:
            agg_basic = target_df.groupby(line_col, as_index=False)[val_col].sum()
            fig1 = px.bar(
                agg_basic, x=line_col, y=val_col, color=line_col, text_auto=True,
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"① {line_col}별 총 {val_col} 합계 비교",
            )
            fig1.update_traces(textposition="outside")
            st.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_fig1")

    with row1_c2:
        if defect_type_col and val_col in num_cols:
            agg_defect = target_df.groupby([defect_type_col, line_col], as_index=False)[val_col].sum()
            fig2 = px.bar(
                agg_defect, x=defect_type_col, y=val_col, color=line_col, barmode="group", text_auto=True,
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"② {defect_type_col}(원인)별 · {line_col}별 {val_col} 세부 분류",
            )
            fig2.update_traces(textposition="outside")
            fig2.update_layout(yaxis_title=f"{val_col} 합계", margin=dict(t=50, b=20))
            st.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_fig2")
        elif len(num_cols) >= 2:
            second_val = num_cols[1] if num_cols[1] != val_col else num_cols[0]
            agg_sec = target_df.groupby(line_col, as_index=False)[second_val].sum()
            fig2 = px.pie(
                agg_sec, names=line_col, values=second_val, hole=0.4,
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"② {line_col}별 {second_val} 점유 비중",
            )
            fig2.update_traces(textinfo="label+percent+value")
            st.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_fig2")

    row2_c1, row2_c2 = st.columns(2)
    with row2_c1:
        if time_col:
            plot_df = target_df.copy()
            if "생산량" in plot_df.columns and "불량수" in plot_df.columns:
                agg_time = plot_df.groupby([time_col, line_col], as_index=False)[["생산량", "불량수"]].sum()
                agg_time["불량률(%)"] = (agg_time["불량수"] / agg_time["생산량"] * 100).round(2)
                fig3 = px.line(
                    agg_time, x=time_col, y="불량률(%)", color=line_col, markers=True, text="불량률(%)",
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"③ {time_col}별 · {line_col} 불량률(%) 추이 (사내 기준 1.5% 선 표시)",
                )
                fig3.add_hline(y=1.5, line_dash="dash", line_color="#DC2626", annotation_text="사내 기준 (1.5%)")
                fig3.update_traces(textposition="top center")
                st.plotly_chart(fig3, use_container_width=True, key=f"{key_prefix}_fig3")
            else:
                agg_time = plot_df.groupby([time_col, line_col], as_index=False)[val_col].sum()
                fig3 = px.line(
                    agg_time, x=time_col, y=val_col, color=line_col, markers=True,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"③ {time_col}별 {val_col} 변화 추이",
                )
                st.plotly_chart(fig3, use_container_width=True, key=f"{key_prefix}_fig3")

    with row2_c2:
        if shift_col:
            agg_shift = target_df.groupby([shift_col, defect_type_col or line_col], as_index=False)[val_col].sum()
            fig4 = px.bar(
                agg_shift, x=shift_col, y=val_col, color=defect_type_col or line_col, barmode="group", text_auto=True,
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"④ {shift_col}(주/야간)별 · {defect_type_col or line_col} 비교",
            )
            fig4.update_traces(textposition="outside")
            st.plotly_chart(fig4, use_container_width=True, key=f"{key_prefix}_fig4")
        elif defect_type_col:
            agg_pie = target_df.groupby(defect_type_col, as_index=False)[val_col].sum()
            fig4 = px.pie(
                agg_pie, names=defect_type_col, values=val_col, hole=0.4,
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"④ 전체 {defect_type_col} 점유율 도넛 차트",
            )
            fig4.update_traces(textinfo="label+percent+value")
            st.plotly_chart(fig4, use_container_width=True, key=f"{key_prefix}_fig4")


@st.dialog("📊 AI 모델별 상세 시각화 대시보드 & 인사이트 리포트", width="large")
def show_model_visual_popup(m_code: str, res_data: dict, target_df: pd.DataFrame):
    st.markdown(f"## {MODEL_LABELS.get(m_code, m_code)}")
    st.caption(f"⏱️ 분석 소요 시간: **{res_data['elapsed_time']}초** | 분석 대상 데이터: **{len(target_df)}건**")

    num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
    k1, k2, k3, k4 = st.columns(4)
    k1.metric("조회 데이터 건수", f"{len(target_df):,} 건")
    if "생산량" in target_df.columns:
        k2.metric("총 생산량", f"{int(target_df['생산량'].sum()):,} 개")
    elif num_cols:
        k2.metric(f"평균 {num_cols[0]}", f"{target_df[num_cols[0]].mean():,.1f}")

    if "불량수" in target_df.columns:
        k3.metric("총 불량수", f"{int(target_df['불량수'].sum()):,} 개")
    if "생산량" in target_df.columns and "불량수" in target_df.columns and target_df["생산량"].sum() > 0:
        defect_rate = (target_df["불량수"].sum() / target_df["생산량"].sum()) * 100
        k4.metric("평균 불량률", f"{defect_rate:.2f} %")
    else:
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
def render_visualization_fragment(target_df: pd.DataFrame):
    st.subheader("2️⃣ 조회된 데이터 다차원 시각화 & 한글 기초통계 집계")
    v_tab1, v_tab2, v_tab3 = st.tabs([
        "🎯 핵심 다차원 대시보드 (라인·불량유형·시간대별)",
        "🛠️ 자유 조합 커스텀 차트 (축 변경 시 화면 초기화 없음)",
        "📋 한글 기초통계 & 라인·불량유형·주차별 집계표",
    ])

    with v_tab1:
        render_smart_charts(target_df, key_prefix="main_dash")

    with v_tab2:
        st.caption("💡 독립 렌더링(@st.fragment)이 적용되어 X축·Y축·색상 항목을 변경해도 아래 3번 AI 분석 결과가 초기화되지 않습니다.")
        num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
        all_cols = target_df.columns.tolist()
        if num_cols and len(target_df) > 0:
            c1, c2, c3, c4 = st.columns(4)
            x_col = c1.selectbox("X축 (기준 항목)", options=all_cols, index=0, key="frag_x")
            y_col = c2.selectbox("Y축 (수치 항목)", options=num_cols, index=min(1, len(num_cols) - 1), key="frag_y")
            color_options = ["(색상 분류 없음)"] + all_cols
            default_color_idx = (
                color_options.index("불량유형")
                if "불량유형" in color_options
                else (color_options.index("라인") if "라인" in color_options else 0)
            )
            color_sel = c3.selectbox("색상 구분 (세부 분류)", options=color_options, index=default_color_idx, key="frag_color")
            agg_func = c4.selectbox("집계 방식", options=["합계 (SUM)", "평균 (MEAN)", "건수 (COUNT)"], key="frag_agg")

            color_col = None if color_sel == "(색상 분류 없음)" else color_sel
            grp_cols = [x_col] + ([color_col] if color_col and color_col != x_col else [])

            if "합계" in agg_func:
                custom_df = target_df.groupby(grp_cols, as_index=False)[y_col].sum()
            elif "평균" in agg_func:
                custom_df = target_df.groupby(grp_cols, as_index=False)[y_col].mean().round(2)
            else:
                custom_df = target_df.groupby(grp_cols, as_index=False)[y_col].count()

            fig_custom = px.bar(
                custom_df, x=x_col, y=y_col, color=color_col or x_col, barmode="group",
                text_auto=True, color_discrete_sequence=DISTINCT_PALETTE,
                title=f"{x_col}별 {y_col} ({agg_func}) 비교",
            )
            fig_custom.update_traces(textposition="outside")
            st.plotly_chart(fig_custom, use_container_width=True, key="frag_custom_chart")

    with v_tab3:
        st.caption("엑셀 `2_기초통계_및_집계` 시트에 차트와 함께 저장되는 100% 한글 실무 집계표입니다.")
        korean_tables = build_korean_stat_tables(target_df)
        for t_title, t_df in korean_tables.items():
            st.markdown(f"**■ {t_title}**")
            st.dataframe(t_df, use_container_width=True, hide_index=True)


@st.fragment
def render_agent_analysis_fragment(
    target_df: pd.DataFrame,
    raw_df: pd.DataFrame,
    summary_df: pd.DataFrame,
    selected_models: list[str],
    active_filter_text: str,
):
    st.subheader("3️⃣ AI 비서 맞춤 질문 & 에이전트 동시 비교 분석 (Groq vs Gemini vs Ollama 정밀 모드)")

    auto_assistant_q = generate_smart_assistant_question(target_df, raw_df, active_filter_text)

    q_mode = st.radio(
        "🤖 질문 생성 방식 선택",
        options=[
            "🤖 AI 비서 자동 맞춤 질문 (1번 조회 데이터·수치·불량유형 자동 조합)",
            "✍️ 사용자 직접 질문 입력/수정",
        ],
        horizontal=True,
    )

    if q_mode.startswith("🤖"):
        st.info(f"**💡 AI 비서가 현재 조회된 데이터({len(target_df)}건)에 맞춰 조합한 분석 질문입니다:**\n\n{auto_assistant_q}")
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
                f"■ {k}\n{v.to_string(index=False)}" for k, v in build_korean_stat_tables(target_df).items()
            )
            data_info = (
                f"[조회 조건]: {active_filter_text or '전체 데이터'}\n"
                f"[조회 데이터 건수]: {len(target_df)}건\n"
                f"[컬럼 목록]: {list(target_df.columns)}\n\n"
                f"[파이썬 정밀 집계표]\n{stat_summary_str}\n\n"
                f"[상세 행 데이터]\n{target_df.to_string(index=False)}"
            )

            new_results = {}
            status_box = st.status("⚡ 선택한 AI 모델(정밀 모드)들을 동시에 병렬 실행하고 있습니다...", expanded=True)

            def _run_single_model(m_code: str):
                res = agent_app.invoke({
                    "question": question,
                    "mode": m_code,
                    "data_summary": data_info,
                    "analysis_result": {},
                    "report": "",
                    "elapsed_time": 0.0,
                })
                excel_bytes = export_analysis_to_excel(target_df, summary_df, res["report"], res["elapsed_time"], m_code)
                res["excel_bytes"] = excel_bytes
                return m_code, res

            with ThreadPoolExecutor(max_workers=len(selected_models)) as executor:
                future_map = {executor.submit(_run_single_model, m): m for m in selected_models}
                for future in as_completed(future_map):
                    m_done, res_done = future.result()
                    new_results[m_done] = res_done
                    status_box.write(f"✅ **{MODEL_LABELS.get(m_done, m_done)}** 정밀 분석 완료! (`{res_done['elapsed_time']}초`)")

            status_box.update(label="🎉 모든 모델 정밀 분석 완료! (차트 조작 시에도 이 결과는 초기화되지 않습니다)", state="complete", expanded=False)
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
            first_line = next((l.strip("> ") for l in r_item["report"].splitlines() if "확인 필요" in l or "이상 없음" in l), "본문 참조")
            matrix_rows.append({
                "AI 모델": MODEL_LABELS.get(m_code, m_code),
                "응답 소요 시간": f"{r_item['elapsed_time']} 초",
                "자동 기준 판정 (IF/ELSE)": first_line,
            })
        st.dataframe(pd.DataFrame(matrix_rows), use_container_width=True, hide_index=True)

        combined_excel_bytes = export_combined_analysis_to_excel(popup_df, saved, MODEL_LABELS)
        top_c1, top_c2 = st.columns([2, 1])
        with top_c1:
            st.caption("👆 **각 모델의 제목 버튼을 클릭하면 상세 다차원 시각화 팝업창이 열립니다.**")
        with top_c2:
            st.download_button(
                label="📚 전체 모델 종합 비교 엑셀 추출 (줄바꿈·차트 완비 통합본 .xlsx)",
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
                    label=f"📥 {m_code} 개별 엑셀 추출 (줄바꿈·차트 포함 .xlsx)",
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
            "**💡 로컬 Ollama 정밀 모드란?**\n"
            "- **보안:** 사내 데이터를 외부 인터넷으로 전송하지 않고 내 PC 안에서만 2단계(Analyzer → Reporter) 정밀 분석을 수행합니다.\n"
            "- **토큰 한도:** 최대 **3,500 토큰**(`num_ctx=8192`)으로 확장되어 긴 보고서도 잘림 없이 완주합니다.\n\n"
            "**📌 세부 모델 차이점**\n"
            "- `qwen2.5-coder:7b`: **(기본 추천)** 표·숫자·SQL 분석 정확도 최고\n"
            "- `qwen3:8b` / `llama3.1:8b`: 한국어 보고서 문장력 우수\n"
            "- `qwen3:4b` / `qwen2.5:3b`: 일반 사무용 PC용 경량 모델"
        )
        chosen_ollama = st.selectbox("구동할 로컬 모델 변경", options=ollama_models, index=0)
        os.environ["LOCAL_MODEL_NAME"] = chosen_ollama

    if st.session_state["saved_results"]:
        if st.button("🗑️ 이전 분석 결과 초기화", use_container_width=True):
            st.session_state["saved_results"] = {}
            st.rerun()

uploaded_file = st.file_uploader("📂 사내 데이터 파일(CSV, Excel) 업로드 (즉시 SQLite DB 자동 구축)", type=["csv", "xlsx"])

if uploaded_file is not None:
    raw_df = load_and_sync_to_db(uploaded_file)

    # 1. 자연어 쿼리 자동 생성 및 SQLite DB 조회
    st.subheader("1️⃣ 자연어 기반 데이터 수집 및 SQLite DB 자동 조회")
    tab_nl, tab_kw, tab_sql = st.tabs(["🤖 자연어 → SQL 자동 변환", "🔍 키워드 조합 필터", "🗄️ 수동 SQL 편집기 (.db / .sql 추출)"])

    active_filter_text = ""

    with tab_nl:
        st.caption(
            f"테이블(`{TABLE_NAME}`) 컬럼: `{list(raw_df.columns)}` | "
            "입력 예시: **'A라인랑 B라인 조회해줘'**, **'B라인만 조회해줘'**, **'도장얼룩 불량만 보여줘'**, **'야간 시간대 불량수 20 넘는 행'**"
        )
        nl_question = st.text_input("조회 조건을 자연어로 입력하세요 (비워두면 전체 데이터를 분석합니다):", value="")
        auto_sql = build_smart_sql(nl_question, raw_df, TABLE_NAME)
        st.code(auto_sql, language="sql")
        try:
            df = run_sql_query(auto_sql)
            st.success(f"✅ SQL 자동 실행 완료: 전체 **{len(raw_df)}건** 중 **{len(df)}건** 조회됨")
            st.dataframe(df, use_container_width=True, hide_index=True)
            if nl_question.strip():
                active_filter_text = nl_question.strip()
        except Exception as e:
            st.error(f"SQL 실행 오류: {e}")
            df = raw_df

    with tab_kw:
        c1, c2 = st.columns([3, 1])
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,` 구분, 예: `A라인, B라인`)", value="")
        match_type = c2.radio("조건", options=["OR (하나라도 포함)", "AND (모두 포함)"], horizontal=True)
        if kw_input.strip():
            df = filter_by_keywords(raw_df, kw_input, match_all=match_type.startswith("AND"))
            active_filter_text = f"키워드({kw_input.strip()})"
        st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_sql:
        st.caption("직접 SQL 쿼리를 수정·실행하고, 생성된 데이터베이스를 `.db` 원본 또는 **Notepad++ 열람용 `.sql` 스크립트**로 내려받을 수 있습니다.")
        sql_input = st.text_area("수동 SQL 입력", value=auto_sql, height=80)
        if st.button("▶️ 수동 SQL 실행"):
            try:
                df = run_sql_query(sql_input)
                st.success(f"쿼리 실행 성공: {len(df)}건 조회됨")
                st.dataframe(df, use_container_width=True, hide_index=True)
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

    render_visualization_fragment(target_df)
    st.divider()
    render_agent_analysis_fragment(target_df, raw_df, summary_df, selected_models, active_filter_text)