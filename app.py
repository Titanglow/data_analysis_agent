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
    """웹 화면 표에서도 엑셀 배경색 태그(🔴빨강, 🟢초록, 🟡노랑 등)가 있는 셀에 실제 배경 색상을 입혀 보여줍니다."""
    def _highlight_cell(val):
        s = str(val)
        if "🔴" in s or "빨강" in s or "FAIL" in s.upper():
            return "background-color: #FEE2E2; color: #991B1B; font-weight: bold;"
        if "🟢" in s or "초록" in s:
            return "background-color: #DCFCE7; color: #166534; font-weight: bold;"
        if "🟡" in s or "노랑" in s or "🟠" in s or "주황" in s:
            return "background-color: #FEF9C3; color: #854D0E; font-weight: bold;"
        if "🎨" in s or "🔵" in s:
            return "background-color: #DBEAFE; color: #1E40AF; font-weight: bold;"
        return ""

    if len(df) <= 300:
        return df.style.map(_highlight_cell)
    return df


def build_smart_sql(nl_query: str, df: pd.DataFrame, table_name: str) -> str:
    """한국어 자연어 질문('A라인랑 B라인', '빨강 배경만', 'Clock 앱만')을 분석해 SQL을 자동 생성합니다."""
    if not nl_query or not nl_query.strip():
        return f"SELECT * FROM {table_name} LIMIT 100;"

    q = nl_query.strip()
    col_matches: dict[str, list[str]] = {}

    # 1. 색상 키워드('빨강', '초록', '노랑', '배경') 검색 시 전체 컬럼 LIKE 검색 지원
    if any(k in q for k in ["빨강", "빨간", "초록", "노랑", "배경", "색상", "색깔"]):
        color_kw = "빨강" if ("빨강" in q or "빨간" in q) else ("초록" if "초록" in q else ("노랑" if "노랑" in q else "배경"))
        or_likes = [f'"{col}" LIKE \'%{color_kw}%\'' for col in df.columns]
        return f"SELECT * FROM {table_name} WHERE " + " OR ".join(or_likes) + " LIMIT 100;"

    # 2. 범주형 컬럼 값 매칭
    str_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    for col in str_cols:
        matched_vals = []
        for val in df[col].dropna().astype(str).unique():
            val_clean = val.strip()
            if not val_clean or val_clean == "-" or len(val_clean) > 40:
                continue
            if val_clean.lower() in q.lower() or f"{val_clean}{col}".lower() in q.lower():
                matched_vals.append(val_clean)
        if matched_vals:
            col_matches[col] = list(dict.fromkeys(matched_vals))

    where_clauses = []
    for col, vals in col_matches.items():
        if len(vals) == 1:
            where_clauses.append(f'"{col}" = \'{vals[0]}\'')
        else:
            joined_vals = ", ".join(f"'{v}'" for v in vals)
            where_clauses.append(f'"{col}" IN ({joined_vals})')

    # 3. 수치형 조건 매칭
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
            where_clauses.append(f'"{target_col}" {op} {target_num}')

    sql = f"SELECT * FROM {table_name}"
    if where_clauses:
        sql += " WHERE " + " AND ".join(where_clauses)
    sql += " LIMIT 100;"
    return sql


def generate_smart_assistant_question(target_df: pd.DataFrame, raw_df: pd.DataFrame, user_filter_text: str) -> str:
    """생산/품질 파일이든 TC/일반 엑셀 파일이든 데이터 특성을 자동 파악해 비서처럼 맞춤 질문을 조합합니다."""
    if len(target_df) == 0:
        return "현재 데이터의 핵심 현황, 이상·주의 항목, 원인 가설 및 실무 조치 제안을 정리해 줘."

    all_cols = target_df.columns.tolist()
    scope_parts = []
    if user_filter_text and user_filter_text.strip():
        scope_parts.append(f"검색조건('{user_filter_text.strip()}')")

    if "시트명" in all_cols:
        sheets = target_df["시트명"].unique().tolist()
        scope_parts.append(f"대상 시트 {len(sheets)}개({', '.join(map(str, sheets[:3]))})")

    stat_snippets = [f"총 {len(target_df)}건"]

    # 생산 데이터인 경우 vs 범주형/TC 엑셀 데이터인 경우 분기
    if "생산량" in all_cols and "불량수" in all_cols:
        p_sum = int(target_df["생산량"].sum())
        d_sum = int(target_df["불량수"].sum())
        rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
        stat_snippets.append(f"총 생산량 {p_sum:,}개, 총 불량수 {d_sum:,}개(불량률 {rate:.2f}%)")
    else:
        row_text = target_df.astype(str).agg(" ".join, axis=1)
        red_cnt = int(row_text.str.contains("빨강|불합격|FAIL|주의", case=False, regex=True).sum())
        if red_cnt > 0:
            stat_snippets.append(f"🔴빨강배경/주의 표시 {red_cnt}건 감지")
        for c_cand in ["Domain", "App", "분류", "부서"]:
            if c_cand in all_cols:
                vals = target_df[c_cand].unique().tolist()
                stat_snippets.append(f"{c_cand}[{', '.join(map(str, vals[:3]))}]")

    scope_title = " · ".join(scope_parts) if scope_parts else "현재 조회 데이터"
    stat_str = ", ".join(stat_snippets)

    return (
        f"[AI 비서 자동 맞춤 질문] {scope_title} ({stat_str}) 데이터를 바탕으로 "
        f"① 시트·분류별 핵심 현황 및 비중(%), ② 빨간색 배경 표시(🔴주의/불합격)·미기재 항목 또는 기준 초과 이상 항목, "
        f"③ 데이터에 근거한 원인 가설, ④ 담당 부서별 즉시 조치(24h)·근본 대책(1개월) 및 목표 KPI를 정밀 분석해 줘."
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
    elif "App" in all_cols or "Domain" in all_cols:
        for c in ["Domain", "App"]:
            if c in all_cols:
                vals = [v for v in target_df[c].astype(str).unique() if v != "-"]
                keywords.extend(vals[:1])
        keywords.append("SW 테스트 자동화 품질 검증")
    else:
        cat_cols = target_df.select_dtypes(include=["object", "category"]).columns.tolist()
        if cat_cols:
            vals = [v for v in target_df[cat_cols[0]].astype(str).unique() if v != "-"]
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
# 📊 범용 스마트 차트 엔진: 수치형이 없어도(TC/텍스트 엑셀) 4대 차트 100% 생성!
# =====================================================================
def render_smart_charts(target_df: pd.DataFrame, key_prefix: str = "main"):
    """
    1) 수치형 컬럼이 있는 경우: 라인별·불량유형별·시계열·시간대별 수치/불량률 차트 생성
    2) 수치형 컬럼이 없는 경우(TC 시트 등 텍스트 엑셀): 시트/분류별 건수, 배경색/판정 상태 분포, 컬럼별 작성 완료율(%) 등 4대 차트 자동 생성!
    """
    if len(target_df) == 0:
        st.info("조회된 데이터가 없습니다.")
        return

    num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
    all_cols = target_df.columns.tolist()

    # [모드 A] 수치형 컬럼이 있는 데이터 (생산·품질·매출 등)
    if num_cols:
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
                fig1.update_traces(textposition="outside")
                st.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_num1")
            else:
                agg_basic = target_df.groupby(line_col, as_index=False)[val_col].sum()
                fig1 = px.bar(
                    agg_basic, x=line_col, y=val_col, color=line_col, text_auto=True,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"① {line_col}별 {val_col} 합계 비교",
                )
                fig1.update_traces(textposition="outside")
                st.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_num1")

        with row1_c2:
            if defect_type_col:
                agg_defect = target_df.groupby([defect_type_col, line_col], as_index=False)[val_col].sum()
                fig2 = px.bar(
                    agg_defect, x=defect_type_col, y=val_col, color=line_col, barmode="group", text_auto=True,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"② {defect_type_col}별 · {line_col}별 {val_col} 분류",
                )
                fig2.update_traces(textposition="outside")
                st.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_num2")
            else:
                agg_pie = target_df.groupby(line_col, as_index=False)[val_col].sum()
                fig2 = px.pie(
                    agg_pie, names=line_col, values=val_col, hole=0.4,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"② {line_col}별 {val_col} 비중",
                )
                st.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_num2")

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
                        title=f"③ {time_col}별 불량률(%) 추이 (기준 1.5% 점선)",
                    )
                    fig3.add_hline(y=1.5, line_dash="dash", line_color="#DC2626", annotation_text="사내 기준 (1.5%)")
                    st.plotly_chart(fig3, use_container_width=True, key=f"{key_prefix}_num3")
        with row2_c2:
            if shift_col:
                agg_shift = target_df.groupby([shift_col, defect_type_col or line_col], as_index=False)[val_col].sum()
                fig4 = px.bar(
                    agg_shift, x=shift_col, y=val_col, color=defect_type_col or line_col, barmode="group", text_auto=True,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"④ {shift_col}별 비교",
                )
                st.plotly_chart(fig4, use_container_width=True, key=f"{key_prefix}_num4")
        return

    # =========================================================================
    # [모드 B] 수치형 컬럼이 없는 텍스트/범주형 엑셀 파일 (TC 시트, 점검표, 이슈 목록 등)
    # =========================================================================
    st.caption("📊 **텍스트·범주형 문서(TC 검증 시트 등)가 감지되어 [건수 · 작성률(%) · 배경색/판정 분포] 다차원 시각화로 자동 전환되었습니다.**")

    primary_cat = next((c for c in ["시트명", "App", "Domain", "분류", "부서"] if c in all_cols), all_cols[0])
    secondary_cat = next((c for c in ["Domain", "App", "Pass/Fail Criteria", "Precondition"] if c in all_cols and c != primary_cat), all_cols[min(1, len(all_cols) - 1)])
    result_col = next((c for c in all_cols if any(k in c.lower() for k in ["actual", "result", "결과", "판정", "상태"])), all_cols[-1])

    r1_c1, r1_c2 = st.columns(2)

    # [차트 1] 시트별 / 주요 분류(App·Domain)별 항목 건수 및 비중 막대그래프
    with r1_c1:
        vc1 = target_df[primary_cat].astype(str).value_counts().reset_index()
        vc1.columns = [primary_cat, "건수"]
        vc1["비중(%)"] = (vc1["건수"] / len(target_df) * 100).round(1)
        vc1["표시라벨"] = vc1.apply(lambda r: f"{r['건수']}건 ({r['비중(%)']}%)", axis=1)
        fig_c1 = px.bar(
            vc1.head(15), x=primary_cat, y="건수", color=primary_cat, text="표시라벨",
            color_discrete_sequence=DISTINCT_PALETTE,
            title=f"① [{primary_cat}]별 데이터 건수 및 비중(%) 현황",
        )
        fig_c1.update_traces(textposition="outside")
        fig_c1.update_layout(yaxis_title="데이터 건수 (건)", margin=dict(t=50, b=20))
        st.plotly_chart(fig_c1, use_container_width=True, key=f"{key_prefix}_cat1")

    # [차트 2] 결과/판정(Actual Result 등 배경색 포함) 분포 현황
    with r1_c2:
        vc_res = target_df[result_col].astype(str).value_counts().reset_index()
        vc_res.columns = [result_col, "건수"]
        fig_c2 = px.pie(
            vc_res.head(10), names=result_col, values="건수", hole=0.4,
            color_discrete_sequence=["#DC2626", "#2563EB", "#059669", "#D97706", "#7C3AED"],
            title=f"② [{result_col}] 판정 · 셀 배경색(🔴빨강 등) 및 기재 상태 비중",
        )
        fig_c2.update_traces(textinfo="label+percent+value")
        st.plotly_chart(fig_c2, use_container_width=True, key=f"{key_prefix}_cat2")

    r2_c1, r2_c2 = st.columns(2)

    # [차트 3] 전체 컬럼별 데이터 작성 완료율(%) vs 미기재/주의 현황
    with r2_c1:
        comp_rows = []
        for col in all_cols:
            s_str = target_df[col].astype(str).str.strip()
            empty_c = int(s_str.isin(["-", "", "None", "nan"]).sum())
            color_c = int(s_str.str.contains("배경|빨강|초록|노랑|주황", regex=True).sum())
            normal_c = max(0, len(target_df) - empty_c - color_c)
            comp_rows.append({"컬럼명": col, "구분": "일반 텍스트 기재", "건수": normal_c})
            if color_c > 0:
                comp_rows.append({"컬럼명": col, "구분": "🔴 색상강조/배경표시", "건수": color_c})
            if empty_c > 0:
                comp_rows.append({"컬럼명": col, "구분": "⚪ 미기재(빈칸)", "건수": empty_c})
        comp_df = pd.DataFrame(comp_rows)
        fig_c3 = px.bar(
            comp_df, x="컬럼명", y="건수", color="구분", barmode="stack", text_auto=True,
            color_discrete_map={
                "일반 텍스트 기재": "#059669",
                "🔴 색상강조/배경표시": "#DC2626",
                "⚪ 미기재(빈칸)": "#9CA3AF",
            },
            title="③ 컬럼별 데이터 작성 현황 (텍스트 기재 vs 🔴색상표시 vs ⚪미기재)",
        )
        st.plotly_chart(fig_c3, use_container_width=True, key=f"{key_prefix}_cat3")

    # [차트 4] 교차 분류(시트명 x App/Domain) 또는 세부 분류 분포
    with r2_c2:
        if primary_cat != secondary_cat:
            cross_df = target_df.groupby([primary_cat, secondary_cat], as_index=False).size()
            cross_df.columns = [primary_cat, secondary_cat, "건수"]
            fig_c4 = px.bar(
                cross_df.head(20), x=primary_cat, y="건수", color=secondary_cat, barmode="group", text_auto=True,
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"④ [{primary_cat}] x [{secondary_cat}] 교차 분류 건수",
            )
            fig_c4.update_traces(textposition="outside")
            st.plotly_chart(fig_c4, use_container_width=True, key=f"{key_prefix}_cat4")


@st.dialog("📊 AI 모델별 상세 시각화 대시보드 & 인사이트 리포트", width="large")
def show_model_visual_popup(m_code: str, res_data: dict, target_df: pd.DataFrame):
    st.markdown(f"## {MODEL_LABELS.get(m_code, m_code)}")
    st.caption(f"⏱️ 분석 소요 시간: **{res_data['elapsed_time']}초** | 분석 대상 데이터: **{len(target_df)}건**")

    row_text = target_df.astype(str).agg(" ".join, axis=1)
    red_cnt = int(row_text.str.contains("빨강|불합격|FAIL", case=False, regex=True).sum())

    k1, k2, k3, k4 = st.columns(4)
    k1.metric("조회 데이터 건수", f"{len(target_df):,} 건")
    k2.metric("전체 컬럼 수", f"{len(target_df.columns)} 개")
    k3.metric("🔴 색상강조/주의 감지", f"{red_cnt:,} 건")
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
        "🎯 핵심 다차원 대시보드 (수치형·범주형·배경색 자동 대응)",
        "🛠️ 자유 조합 커스텀 차트 (모든 엑셀/CSV 범용)",
        "📋 한글 기초통계 & 시트·분류·판정별 집계표 (엑셀 2번 시트)",
    ])

    with v_tab1:
        render_smart_charts(target_df, key_prefix="main_dash")

    with v_tab2:
        st.caption("💡 숫자 컬럼이 없는 텍스트 엑셀 파일도 **[건수 (COUNT)]** 집계로 모든 컬럼을 막대 차트로 비교할 수 있습니다.")
        num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
        all_cols = target_df.columns.tolist()
        if len(target_df) > 0:
            c1, c2, c3, c4 = st.columns(4)
            x_col = c1.selectbox("X축 (기준 항목)", options=all_cols, index=0, key="frag_x")
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

            fig_custom = px.bar(
                custom_df.head(30), x=x_col, y=y_plot, color=color_col or x_col, barmode="group",
                text_auto=True, color_discrete_sequence=DISTINCT_PALETTE,
                title=f"[{x_col}]별 {y_plot} 비교",
            )
            fig_custom.update_traces(textposition="outside")
            st.plotly_chart(fig_custom, use_container_width=True, key="frag_custom_chart")

    with v_tab3:
        st.caption("엑셀 `2_기초통계_및_집계` 시트에 요약 차트와 함께 기록되는 100% 한글 집계표입니다.")
        korean_tables = build_korean_stat_tables(target_df)
        for t_title, t_df in korean_tables.items():
            st.markdown(f"**■ {t_title}**")
            st.dataframe(style_dataframe_with_colors(t_df), use_container_width=True, hide_index=True)


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
            "🤖 AI 비서 자동 맞춤 질문 (조회된 시트·컬럼·배경색/수치 자동 조합)",
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
            # 행 데이터가 너무 길 경우 상위 40건 + 색상 강조된 중요 행 우선 전달
            sample_rows_df = target_df.head(40)
            data_info = (
                f"[조회 조건]: {active_filter_text or '전체 시트 데이터'}\n"
                f"[조회 데이터 건수]: {len(target_df)}건\n"
                f"[컬럼 목록]: {list(target_df.columns)}\n\n"
                f"[파이썬 정밀 집계표]\n{stat_summary_str}\n\n"
                f"[상세 행 데이터 (배경색 태그 포함)]\n{sample_rows_df.to_string(index=False)}"
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

            status_box.update(label="🎉 모든 모델 정밀 분석 완료!", state="complete", expanded=False)
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
                    show_model_visual_popup(m_code, res_item, popup_df)

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
            "- **토큰 한도:** 최대 **3,500 토큰**(`num_ctx=8192`)으로 확장되어 긴 보고서도 잘림 없이 완주합니다."
        )
        chosen_ollama = st.selectbox("구동할 로컬 모델 변경", options=ollama_models, index=0)
        os.environ["LOCAL_MODEL_NAME"] = chosen_ollama

    if st.session_state["saved_results"]:
        if st.button("🗑️ 이전 분석 결과 초기화", use_container_width=True):
            st.session_state["saved_results"] = {}
            st.rerun()

uploaded_file = st.file_uploader("📂 사내 데이터 파일(CSV, Excel) 업로드 (다중 시트 & 셀 배경색 자동 인식)", type=["csv", "xlsx"])

if uploaded_file is not None:
    # 1차로 전체 시트 스캔 및 통합 로드
    raw_df_all, sheet_names, sheet_counts = load_and_sync_to_db(uploaded_file, selected_sheet="ALL")

    # 다중 시트 엑셀(.xlsx)일 경우 시트 선택 바 제공 (기본값: 모든 시트 통합 조회!)
    if len(sheet_names) > 1:
        sheet_summary = " | ".join(f"**{s}**: {cnt}건" for s, cnt in sheet_counts.items())
        st.info(f"📑 **다중 시트 엑셀 파일 감지 (총 {len(sheet_names)}개 시트, 합계 {len(raw_df_all)}건):** {sheet_summary}")
        sheet_options = ["ALL"] + sheet_names
        chosen_sheet = st.radio(
            "조회할 엑셀 시트 범위 선택:",
            options=sheet_options,
            index=0,
            format_func=lambda x: f"📚 전체 시트 통합 조회 (총 {len(raw_df_all)}건)" if x == "ALL" else f"📄 {x} ({sheet_counts.get(x, 0)}건)",
            horizontal=True,
        )
        if chosen_sheet == "ALL":
            raw_df = raw_df_all
        else:
            raw_df, _, _ = load_and_sync_to_db(uploaded_file, selected_sheet=chosen_sheet)
    else:
        raw_df = raw_df_all

    # 1. 자연어 쿼리 자동 생성 및 SQLite DB 조회
    st.subheader("1️⃣ 자연어 기반 데이터 수집 및 SQLite DB 자동 조회")
    tab_nl, tab_kw, tab_sql = st.tabs(["🤖 자연어 → SQL 자동 변환", "🔍 키워드 조합 필터", "🗄️ 수동 SQL 편집기 (.db / .sql 추출)"])

    active_filter_text = ""

    with tab_nl:
        st.caption(
            f"테이블(`{TABLE_NAME}`) 컬럼: `{list(raw_df.columns)}` | "
            "입력 예시: **'빨강 배경만 조회해줘'**, **'Clock만 보여줘'**, **'A라인랑 B라인 조회해줘'**"
        )
        nl_question = st.text_input("조회 조건을 자연어로 입력하세요 (비워두면 전체 시트 데이터를 분석합니다):", value="")
        auto_sql = build_smart_sql(nl_question, raw_df, TABLE_NAME)
        st.code(auto_sql, language="sql")
        try:
            df = run_sql_query(auto_sql)
            st.success(f"✅ SQL 자동 실행 완료: 전체 **{len(raw_df)}건** 중 **{len(df)}건** 조회됨 (셀 배경색 자동 표시 적용)")
            st.dataframe(style_dataframe_with_colors(df), use_container_width=True, hide_index=True)
            if nl_question.strip():
                active_filter_text = nl_question.strip()
        except Exception as e:
            st.error(f"SQL 실행 오류: {e}")
            df = raw_df

    with tab_kw:
        c1, c2 = st.columns([3, 1])
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,` 구분, 예: `빨강, Clock, 알람`)", value="")
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

    render_visualization_fragment(target_df)
    st.divider()
    render_agent_analysis_fragment(target_df, raw_df, summary_df, selected_models, active_filter_text)