import os
import re
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
from data_agent.tools.data_loader import load_and_sync_to_db, filter_by_keywords, run_sql_query, DB_PATH, TABLE_NAME

# Streamlit 기본 빨간색(#FF4B4B) 테마를 근본적으로 초록색(#059669)으로 바꾸기 위해 config.toml 자동 생성
CONFIG_DIR = Path(".streamlit")
CONFIG_FILE = CONFIG_DIR / "config.toml"
if not CONFIG_FILE.exists():
    CONFIG_DIR.mkdir(exist_ok=True)
    CONFIG_FILE.write_text(
        '[theme]\nprimaryColor = "#059669"\n',
        encoding="utf-8",
    )

st.set_page_config(page_title="사내 데이터 분석 Agent", layout="wide")

# 라디오 버튼, 멀티셀렉트 태그(칩), 실행 버튼 등 모든 빨간색 요소를 초록색(#059669)으로 강제 변환
st.markdown(
    """
    <style>
    :root {
        --primary-color: #059669 !important;
    }
    /* 멀티셀렉트 선택된 칩(Tag) 초록색 강제 적용 */
    span[data-baseweb="tag"],
    div[data-baseweb="select"] span[data-baseweb="tag"],
    [data-testid="stMultiSelect"] span[data-baseweb="tag"] {
        background-color: #059669 !important;
        color: #ffffff !important;
        border-radius: 6px !important;
        font-weight: 600 !important;
    }
    span[data-baseweb="tag"] span {
        color: #ffffff !important;
    }
    /* 라디오 버튼 선택 동그라미 초록색 강제 적용 */
    div[role="radiogroup"] label > div:first-child {
        border-color: #059669 !important;
    }
    div[role="radiogroup"] label[data-baseweb="radio"] > div:first-child {
        background-color: #059669 !important;
        border-color: #059669 !important;
    }
    /* Primary 버튼 초록색 적용 */
    button[kind="primary"],
    [data-testid="stBaseButton-primary"] {
        background-color: #059669 !important;
        border-color: #059669 !important;
        color: white !important;
    }
    button[kind="primary"]:hover,
    [data-testid="stBaseButton-primary"]:hover {
        background-color: #047857 !important;
        border-color: #047857 !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

st.title("📊 자연어 기반 사내 데이터 분석 Agent")

MODEL_LABELS = {
    "cloud_groq": "⚡ Cloud: Groq (초고속 무료 API)",
    "cloud_gemini": "☁️ Cloud: Gemini (Google 무료 API)",
    "local_ollama_fast": "🚀 Local: Ollama 빠른 모드 (사내보안·고속)",
    "local_ollama": "🔒 Local: Ollama 정밀 모드 (사내보안·2단계)",
    "mock": "🧪 테스트 모드 (Mock·0.2초 검증)",
}

if "saved_results" not in st.session_state:
    st.session_state["saved_results"] = {}
if "last_analyzed_df" not in st.session_state:
    st.session_state["last_analyzed_df"] = None


# =====================================================================
# 📊 모델 제목 클릭 시 뜨는 전용 UI 시각화 모달 팝업창 (@st.dialog)
# =====================================================================
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

    st.markdown("#### 📈 데이터 다각도 시각화 & 한글 통계 집계")
    p_tab1, p_tab2, p_tab3, p_tab4 = st.tabs([
        "📊 항목별 비교 막대 차트",
        "📈 추세 및 이상치 선 차트",
        "🍩 비중 분석 도넛 차트",
        "📋 한글 기초통계 & 라인별 집계표",
    ])

    all_cols = target_df.columns.tolist()
    x_default = "주차" if "주차" in all_cols else all_cols[0]
    color_col = "라인" if "라인" in all_cols else None

    with p_tab1:
        if num_cols:
            y_bar = "불량수" if "불량수" in num_cols else num_cols[0]
            fig_bar = px.bar(
                target_df, x=x_default, y=y_bar, color=color_col, barmode="group",
                text_auto=True, title=f"{x_default}별 {y_bar} 비교 현황"
            )
            st.plotly_chart(fig_bar, use_container_width=True, key=f"popup_bar_{m_code}")

    with p_tab2:
        if num_cols:
            plot_df = target_df.copy()
            if "생산량" in plot_df.columns and "불량수" in plot_df.columns:
                plot_df["불량률(%)"] = (plot_df["불량수"] / plot_df["생산량"] * 100).round(2)
                y_line = "불량률(%)"
            else:
                y_line = num_cols[-1]
            fig_line = px.line(
                plot_df, x=x_default, y=y_line, color=color_col, markers=True,
                title=f"{x_default}별 {y_line} 추이 및 이상 징후"
            )
            st.plotly_chart(fig_line, use_container_width=True, key=f"popup_line_{m_code}")

    with p_tab3:
        if num_cols:
            y_pie = "불량수" if "불량수" in num_cols else num_cols[0]
            name_col = color_col if color_col else x_default
            fig_pie = px.pie(
                target_df, names=name_col, values=y_pie, hole=0.4,
                title=f"{name_col}별 {y_pie} 점유율 비중"
            )
            st.plotly_chart(fig_pie, use_container_width=True, key=f"popup_pie_{m_code}")

    with p_tab4:
        stat_tables = build_korean_stat_tables(target_df)
        for t_title, t_df in stat_tables.items():
            st.markdown(f"**■ {t_title}**")
            st.dataframe(t_df, use_container_width=True, hide_index=True)

    st.divider()

    analysis_json = res_data.get("analysis_result", {})
    if analysis_json:
        st.markdown("#### 🎯 AI 핵심 인사이트 요약 카드")
        ic1, ic2 = st.columns(2)
        with ic1:
            st.warning("**🚨 감지된 이상 징후 (Anomalies)**\n\n" + "\n".join(f"- {a}" for a in analysis_json.get("anomalies", [])))
        with ic2:
            st.info("**📌 주요 트렌드 (Trends)**\n\n" + "\n".join(f"- {t}" for t in analysis_json.get("trends", [])))

    st.markdown("#### 📝 AI 생성 전체 리포트")
    st.markdown(res_data["report"])

    st.download_button(
        label=f"📥 {MODEL_LABELS.get(m_code, m_code)} 엑셀 리포트 다운로드 (.xlsx)",
        data=res_data["excel_bytes"],
        file_name=f"report_{m_code}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key=f"popup_dl_{m_code}",
        type="primary",
    )


def build_smart_sql(nl_query: str, df: pd.DataFrame, table_name: str) -> str:
    """한국어 자연어 질문을 분석해 SQL을 자동 생성합니다."""
    if not nl_query or not nl_query.strip():
        return f"SELECT * FROM {table_name} LIMIT 50;"

    q = nl_query.strip()
    where_clauses = []

    str_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    for col in str_cols:
        for val in df[col].dropna().astype(str).unique():
            val_clean = val.strip()
            if val_clean and (val_clean.lower() in q.lower() or f"{val_clean}{col}".lower() in q.lower()):
                where_clauses.append(f"{col} = '{val_clean}'")

    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    nums = re.findall(r"\d+(?:\.\d+)?", q)
    if nums and num_cols:
        target_num = nums[0]
        target_col = next((col for col in num_cols if col.lower() in q.lower()), None)
        if target_col is None:
            for col in num_cols:
                if ("불량" in q and "불량" in col) or ("생산" in q and "생산" in col):
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
        sql += " WHERE " + " AND ".join(list(dict.fromkeys(where_clauses)))
    sql += " LIMIT 50;"
    return sql


# =====================================================================
# ⚙️ 사이드바: 모델 선택 & 로컬 Ollama 설명 접기(Expander)
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
            default=["cloud_groq", "cloud_gemini", "local_ollama_fast"],
            format_func=lambda x: MODEL_LABELS[x],
        )

    st.divider()
    ollama_models = get_installed_ollama_models()
    # 평소에는 기본값(qwen2.5-coder:7b)으로 자동 동작하고, 궁금하거나 바꿀 때만 펼쳐보도록 정리
    with st.expander("🖥️ 로컬 Ollama 상세 설정 & 모델 안내 (클릭)", expanded=False):
        st.markdown(
            "**💡 로컬 Ollama를 왜 쓰나요?**\n"
            "- **보안:** 사내 데이터를 외부 인터넷으로 전송하지 않고 내 PC 안에서만 안전하게 분석합니다.\n"
            "- **무료:** API 요금이나 한도 없이 무제한 실행됩니다.\n\n"
            "**📌 세부 모델 차이점**\n"
            "- `qwen2.5-coder:7b`: **(기본 추천)** 표·숫자·SQL 분석 정확도 최고\n"
            "- `qwen3:8b` / `llama3.1:8b`: 한국어 보고서 문장력 우수\n"
            "- `qwen3:4b` / `qwen2.5:3b`: 일반 사무용 PC용 초고속 경량 모델"
        )
        chosen_ollama = st.selectbox(
            "구동할 로컬 모델 변경",
            options=ollama_models,
            index=0,
        )
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
    tab_nl, tab_kw, tab_sql = st.tabs(["🤖 자연어 → SQL 자동 변환", "🔍 키워드 조합 필터", "🗄️ 수동 SQL 편집기"])

    with tab_nl:
        st.caption(f"테이블(`{TABLE_NAME}`) 컬럼: `{list(raw_df.columns)}` | 예: **'A라인만 조회해줘'**, **'불량수 20이 넘는 행 찾아줘'**")
        nl_question = st.text_input("조회 조건을 자연어로 입력하세요:", value="A라인만 조회해줘")
        auto_sql = build_smart_sql(nl_question, raw_df, TABLE_NAME)
        st.code(auto_sql, language="sql")
        try:
            df = run_sql_query(auto_sql)
            st.success(f"✅ SQL 자동 실행 완료: 총 **{len(raw_df)}건** 중 **{len(df)}건** 추출됨")
            st.dataframe(df, use_container_width=True, hide_index=True)
        except Exception as e:
            st.error(f"SQL 실행 오류: {e}")
            df = raw_df

    with tab_kw:
        c1, c2 = st.columns([3, 1])
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,` 구분, 예: `B라인, W44`)", value="")
        match_type = c2.radio("조건", options=["OR", "AND"], horizontal=True)
        if kw_input.strip():
            df = filter_by_keywords(raw_df, kw_input, match_all=(match_type == "AND"))
        st.dataframe(df, use_container_width=True, hide_index=True)

    with tab_sql:
        sql_input = st.text_area("수동 SQL 입력", value=auto_sql, height=80)
        if st.button("수동 SQL 실행"):
            try:
                df = run_sql_query(sql_input)
                st.dataframe(df, use_container_width=True, hide_index=True)
            except Exception as e:
                st.error(f"오류: {e}")
        if DB_PATH.exists():
            with open(DB_PATH, "rb") as f:
                st.download_button("💾 SQLite DB 파일(.db) 다운로드", data=f, file_name="internal_analysis.db")

    target_df = df if len(df) > 0 else raw_df
    summary_df = target_df.describe(include="all").fillna("")
    st.divider()

    # 2. 시각화 차트 & 한글 기초통계/라인별 집계 탭
    st.subheader("2️⃣ 조회된 데이터 시각화 & 한글 기초통계 집계")
    v_tab1, v_tab2 = st.tabs(["📈 인터랙티브 차트 시각화", "📋 한글 기초통계 & 라인·주차별 집계표 (엑셀 2번 시트 미리보기)"])

    with v_tab1:
        num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
        all_cols = target_df.columns.tolist()
        if len(num_cols) > 0 and len(target_df) > 0:
            c1, c2, c3 = st.columns(3)
            x_col = c1.selectbox("X축 (기준)", options=all_cols, index=0)
            y_col = c2.selectbox("Y축 (수치)", options=num_cols, index=min(1, len(num_cols) - 1))
            chart_type = c3.selectbox("차트 종류", options=["막대 차트 (Bar)", "선 차트 (Line)", "산점도 (Scatter)"])
            color_col = "라인" if "라인" in all_cols else None
            if "막대" in chart_type:
                fig = px.bar(target_df, x=x_col, y=y_col, color=color_col, barmode="group", text_auto=True)
            elif "선" in chart_type:
                fig = px.line(target_df, x=x_col, y=y_col, color=color_col, markers=True)
            else:
                fig = px.scatter(target_df, x=x_col, y=y_col, color=color_col, size=y_col)
            st.plotly_chart(fig, use_container_width=True)

    with v_tab2:
        st.caption("기존 영어(count, mean, 25% 등) 표기를 제거하고, 엑셀 `2_기초통계_및_집계` 시트에 저장되는 100% 한글 집계표입니다.")
        korean_tables = build_korean_stat_tables(target_df)
        for t_title, t_df in korean_tables.items():
            st.markdown(f"**■ {t_title}**")
            st.dataframe(t_df, use_container_width=True, hide_index=True)

    st.divider()

    # =====================================================================
    # 3. 다중 AI 모델 병렬(동시) 비교 분석 & 통합 엑셀 추출
    # =====================================================================
    st.subheader("3️⃣ 자연어 질문 및 AI 에이전트 동시 비교 분석 (Groq vs Gemini vs Ollama)")
    question = st.text_input("💡 에이전트 분석 질문:", value="현재 조회된 데이터의 핵심 추세, 이상 항목, 원인 가설 및 조치 제안을 정리해 줘.")

    if st.button("🚀 에이전트 분석 실행 (선택 모델 동시 병렬 실행)", type="primary"):
        if not selected_models:
            st.warning("좌측 사이드바에서 최소 1개 이상의 모델을 선택해주세요.")
        else:
            agent_app = build_agent_graph()
            # 한글 집계표도 함께 요약 텍스트에 포함하여 LLM 분석 정확도 향상
            stat_summary_str = "\n".join(
                f"[{k}]\n{v.to_string(index=False)}" for k, v in build_korean_stat_tables(target_df).items()
            )
            data_info = (
                f"데이터 건수: {len(target_df)}건\n컬럼: {list(target_df.columns)}\n"
                f"핵심 집계:\n{stat_summary_str}\n\n원본 데이터:\n{target_df.to_string(index=False)}"
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
                excel_bytes = export_analysis_to_excel(target_df, summary_df, res["report"], res["elapsed_time"], m_code)
                res["excel_bytes"] = excel_bytes
                return m_code, res

            with ThreadPoolExecutor(max_workers=len(selected_models)) as executor:
                future_map = {executor.submit(_run_single_model, m): m for m in selected_models}
                for future in as_completed(future_map):
                    m_done, res_done = future.result()
                    new_results[m_done] = res_done
                    status_box.write(f"✅ **{MODEL_LABELS.get(m_done, m_done)}** 분석 완료! (`{res_done['elapsed_time']}초`)")

            status_box.update(label="🎉 모든 모델 분석 완료! (모델 제목 클릭 시 시각화 팝업 오픈 / 통합 엑셀 추출 가능)", state="complete", expanded=False)
            st.session_state["saved_results"] = new_results
            st.session_state["last_analyzed_df"] = target_df

    saved = st.session_state.get("saved_results", {})
    if saved:
        display_models = [m for m in selected_models if m in saved] or list(saved.keys())
        popup_df = st.session_state.get("last_analyzed_df") if st.session_state.get("last_analyzed_df") is not None else target_df

        # ⭐ 다중 모델 결과를 한 파일로 묶어주는 [종합 엑셀 리포트 추출] 버튼
        combined_excel_bytes = export_combined_analysis_to_excel(popup_df, saved, MODEL_LABELS)
        top_c1, top_c2 = st.columns([2, 1])
        with top_c1:
            st.caption("👆 **각 모델의 제목 버튼을 클릭하면 상세 시각화 차트 & 한글 집계 팝업창이 열립니다.**")
        with top_c2:
            st.download_button(
                label="📚 전체 모델 종합 비교 엑셀 추출 (통합본 .xlsx)",
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
                    label=f"📥 {m_code} 개별 엑셀 추출 (.xlsx)",
                    data=res_item["excel_bytes"],
                    file_name=f"report_{m_code}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_persist_{m_code}_{idx}",
                    use_container_width=True,
                )