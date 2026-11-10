import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path
from data_agent.graph import build_agent_graph
from data_agent.excel_exporter import export_analysis_to_excel
from data_agent.tools.data_loader import load_and_sync_to_db, filter_by_keywords, run_sql_query, DB_PATH, TABLE_NAME

st.set_page_config(page_title="사내 데이터 분석 Agent", layout="wide")
st.title("📊 자연어 기반 사내 데이터 분석 Agent")

MODEL_LABELS = {
    "local_ollama": "🔒 Local: Ollama (사내보안)",
    "cloud_openai": "☁️ Cloud: ChatGPT (OpenAI)",
    "cloud_gemini": "☁️ Cloud: Gemini (Google)",
    "mock": "🧪 테스트 모드 (Mock)"
}

with st.sidebar:
    st.header("⚙️ 실행 및 비교 모델 설정")
    run_type = st.radio(
        "실행 방식 선택",
        options=["단일 모델 실행", "다중 모델 동시 비교 (PoC)"],
        index=1
    )

    if run_type == "단일 모델 실행":
        selected_models = [st.selectbox("사용할 AI 모델", options=list(MODEL_LABELS.keys()), format_func=lambda x: MODEL_LABELS[x])]
    else:
        selected_models = st.multiselect(
            "비교할 AI 모델들을 모두 선택하세요",
            options=["local_ollama", "cloud_openai", "cloud_gemini", "mock"],
            default=["local_ollama", "cloud_openai", "cloud_gemini"],
            format_func=lambda x: MODEL_LABELS[x]
        )

uploaded_file = st.file_uploader("📂 사내 데이터 파일(CSV, Excel) 업로드 (SQLite DB 자동 구축)", type=["csv", "xlsx"])

if uploaded_file is not None:
    raw_df = load_and_sync_to_db(uploaded_file)

    # 1. 키워드 조합 검색 및 SQL DB 필터링
    st.subheader("1️⃣ 키워드 조합 검색 & 자동 구축된 SQL DB 조회")
    tab_kw, tab_sql = st.tabs(["🔍 키워드 조합 검색 (데이터 수집/필터)", "🗄️ SQLite DB 직접 쿼리 & 파일 추출"])

    with tab_kw:
        c1, c2 = st.columns([3, 1])
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,`로 조합 가능, 예: `B라인, W44`)", value="")
        match_type = c2.radio("조합 조건", options=["OR (하나라도 포함)", "AND (모두 포함)"], horizontal=True)
        df = filter_by_keywords(raw_df, kw_input, match_all=(match_type.startswith("AND")))
        st.caption(f"전체 {len(raw_df)}건 중 **{len(df)}건** 필터링됨")
        st.dataframe(df, use_container_width=True)

    with tab_sql:
        default_sql = f"SELECT * FROM {TABLE_NAME} LIMIT 10;"
        sql_input = st.text_area("SQL 쿼리 입력 (테이블명: `internal_data`)", value=default_sql, height=80)
        if st.button("SQL 실행"):
            try:
                st.dataframe(run_sql_query(sql_input), use_container_width=True)
            except Exception as e:
                st.error(f"SQL 오류: {e}")
        if DB_PATH.exists():
            with open(DB_PATH, "rb") as f:
                st.download_button("💾 자동 생성된 SQLite DB 파일(.db) 다운로드", data=f, file_name="internal_analysis.db")

    summary_df = df.describe(include="all").fillna("")
    st.divider()

    # 2. 시각화 차트
    st.subheader("2️⃣ 필터링 데이터 시각화 차트")
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    all_cols = df.columns.tolist()
    if len(num_cols) > 0 and len(df) > 0:
        c1, c2, c3 = st.columns(3)
        x_col = c1.selectbox("X축 (기준)", options=all_cols, index=0)
        y_col = c2.selectbox("Y축 (수치)", options=num_cols, index=0)
        chart_type = c3.selectbox("차트 형태", options=["막대 차트 (Bar)", "선 차트 (Line)", "산점도 (Scatter)"])
        fig = px.bar(df, x=x_col, y=y_col) if "막대" in chart_type else (px.line(df, x=x_col, y=y_col) if "선" in chart_type else px.scatter(df, x=x_col, y=y_col))
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # 3. 자연어 분석 및 모델 비교 실행
    st.subheader("3️⃣ 자연어 질문 및 AI 에이전트 분석 (Ollama vs ChatGPT vs Gemini)")
    question = st.text_input("💡 질문 입력:", value="현재 데이터의 핵심 추세, 이상 항목, 원인 가설 및 조치 제안을 정리해 줘.")

    if st.button("🚀 에이전트 분석 실행", type="primary"):
        if not selected_models:
            st.warning("좌측 사이드바에서 최소 1개 이상의 모델을 선택해 주세요.")
        else:
            agent_app = build_agent_graph()
            data_info = f"데이터 건수: {len(df)}건\n컬럼: {list(df.columns)}\n통계 요약:\n{df.describe(include='all').to_string()}"

            cols = st.columns(len(selected_models))
            for idx, m_code in enumerate(selected_models):
                with cols[idx]:
                    st.markdown(f"#### {MODEL_LABELS[m_code]}")
                    with st.spinner(f"{MODEL_LABELS[m_code]} 분석 중..."):
                        res = agent_app.invoke({
                            "question": question,
                            "mode": m_code,
                            "data_summary": data_info,
                            "analysis_result": {},
                            "report": "",
                            "elapsed_time": 0.0
                        })
                        st.info(f"⏱️ 소요 시간: **{res['elapsed_time']}초**")
                        st.markdown(res["report"])

                        excel_bytes = export_analysis_to_excel(df, summary_df, res["report"], res["elapsed_time"], m_code)
                        st.download_button(
                            label=f"📥 {m_code} 결과 엑셀 추출",
                            data=excel_bytes,
                            file_name=f"report_{m_code}.xlsx",
                            key=f"dl_{m_code}"
                        )