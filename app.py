import streamlit as st
import pandas as pd
import plotly.express as px
from data_agent.graph import build_agent_graph
from data_agent.excel_exporter import export_analysis_to_excel

st.set_page_config(page_title="사내 데이터 분석 Agent", layout="wide")
st.title("📊 자연어 기반 사내 데이터 분석 Agent (프로토타입)")

# 사이드바: 보안 및 PoC 비교 모드 선택
with st.sidebar:
    st.header("⚙️ 실행 환경 설정 (PoC 비교)")
    mode = st.selectbox(
        "LLM 엔진 선택",
        options=["mock", "cloud_gemini", "cloud_openai", "local_ollama"],
        format_func=lambda x: {
            "mock": "🧪 테스트 모드 (LLM 없이 UI/기능 검증)",
            "cloud_gemini": "☁️ Cloud: Gemini 2.5 Flash (비교용)",
            "cloud_openai": "☁️ Cloud: GPT-4o-mini (비교용)",
            "local_ollama": "🔒 Local: Ollama 사내 보안 모드 (추후 연동)"
        }[x]
    )
    st.caption("※ Ollama 설치 전에는 '테스트 모드' 또는 'Cloud 모드'로 전체 기능을 테스트할 수 있습니다.")

# 1. 파일 업로드 영역
uploaded_file = st.file_uploader("📂 분석할 사내 데이터 파일(CSV 또는 Excel)을 업로드하세요", type=["csv", "xlsx"])

if uploaded_file is not None:
    # 파일 읽기
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    st.subheader("1️⃣ 데이터 미리보기 및 기본 스키마")
    col1, col2 = st.columns([2, 1])
    with col1:
        st.dataframe(df.head(10), use_container_width=True)
    with col2:
        summary_df = df.describe(include="all").fillna("")
        st.dataframe(summary_df, use_container_width=True)

    st.divider()

    # 2. 자동 시각화 영역
    st.subheader("2️⃣ 데이터 시각화 차트")
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    all_cols = df.columns.tolist()

    if len(num_cols) > 0:
        c1, c2, c3 = st.columns(3)
        x_col = c1.selectbox("X축 (기준 컬럼)", options=all_cols, index=0)
        y_col = c2.selectbox("Y축 (수치 컬럼)", options=num_cols, index=0)
        chart_type = c3.selectbox("차트 종류", options=["막대 차트 (Bar)", "선 차트 (Line)", "산점도 (Scatter)"])

        if chart_type == "막대 차트 (Bar)":
            fig = px.bar(df, x=x_col, y=y_col, title=f"{x_col}별 {y_col} 현황")
        elif chart_type == "선 차트 (Line)":
            fig = px.line(df, x=x_col, y=y_col, title=f"{x_col} 추이")
        else:
            fig = px.scatter(df, x=x_col, y=y_col, title=f"{x_col} vs {y_col} 분포")

        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("수치형 데이터가 없어 차트를 생략합니다.")

    st.divider()

    # 3. 자연어 질문 및 에이전트 실행 영역
    st.subheader("3️⃣ 자연어 질문 및 AI 에이전트 분석")
    question = st.text_input("💡 데이터에 대해 궁금한 점을 자연어로 질문하세요:", value="가장 두드러지는 핵심 트렌드와 이상 징후를 요약해 줘.")

    if st.button("🚀 에이전트 분석 실행", type="primary"):
        with st.spinner("LangGraph 에이전트가 데이터를 분석 중입니다..."):
            agent_app = build_agent_graph()
            data_info = f"컬럼 목록: {list(df.columns)}\n기초 통계:\n{df.describe().to_string()}"

            # 그래프 실행
            result_state = agent_app.invoke({
                "question": question,
                "mode": mode,
                "data_summary": data_info,
                "report": "",
                "elapsed_time": 0.0
            })

            st.success(f"분석 완료! (실행 모드: `{mode}` | 소요 시간: `{result_state['elapsed_time']}초`)")
            st.markdown(result_state["report"])

            # 4. 엑셀 파일 추출 및 다운로드 버튼
            excel_bytes = export_analysis_to_excel(
                df=df,
                summary_df=summary_df,
                report_text=result_state["report"],
                elapsed_time=result_state["elapsed_time"],
                mode=mode
            )
            st.download_button(
                label="📥 분석 결과 + 리포트 엑셀(.xlsx) 다운로드",
                data=excel_bytes,
                file_name=f"analysis_report_{mode}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )