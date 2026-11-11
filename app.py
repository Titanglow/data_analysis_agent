import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
import streamlit as st
import pandas as pd
import plotly.express as px
from data_agent.graph import build_agent_graph
from data_agent.excel_exporter import export_analysis_to_excel
from data_agent.llm import get_installed_ollama_models
from data_agent.tools.data_loader import load_and_sync_to_db, filter_by_keywords, run_sql_query, DB_PATH, TABLE_NAME

st.set_page_config(page_title="사내 데이터 분석 Agent", layout="wide")

# 초록색 테마 및 클릭 가능한 모델 타이틀 버튼 스타일
st.markdown(
    """
    <style>
    span[data-baseweb="tag"] {
        background-color: #059669 !important;
        color: white !important;
        border-radius: 6px !important;
    }
    button[kind="primary"] {
        background-color: #059669 !important;
        border-color: #059669 !important;
        color: white !important;
    }
    button[kind="primary"]:hover {
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

# 세션 상태 초기화 (엑셀 다운로드나 팝업 클릭 시에도 분석 결과가 절대 사라지지 않도록 보존)
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

    # 1. 상단 핵심 KPI 메트릭 카드
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

    # 2. 팝업 내 인터랙티브 시각화 차트 3종 탭
    st.markdown("#### 📈 데이터 다각도 시각화 분석")
    p_tab1, p_tab2, p_tab3 = st.tabs(["📊 항목별 비교 막대 차트", "📈 추세 및 이상치 선 차트", "🍩 비중 분석 파이 차트"])

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

    st.divider()

    # 3. Analyzer가 추출한 구조화 인사이트 요약 박스
    analysis_json = res_data.get("analysis_result", {})
    if analysis_json:
        st.markdown("#### 🎯 AI 핵심 인사이트 요약 카드")
        ic1, ic2 = st.columns(2)
        with ic1:
            st.warning("**🚨 감지된 이상 징후 (Anomalies)**\n\n" + "\n".join(f"- {a}" for a in analysis_json.get("anomalies", [])))
        with ic2:
            st.info("**📌 주요 트렌드 (Trends)**\n\n" + "\n".join(f"- {t}" for t in analysis_json.get("trends", [])))

    # 4. 전체 리포트 원문 및 팝업 내 엑셀 다운로드
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
# ⚙️ 사이드바: 모델 선택 & Ollama 세부 모델(7B / 8B / 경량) 선택
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
    st.subheader("🖥️ 로컬 Ollama 모델 상세 설정")
    ollama_models = get_installed_ollama_models()
    chosen_ollama = st.selectbox(
        "로컬 구동 모델 선택 (7B / 8B / 경량)",
        options=ollama_models,
        index=0,
        help="PC에 설치된 Ollama 모델을 선택합니다. 새 모델은 터미널에서 'ollama pull qwen3:8b'로 추가할 수 있습니다.",
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
            st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.error(f"SQL 실행 오류: {e}")
            df = raw_df

    with tab_kw:
        c1, c2 = st.columns([3, 1])
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,` 구분, 예: `B라인, W44`)", value="")
        match_type = c2.radio("조건", options=["OR", "AND"], horizontal=True)
        if kw_input.strip():
            df = filter_by_keywords(raw_df, kw_input, match_all=(match_type == "AND"))
        st.dataframe(df, use_container_width=True)

    with tab_sql:
        sql_input = st.text_area("수동 SQL 입력", value=auto_sql, height=80)
        if st.button("수동 SQL 실행"):
            try:
                df = run_sql_query(sql_input)
                st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"오류: {e}")
        if DB_PATH.exists():
            with open(DB_PATH, "rb") as f:
                st.download_button("💾 SQLite DB 파일(.db) 다운로드", data=f, file_name="internal_analysis.db")

    summary_df = df.describe(include="all").fillna("") if len(df) > 0 else raw_df.describe(include="all").fillna("")
    st.divider()

    # 2. 기본 시각화 차트
    st.subheader("2️⃣ 조회된 데이터 시각화 차트")
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    all_cols = df.columns.tolist()
    if len(num_cols) > 0 and len(df) > 0:
        c1, c2, c3 = st.columns(3)
        x_col = c1.selectbox("X축 (기준)", options=all_cols, index=0)
        y_col = c2.selectbox("Y축 (수치)", options=num_cols, index=min(1, len(num_cols) - 1))
        chart_type = c3.selectbox("차트 종류", options=["막대 차트 (Bar)", "선 차트 (Line)", "산점도 (Scatter)"])
        color_col = "라인" if "라인" in all_cols else None
        if "막대" in chart_type:
            fig = px.bar(df, x=x_col, y=y_col, color=color_col, barmode="group")
        elif "선" in chart_type:
            fig = px.line(df, x=x_col, y=y_col, color=color_col, markers=True)
        else:
            fig = px.scatter(df, x=x_col, y=y_col, color=color_col, size=y_col)
        st.plotly_chart(fig, use_container_width=True)

    st.divider()

    # =====================================================================
    # 3. 다중 AI 모델 병렬(동시) 비교 분석 & 결과 영구 유지 & 팝업 호출
    # =====================================================================
    st.subheader("3️⃣ 자연어 질문 및 AI 에이전트 동시 비교 분석 (Groq vs Gemini vs Ollama)")
    question = st.text_input("💡 에이전트 분석 질문:", value="현재 조회된 데이터의 핵심 추세, 이상 항목, 원인 가설 및 조치 제안을 정리해 줘.")

    target_df = df if len(df) > 0 else raw_df

    # 실행 버튼 클릭 시: ThreadPoolExecutor로 선택한 모든 모델을 '동시에(병렬)' 실행하여 세션에 저장
    if st.button("🚀 에이전트 분석 실행 (선택 모델 동시 병렬 실행)", type="primary"):
        if not selected_models:
            st.warning("좌측 사이드바에서 최소 1개 이상의 모델을 선택해주세요.")
        else:
            agent_app = build_agent_graph()
            data_info = f"데이터 건수: {len(target_df)}건\n컬럼: {list(target_df.columns)}\n데이터 내역:\n{target_df.to_string(index=False)}"

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

            # 병렬 실행: 로컬 모델을 기다리지 않고 Groq, Gemini, Ollama가 동시에 출발!
            with ThreadPoolExecutor(max_workers=len(selected_models)) as executor:
                future_map = {executor.submit(_run_single_model, m): m for m in selected_models}
                for future in as_completed(future_map):
                    m_done, res_done = future.result()
                    new_results[m_done] = res_done
                    status_box.write(f"✅ **{MODEL_LABELS.get(m_done, m_done)}** 분석 완료! (`{res_done['elapsed_time']}초`)")

            status_box.update(label="🎉 모든 모델 분석 완료! (아래 모델 제목 버튼을 누르면 시각화 팝업창이 열립니다)", state="complete", expanded=False)
            st.session_state["saved_results"] = new_results
            st.session_state["last_analyzed_df"] = target_df

    # 저장된 분석 결과 렌더링 (엑셀 다운로드 버튼을 눌러도 절대 사라지지 않음!)
    saved = st.session_state.get("saved_results", {})
    if saved:
        display_models = [m for m in selected_models if m in saved] or list(saved.keys())
        popup_df = st.session_state.get("last_analyzed_df") if st.session_state.get("last_analyzed_df") is not None else target_df

        st.caption("👆 **각 모델의 제목 버튼을 클릭하면 상세 시각화 차트 & KPI 팝업창이 열립니다.** (엑셀 추출을 눌러도 결과가 유지됩니다)")
        cols = st.columns(len(display_models))

        for idx, m_code in enumerate(display_models):
            res_item = saved[m_code]
            with cols[idx]:
                # 모델 제목 자체를 클릭 가능한 팝업 오픈 버튼으로 배치
                if st.button(f"📊 {MODEL_LABELS.get(m_code, m_code)}\n(클릭 시 시각화 팝업 열기 🔍)", key=f"title_btn_{m_code}", use_container_width=True):
                    show_model_visual_popup(m_code, res_item, popup_df)

                st.info(f"⏱️ 소요 시간: **{res_item['elapsed_time']}초**")
                st.markdown(res_item["report"])

                st.download_button(
                    label=f"📥 {m_code} 엑셀 리포트 추출 (.xlsx)",
                    data=res_item["excel_bytes"],
                    file_name=f"report_{m_code}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    key=f"dl_persist_{m_code}_{idx}",
                )