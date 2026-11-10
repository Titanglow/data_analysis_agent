import re
import streamlit as st
import pandas as pd
import plotly.express as px
from data_agent.graph import build_agent_graph
from data_agent.excel_exporter import export_analysis_to_excel
from data_agent.tools.data_loader import load_and_sync_to_db, filter_by_keywords, run_sql_query, DB_PATH, TABLE_NAME

st.set_page_config(page_title="사내 데이터 분석 Agent", layout="wide")

# UI 커스텀 스타일: 빨간색(기본값) 멀티셀렉트 태그와 버튼을 초록색(에메랄드 그린)으로 변경
st.markdown(
    """
    <style>
    /* 멀티셀렉트 선택 칩(Tag) 초록색 변경 */
    span[data-baseweb="tag"] {
        background-color: #059669 !important;
        color: white !important;
        border-radius: 6px !important;
    }
    /* 기본 Primary 버튼 초록색 변경 */
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
    "local_ollama": "🔒 Local: Ollama (사내보안·무료)",
    "cloud_openai": "☁️ Cloud: ChatGPT (OpenAI)",
    "cloud_gemini": "☁️ Cloud: Gemini (Google)",
    "mock": "🧪 테스트 모드 (Mock·0.3초 검증)"
}


def build_smart_sql(nl_query: str, df: pd.DataFrame, table_name: str) -> str:
    """한국어 자연어 질문('A라인만 조회해줘', '불량수 20이 넘는 행 찾아줘')을 분석해 정확한 SQL을 생성합니다."""
    if not nl_query or not nl_query.strip():
        return f"SELECT * FROM {table_name} LIMIT 20;"

    q = nl_query.strip()
    where_clauses = []

    # 1. 범주형(문자열) 컬럼 값 매칭 (조사 '만', '이', '가', '은', '는' 등이 붙어도 정확히 인식)
    str_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
    for col in str_cols:
        unique_vals = df[col].dropna().astype(str).unique()
        for val in unique_vals:
            val_clean = val.strip()
            if not val_clean:
                continue
            # 예: 데이터 값이 'A라인'이거나 'A'일 때 질문에 'A라인'이 포함된 경우
            if val_clean.lower() in q.lower() or f"{val_clean}{col}".lower() in q.lower():
                where_clauses.append(f"{col} = '{val_clean}'")

    # 2. 수치형 컬럼 조건 매칭 (예: '불량수 20이 넘는', '생산량 1200 이상')
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    nums = re.findall(r"\d+(?:\.\d+)?", q)
    if nums and num_cols:
        target_num = nums[0]
        # 질문에 언급된 수치형 컬럼 찾기 (언급된 컬럼이 없으면 '불량' 관련 키워드 우선 매칭)
        target_col = None
        for col in num_cols:
            if col.lower() in q.lower():
                target_col = col
                break
        if target_col is None:
            for col in num_cols:
                if "불량" in q and "불량" in col:
                    target_col = col
                    break
                elif "생산" in q and "생산" in col:
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

    # 3. 정렬(ORDER BY) 조건 감지
    order_clause = ""
    for col in num_cols:
        if col in q and any(k in q for k in ["순", "정렬", "가장 높", "가장 많"]):
            order_clause = f" ORDER BY {col} DESC"
            break

    sql = f"SELECT * FROM {table_name}"
    if where_clauses:
        # 중복 제거 후 결합
        unique_where = list(dict.fromkeys(where_clauses))
        sql += " WHERE " + " AND ".join(unique_where)
    sql += order_clause + " LIMIT 50;"
    return sql


with st.sidebar:
    st.header("⚙️ 실행 및 비교 모델 설정")
    run_type = st.radio("실행 방식 선택", options=["단일 모델 실행", "다중 모델 동시 비교 (PoC)"], index=1)

    if run_type == "단일 모델 실행":
        selected_models = [st.selectbox("사용할 AI 모델", options=list(MODEL_LABELS.keys()), format_func=lambda x: MODEL_LABELS[x])]
    else:
        selected_models = st.multiselect(
            "비교할 AI 모델들을 선택하세요 (초록색 활성화)",
            options=["local_ollama", "cloud_openai", "cloud_gemini", "mock"],
            default=["local_ollama", "cloud_openai", "cloud_gemini"],
            format_func=lambda x: MODEL_LABELS[x]
        )

    st.divider()
    st.caption("💡 **API 키 안내**\n- **Ollama:** 사내 로컬 무료 구동\n- **Gemini:** AI Studio에서 `AIzaSy...` 키 발급 시 무료\n- **OpenAI:** 크레딧 충전 계정 필요 (미충전 시 자동 백업 리포트 동작)")

uploaded_file = st.file_uploader("📂 사내 데이터 파일(CSV, Excel) 업로드 (업로드 즉시 SQLite DB 자동 구축)", type=["csv", "xlsx"])

if uploaded_file is not None:
    raw_df = load_and_sync_to_db(uploaded_file)

    # 1. 자연어 쿼리 자동 생성 및 SQLite DB 조회
    st.subheader("1️⃣ 자연어 기반 데이터 수집 및 SQLite DB 자동 조회")
    tab_nl, tab_kw, tab_sql = st.tabs(["🤖 자연어 → SQL 자동 변환 (추천)", "🔍 키워드 조합 필터", "🗄️ 수동 SQL 편집기"])

    with tab_nl:
        st.caption(f"현재 테이블(`{TABLE_NAME}`) 컬럼: `{list(raw_df.columns)}` | 예시: **'A라인만 조회해줘'**, **'불량수 20이 넘는 행 찾아줘'**, **'B라인에서 불량수 25 이상'**")
        nl_question = st.text_input("조회할 데이터 조건을 자연어로 입력하세요 (입력 즉시 SQL이 자동 생성됩니다):", value="A라인만 조회해줘")

        auto_sql = build_smart_sql(nl_question, raw_df, TABLE_NAME)
        st.code(auto_sql, language="sql")

        try:
            df = run_sql_query(auto_sql)
            st.success(f"✅ 자동 생성된 SQL 실행 완료: 총 **{len(raw_df)}건** 중 조건에 맞는 **{len(df)}건**이 추출되었습니다.")
            st.dataframe(df, use_container_width=True)
        except Exception as e:
            st.error(f"SQL 실행 오류: {e}")
            df = raw_df

    with tab_kw:
        c1, c2 = st.columns([3, 1])
        kw_input = c1.text_input("검색 키워드 입력 (쉼표 `,`로 조합, 예: `B라인, W44`)", value="")
        match_type = c2.radio("조건", options=["OR", "AND"], horizontal=True)
        if kw_input.strip():
            df = filter_by_keywords(raw_df, kw_input, match_all=(match_type == "AND"))
        st.dataframe(df, use_container_width=True)

    with tab_sql:
        sql_input = st.text_area("수동 SQL 쿼리 입력", value=auto_sql, height=80)
        if st.button("수동 SQL 실행"):
            try:
                df = run_sql_query(sql_input)
                st.dataframe(df, use_container_width=True)
            except Exception as e:
                st.error(f"오류: {e}")
        if DB_PATH.exists():
            with open(DB_PATH, "rb") as f:
                st.download_button("💾 자동 생성된 SQLite DB 파일(.db) 추출", data=f, file_name="internal_analysis.db")

    summary_df = df.describe(include="all").fillna("") if len(df) > 0 else raw_df.describe(include="all").fillna("")
    st.divider()

    # 2. 시각화 차트
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

    # 3. 다중 AI 모델 동시 비교 분석
    st.subheader("3️⃣ 자연어 질문 및 AI 에이전트 다중 비교 분석")
    question = st.text_input("💡 에이전트 분석 질문:", value="현재 조회된 데이터의 핵심 추세, 이상 항목, 원인 가설 및 조치 제안을 정리해 줘.")

    if st.button("🚀 에이전트 분석 실행 (선택 모델 동시 비교)", type="primary"):
        if not selected_models:
            st.warning("좌측 사이드바에서 최소 1개 이상의 모델을 선택해주세요.")
        else:
            agent_app = build_agent_graph()
            target_df = df if len(df) > 0 else raw_df
            data_info = f"데이터 건수: {len(target_df)}건\n컬럼: {list(target_df.columns)}\n데이터 내역:\n{target_df.to_string(index=False)}"

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

                        excel_bytes = export_analysis_to_excel(target_df, summary_df, res["report"], res["elapsed_time"], m_code)
                        st.download_button(
                            label=f"📥 {m_code} 엑셀 리포트 추출",
                            data=excel_bytes,
                            file_name=f"report_{m_code}.xlsx",
                            key=f"dl_{m_code}_{idx}"
                        )
