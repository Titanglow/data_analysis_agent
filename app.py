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
    analyze_color_cell_contexts,
    cluster_text_series_by_topic,
    shorten_long_label,
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
    .wf-card {
        border-radius: 10px;
        padding: 12px 14px;
        border: 1.5px solid #D1D5DB;
        background: #F9FAFB;
        height: 100%;
    }
    .wf-card-done {
        border-color: #059669;
        background: #ECFDF5;
    }
    .wf-card-wait {
        border-color: #9CA3AF;
        background: #F3F4F6;
        opacity: 0.85;
    }
    .wf-title {
        font-weight: 700;
        font-size: 13.5px;
        color: #065F46;
        margin-bottom: 4px;
    }
    .wf-file {
        font-family: monospace;
        font-size: 11.5px;
        color: #047857;
        background: #D1FAE5;
        padding: 2px 6px;
        border-radius: 4px;
        display: inline-block;
        margin-bottom: 6px;
    }
    .wf-status {
        font-size: 12px;
        color: #1F2937;
        line-height: 1.4;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

DISTINCT_PALETTE = ["#059669", "#2563EB", "#DC2626", "#D97706", "#7C3AED", "#0891B2", "#DB2777"]

st.title("📊 자연어 기반 사내 데이터·문서 통합 분석 AI Agent")

with st.expander("💡 [도입 목적] 왜 엑셀 수작업이나 일반 ChatGPT 대신 이 에이전트를 쓰나요? (클릭)", expanded=False):
    p_c1, p_c2, p_c3, p_c4 = st.columns(4)
    with p_c1:
        st.markdown(
            "**1️⃣ 사내 기밀 100% 보안 유지**\n"
            "- 외부 웹 ChatGPT에 사내 문서를 올리면 외부 서버에 유출·학습됩니다.\n"
            "- **Local Ollama 모드**는 사내 PC 내부에서만 돌아가 기밀 유출이 **0건**입니다."
        )
    with p_c2:
        st.markdown(
            "**2️⃣ AI 숫자 거짓말(환각) 차단**\n"
            "- 일반 AI에 파일을 통째로 주면 합계·비율을 틀리게 계산합니다.\n"
            "- **Python/SQL 엔진이 먼저 100% 정확히 계산**하고 AI는 해석만 수행합니다."
        )
    with p_c3:
        st.markdown(
            "**3️⃣ 엑셀·워드·폴라리온 통합 DB화**\n"
            "- 다중 시트 엑셀(`.xlsx`), 폴라리온 TC 절차서, **워드 보고서(`.docx`)**를 즉시 **SQLite DB**로 정제·변환합니다."
        )
    with p_c4:
        st.markdown(
            "**4️⃣ 주간보고 산출물 1분 자동화**\n"
            "- 매주 2~3시간 걸리던 데이터 필터링·차트 작성·보고서 초안 작업을 **3대 AI 교차 검증 + 마스터 엑셀(`.xlsx`) 추출**로 1분 만에 끝냅니다."
        )

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


def render_live_workflow_tracker(
    file_name: str,
    raw_count: int,
    filtered_count: int,
    aux_count: int,
    saved_results: dict,
):
    has_ai_done = len(saved_results) > 0
    if has_ai_done:
        times_str = ", ".join(f"{k.replace('cloud_', '').replace('local_', '')}: {v.get('elapsed_time', 0)}s" for k, v in saved_results.items())
        ai_badge = f"✅ <b>{len(saved_results)}개 AI 모델</b> 분석 완료!<br/>• ({times_str})"
        excel_badge = f"✅ <b>마스터 엑셀(.xlsx)</b> 패키징 완료<br/>• 종합비교 + 한글집계표 + 차트 내장"
    else:
        ai_badge = "⏳ 실행 대기 (3번 버튼 클릭 시 가동)<br/>• IF/ELSE 판정 + 4단락 리포트 병렬 작성"
        excel_badge = "⏳ AI 분석 후 자동 생성<br/>• 한글 집계표 + BarChart 내장"

    ai_card_cls = "wf-card wf-card-done" if has_ai_done else "wf-card wf-card-wait"

    with st.expander("⚡ [실시간 파이프라인 가동 현황] 각 파이썬(.py) 노드별 실시간 데이터 흐름 보기 (클릭)", expanded=True):
        cols = st.columns(5)
        steps = [
            (
                "1️⃣ 입력 & 정제 노드",
                "data_loader.py",
                "wf-card wf-card-done",
                f"✅ <b>{file_name[:18]}</b> 로드 완료<br/>• Polarion 개행(_x000D_) 정제<br/>• 보조시트/표 {aux_count}개 분리",
            ),
            (
                "2️⃣ SQLite DB & SQL 노드",
                "app.py (build_smart_sql)",
                "wf-card wf-card-done",
                f"✅ 전체 <b>{raw_count:,}건</b> DB 적재 완료<br/>• 현재 SQL 조건 추출: <b>{filtered_count:,}건</b>",
            ),
            (
                "3️⃣ 파이썬 고속 집계 노드",
                "analyzer_node.py",
                "wf-card wf-card-done",
                f"✅ <b>{filtered_count:,}건</b> 통계·클러스터링 완료<br/>• 숫자 환각 방지 선(先)계산 (0.01초)",
            ),
            (
                "4️⃣ 3대 AI 보고서 노드",
                "reporter_node.py + llm.py",
                ai_card_cls,
                ai_badge,