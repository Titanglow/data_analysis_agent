# data_agent/config.py
import os
from dotenv import load_dotenv

load_dotenv()

# 기본 에메랄드 테마 색상 및 차트 팔레트
PRIMARY_COLOR = "#059669"
DISTINCT_PALETTE = [
    "#059669",  # 에메랄드 그린
    "#2563EB",  # 로얄 블루
    "#DC2626",  # 크림슨 레드
    "#D97706",  # 앰버 오렌지
    "#7C3AED",  # 바이올렛
    "#0891B2",  # 시안 블루
    "#DB2777",  # 핑크
]

# 노션(Notion) 데이터베이스 전용 차트 팔레트
NOTION_PALETTE = [
    "#37352F",  # 노션 차콜 블랙
    "#0F7B6C",  # 노션 틸 그린
    "#2EAADC",  # 노션 블루
    "#D9730D",  # 노션 앰버
    "#9065B0",  # 노션 퍼플
    "#E03E3E",  # 노션 코랄 레드
    "#64473A",  # 노션 브라운
]

# 사내 품질 관리 기본 임계치 (사이드바에서 실시간 변경 가능)
DEFAULT_DEFECT_THRESHOLD = 1.5

# 로컬 및 클라우드 모델 기본 설정
DEFAULT_LOCAL_MODEL = os.getenv("LOCAL_MODEL_NAME", "qwen2.5-coder:7b")
OLLAMA_BASE_URL = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")

MODEL_LABELS = {
    "cloud_groq": "⚡ Cloud: Groq (초고속 무료 API)",
    "cloud_gemini": "☁️ Cloud: Gemini (Google 무료 API)",
    "local_ollama": "🔒 Local: Ollama (사내보안 정밀 모드)",
    "mock": "🧪 테스트 모드 (Mock·0.2초 검증)",
}


def get_theme_css(theme_mode: str = "emerald", popup_width_css: str = "min(620px, 48vw)") -> str:
    """
    사이드바에서 선택한 UI 테마('emerald' 또는 'notion')에 맞춰
    전체 대시보드 및 우측 플로팅 팝업 채팅창의 CSS 스타일을 반환합니다.
    """
    if theme_mode == "notion":
        return f"""
        <style>
        :root {{
            --primary-color: #37352F !important;
        }}
        /* 전체 배경 및 노션 본문 서체 */
        .stApp {{
            background-color: #FFFFFF !important;
            color: #37352F !important;
            font-family: ui-sans-serif, -apple-system, BlinkMacSystemFont, "Segoe UI", "Malgun Gothic", "Apple SD Gothic Neo", sans-serif !important;
        }}
        /* 노션 특유의 웜그레이 사이드바 (#F7F6F3) */
        section[data-testid="stSidebar"] {{
            background-color: #F7F6F3 !important;
            border-right: 1px solid #E9E9E7 !important;
        }}
        section[data-testid="stSidebar"] * {{
            color: #37352F !important;
        }}
        /* 제목 타이포그래피 (노션 페이지 헤더 스타일) */
        h1, h2, h3, h4 {{
            color: #37352F !important;
            font-weight: 700 !important;
            letter-spacing: -0.02em !important;
        }}
        /* 멀티셀렉트 태그: 노션 파스텔 그린 태그 스타일 */
        span[data-baseweb="tag"],
        div[data-baseweb="select"] span[data-baseweb="tag"],
        [data-testid="stMultiSelect"] span[data-baseweb="tag"] {{
            background-color: #DBEDDB !important;
            color: #1C3829 !important;
            border: 1px solid #C7E2D4 !important;
            border-radius: 4px !important;
            font-weight: 600 !important;
        }}
        span[data-baseweb="tag"] span {{
            color: #1C3829 !important;
        }}
        /* 주요 실행 버튼: 노션 다크 차콜 버튼 */
        button[kind="primary"], [data-testid="stBaseButton-primary"] {{
            background-color: #37352F !important;
            border: 1px solid #37352F !important;
            color: #FFFFFF !important;
            border-radius: 6px !important;
            font-weight: 600 !important;
            box-shadow: 0 1px 2px rgba(15, 15, 15, 0.08) !important;
        }}
        button[kind="primary"]:hover, [data-testid="stBaseButton-primary"]:hover {{
            background-color: #262521 !important;
            border-color: #262521 !important;
        }}
        /* 일반 보조 버튼: 노션 블록 버튼 */
        button[kind="secondary"], [data-testid="stBaseButton-secondary"] {{
            background-color: #FFFFFF !important;
            border: 1px solid #E3E2E0 !important;
            color: #37352F !important;
            border-radius: 6px !important;
            font-weight: 500 !important;
        }}
        button[kind="secondary"]:hover, [data-testid="stBaseButton-secondary"]:hover {{
            background-color: #F7F6F3 !important;
            border-color: #D3D1CB !important;
        }}
        /* KPI Metric 카드: 노션 갤러리 뷰 카드 */
        div[data-testid="stMetric"] {{
            background-color: #FFFFFF !important;
            border: 1px solid #E9E9E7 !important;
            border-radius: 8px !important;
            padding: 12px 16px !important;
            box-shadow: 0 1px 2px rgba(15, 15, 15, 0.03) !important;
        }}
        /* 알림 박스(st.info / st.success / st.warning): 노션 콜아웃(Callout) 블록 스타일 */
        div[data-testid="stAlert"] {{
            background-color: #F7F6F3 !important;
            border: 1px solid #E9E9E7 !important;
            border-radius: 6px !important;
            color: #37352F !important;
        }}
        /* Expander(토글 블록) 및 표 테두리 */
        details[data-testid="stExpander"] {{
            background-color: #FFFFFF !important;
            border: 1px solid #E9E9E7 !important;
            border-radius: 6px !important;
        }}
        hr {{
            border-color: #E9E9E7 !important;
        }}
        /* 🗗 노션 스타일 플로팅 팝업 채팅창 */
        div[data-testid="stVerticalBlock"]:has(#floating-chat-marker):not(:has(div[data-testid="stVerticalBlock"] #floating-chat-marker)) {{
            position: fixed !important;
            bottom: 22px !important;
            right: 26px !important;
            width: {popup_width_css} !important;
            max-height: 86vh !important;
            overflow-y: auto !important;
            overflow-x: hidden !important;
            background-color: #FFFFFF !important;
            border: 1.5px solid #D3D1CB !important;
            border-radius: 12px !important;
            box-shadow: 0 16px 40px rgba(15, 15, 15, 0.18) !important;
            z-index: 999990 !important;
            padding: 14px 18px 18px 18px !important;
        }}
        </style>
        """

    # 기본 에메랄드 대시보드 테마
    return f"""
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
    /* 🗗 에메랄드 스타일 플로팅 팝업 채팅창 */
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
    """