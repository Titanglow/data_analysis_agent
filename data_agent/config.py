# data_agent/config.py
import os
from dotenv import load_dotenv

load_dotenv()

# UI 테마 및 차트 고대비 컬러 팔레트 (에메랄드 그린 기본 테마)
PRIMARY_COLOR = "#059669"
DISTINCT_PALETTE = [
    "#059669",  # 에메랄드 그린 (1위/기본)
    "#2563EB",  # 로얄 블루
    "#DC2626",  # 크림슨 레드 (주의/경고)
    "#D97706",  # 앰버 오렌지
    "#7C3AED",  # 바이올렛
    "#0891B2",  # 시안 블루
    "#DB2777",  # 핑크
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