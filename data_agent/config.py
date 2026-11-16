import os
from pathlib import Path
from dotenv import load_dotenv

# 프로젝트 루트의 .env 파일 로드
PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"
load_dotenv(dotenv_path=ENV_PATH, override=True)

# 1. 데이터베이스 기본 설정
DB_PATH = PROJECT_ROOT / "internal_analysis.db"
TABLE_NAME = "internal_data"

# 2. 사내 품질 판정 기준값 (IF/ELSE 분기 기준 - 기본 1.5%)
DEFAULT_DEFECT_THRESHOLD = float(os.getenv("DEFECT_THRESHOLD", "1.5"))

# 3. 멀티 LLM 모델 라벨 및 기본 설정
MODEL_LABELS = {
    "cloud_groq": "⚡ Cloud: Groq (초고속 무료 API)",
    "cloud_gemini": "☁️ Cloud: Gemini (Google 무료 API)",
    "local_ollama": "🔒 Local: Ollama (사내보안 정밀 모드)",
    "mock": "🧪 테스트 모드 (Mock·0.2초 검증)",
}

DEFAULT_LOCAL_MODEL = os.getenv("LOCAL_MODEL_NAME", "qwen2.5-coder:7b")

# 4. 대시보드 공통 컬러 팔레트 (에메랄드 그린 테마 및 차트 구분색)
PRIMARY_COLOR = "#059669"
DISTINCT_PALETTE = [
    "#059669",  # 에메랄드 그린 (1위/정상)
    "#2563EB",  # 로열 블루
    "#DC2626",  # 코랄 레드 (주의/강조)
    "#D97706",  # 앰버 오렌지
    "#7C3AED",  # 바이올렛 퍼플
    "#0891B2",  # 시안 블루
    "#DB2777",  # 핑크
]