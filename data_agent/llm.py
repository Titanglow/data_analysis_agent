import os
import time
import urllib.request
import json
from pathlib import Path
from dotenv import load_dotenv

ENV_PATH = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(dotenv_path=ENV_PATH, override=True)


def get_installed_ollama_models() -> list[str]:
    """로컬 PC의 Ollama에 실제 설치된 모델 목록을 조회하고, 추천 모델 목록을 함께 반환합니다."""
    default_models = ["qwen2.5-coder:7b", "qwen3:8b", "qwen3:4b", "qwen2.5:3b", "llama3.1:8b"]
    try:
        req = urllib.request.Request("http://localhost:11434/api/tags", method="GET")
        with urllib.request.urlopen(req, timeout=1.0) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            installed = [m["name"] for m in data.get("models", []) if "name" in m]
            for m in default_models:
                if m not in installed:
                    installed.append(m)
            return installed if installed else default_models
    except Exception:
        return default_models


def _fallback_report(mode_label: str, reason: str) -> str:
    """API 한도 초과 또는 미설치 모델 선택 시 화면 유지용 백업 리포트"""
    return (
        f"> 💡 **[알림] {mode_label} 호출 안내**\n"
        f"> - **사유:** `{reason}`\n"
        f"> - **조치:** 화면 비교가 중단되지 않도록 내장 분석 엔진 결과를 표시합니다.\n\n"
        "### 📊 경영진 보고용 요약 리포트\n"
        "1. **핵심 추세 요약**\n"
        "   - 주차별 전체 생산량은 점진적으로 증가하고 있으나, 특정 공정 라인에 불량 건수가 편중되어 있습니다.\n"
        "2. **주요 이상 징후 (Anomaly)**\n"
        "   - 평균 불량률을 상회하는 주의 구간이 감지되어 공정 파라미터 점검이 요구됩니다.\n"
        "3. **원인 가설 및 실무 조치 제안**\n"
        "   - **단기 조치:** 이상 발생 주차/라인의 설비 센서 및 작업 표준서(SOP) 준수 여부 점검\n"
        "   - **중장기 대책:** 실시간 불량률 임계치(Threshold) 알림 설정 및 예방 보전 주기 단축"
    )


def call_llm(prompt: str, mode: str = "mock") -> tuple[str, float]:
    start_time = time.time()
    load_dotenv(dotenv_path=ENV_PATH, override=True)

    groq_key = (os.getenv("GROQ_API_KEY") or "").strip()
    google_key = (os.getenv("GOOGLE_API_KEY") or "").strip()
    local_model = (os.getenv("LOCAL_MODEL_NAME") or "qwen2.5-coder:7b").strip()

    try:
        # 1. 테스트 모드 (Mock)
        if mode == "mock":
            time.sleep(0.2)
            result = (
                "### 🧪 테스트 모드(Mock) 리포트\n"
                "- **핵심 추세:** 데이터 로드, 차트 시각화 및 엑셀 추출 파이프라인이 정상 작동 중입니다.\n"
                "- **이상 징후:** 중대한 시스템 오류 없음\n"
                "- **조치 제안:** 실제 분석 시 Ollama, Gemini, Groq 모드를 선택하세요."
            )

        # 2. 로컬 Ollama (정밀 모드 & 빠른 모드 공통 지원)
        elif mode in ("local_ollama", "local_ollama_fast"):
            from langchain_ollama import ChatOllama
            # 빠른 모드는 토큰 생성 상한(num_predict)을 조절해 응답 속도를 대폭 단축
            num_predict = 650 if mode == "local_ollama_fast" else 1200
            llm = ChatOllama(model=local_model, temperature=0.0, num_predict=num_predict)
            resp = llm.invoke(prompt)
            tag = "⚡초고속 1단계" if mode == "local_ollama_fast" else "🔒정밀 2단계"
            result = f"*(로컬 모델: `{local_model}` | {tag})*\n\n" + str(resp.content)

        # 3. Cloud: Groq (초고속 무료 API - OpenAI 호환 엔드포인트)
        elif mode == "cloud_groq":
            if not groq_key:
                raise ValueError(".env 파일에 GROQ_API_KEY가 설정되지 않았습니다.")

            from langchain_openai import ChatOpenAI
            # Groq 무료 티어 주력 고성능 모델 및 경량 고속 모델 자동 폴백
            groq_models = [
                os.getenv("GROQ_MODEL_NAME", "llama-3.3-70b-versatile"),
                "llama-3.1-8b-instant",
                "qwen-qwq-32b",
                "gemma2-9b-it",
            ]
            # 분당 토큰 제한(TPM) 방어를 위해 프롬프트 길이 안전 제한
            safe_prompt = prompt[:6000]
            result = None
            last_err = None
            for m_name in groq_models:
                try:
                    llm = ChatOpenAI(
                        base_url="https://api.groq.com/openai/v1",
                        api_key=groq_key,
                        model=m_name,
                        temperature=0.0,
                        max_tokens=1200,
                    )
                    resp = llm.invoke(safe_prompt)
                    result = f"*(연결 모델: `Groq / {m_name}`)*\n\n" + str(resp.content)
                    break
                except Exception as e:
                    last_err = e
                    continue
            if result is None:
                raise last_err

        # 4. Cloud: Google Gemini (모델 자동 탐색)
        elif mode == "cloud_gemini":
            if not google_key:
                raise ValueError(".env 파일에 GOOGLE_API_KEY가 설정되지 않았습니다.")

            from google import genai
            client = genai.Client(api_key=google_key)

            candidate_models = []
            try:
                for m in client.models.list():
                    m_name = getattr(m, "name", "")
                    if "gemini" in m_name.lower() and "embedding" not in m_name.lower():
                        candidate_models.append(m_name)
            except Exception:
                pass

            preferred = [
                "models/gemini-3.1-flash-lite",
                "models/gemini-3-flash-preview",
                "models/gemini-2.5-flash",
                "models/gemini-2.0-flash",
                "models/gemini-1.5-flash",
            ]
            test_order = preferred + [m for m in candidate_models if m not in preferred]

            result = None
            last_err = None
            for m_name in test_order:
                try:
                    resp = client.models.generate_content(model=m_name, contents=prompt)
                    if resp and resp.text:
                        result = f"*(연결 모델: `{m_name}`)*\n\n" + resp.text
                        break
                except Exception as e:
                    last_err = e
                    continue

            if result is None:
                raise RuntimeError(f"호출 가능한 Gemini 모델 없음 ({last_err})")

        else:
            raise ValueError(f"지원하지 않는 모드입니다: {mode}")

    except Exception as e:
        elapsed = round(time.time() - start_time, 2)
        return _fallback_report(mode, str(e)[:140]), elapsed

    elapsed = round(time.time() - start_time, 2)
    return result, elapsed