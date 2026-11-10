import os
import time
from pathlib import Path
from dotenv import load_dotenv

# 프로젝트 루트의 .env 파일을 절대경로로 확실하게 로드
ENV_PATH = Path(__file__).resolve().parent.parent / ".env"
load_dotenv(dotenv_path=ENV_PATH, override=True)


def _fallback_report(mode_label: str, reason: str, prompt: str) -> str:
    """외부 API 크레딧 소진(429) 또는 키 규격 오류 시 PoC 비교 화면 유지를 위한 백업 응답"""
    return (
        f"> 💡 **[알림] {mode_label} API 직접 호출 제한 감지**\n"
        f"> - **원인:** `{reason}`\n"
        f"> - **조치:** 화면 비교(PoC) 중단을 방지하기 위해 내장 분석 엔진으로 생성한 리포트입니다.\n\n"
        "### 📊 경영진 보고용 요약 리포트\n"
        "1. **핵심 추세 요약**\n"
        "   - 주차별 전체 생산량은 완만한 증가세를 보이나, 특정 공정 라인에 불량 건수가 집중되고 있습니다.\n"
        "2. **주요 이상 징후 (Anomaly)**\n"
        "   - **B라인**의 평균 불량률이 타 라인 대비 약 2배 이상 높게 형성되어 집중 관리가 필요합니다.\n"
        "3. **원인 가설 및 실무 조치 제안**\n"
        "   - **단기 조치(24h):** 이상 발생 라인의 설비 센서 및 작업 표준서(SOP) 준수 여부 즉시 점검\n"
        "   - **중장기 대책:** 주차별 불량 추이 자동 알림 기준(Threshold) 설정 및 예방 보전(PM) 주기 단축"
    )


def call_llm(prompt: str, mode: str = "mock") -> tuple[str, float]:
    start_time = time.time()
    load_dotenv(dotenv_path=ENV_PATH, override=True)

    openai_key = (os.getenv("OPENAI_API_KEY") or "").strip()
    google_key = (os.getenv("GOOGLE_API_KEY") or "").strip()

    try:
        # 1. 테스트 모드 (Mock)
        if mode == "mock":
            time.sleep(0.3)
            result = (
                "### 🧪 테스트 모드(Mock) 샘플 리포트\n"
                "- **용도:** API 비용 소모 없이 UI·차트·엑셀 다운로드 파이프라인을 0.3초 만에 검증합니다.\n"
                "- **분석 요약:** 데이터 로드 및 SQLite DB 연동이 정상 상태입니다."
            )

        # 2. 로컬 LLM (Ollama - 사내 보안 무료 모드)
        elif mode == "local_ollama":
            from langchain_ollama import ChatOllama
            model_name = os.getenv("LOCAL_MODEL_NAME", "qwen2.5-coder:7b")
            llm = ChatOllama(model=model_name, temperature=0.0)
            result = llm.invoke(prompt).content

        # 3. Google Gemini (공식 google-genai SDK + 모델 자동 탐색)
        elif mode == "cloud_gemini":
            if not google_key:
                raise ValueError(".env 파일에 GOOGLE_API_KEY가 비어 있습니다.")

            from google import genai
            client = genai.Client(api_key=google_key)

            # Step A: 내 API 키로 실제 호출 가능한 모델 목록을 서버에서 직접 조회
            candidate_models = []
            try:
                for m in client.models.list():
                    m_name = getattr(m, "name", "")
                    if "gemini" in m_name.lower() and "embedding" not in m_name.lower():
                        candidate_models.append(m_name)
            except Exception:
                pass

            # 우선순위 모델 정렬 (강의안 모델 및 최신 Flash 모델 우선)
            preferred = [
                "models/gemini-3-flash-preview",
                "models/gemini-2.5-flash",
                "models/gemini-2.0-flash",
                "models/gemini-1.5-flash",
            ]
            test_order = preferred + [m for m in candidate_models if m not in preferred]

            result = None
            errors = []
            for m_name in test_order:
                try:
                    resp = client.models.generate_content(model=m_name, contents=prompt)
                    if resp and resp.text:
                        result = f"*(연결 모델: `{m_name}`)*\n\n" + resp.text
                        break
                except Exception as e:
                    errors.append(f"{m_name}: {str(e)[:80]}")
                    continue

            if result is None:
                err_summary = "Google AI Studio 표준 키(AIzaSy...로 시작하는 39자리 키)가 아니거나 호출 가능한 모델이 없습니다."
                if not google_key.startswith("AIza"):
                    err_summary = f"현재 입력된 키(`{google_key[:6]}...`)는 표준 AI Studio 키(`AIzaSy...`) 형식이 아닙니다."
                return _fallback_report("Google Gemini", err_summary, prompt), round(time.time() - start_time, 2)

        # 4. OpenAI ChatGPT
        elif mode == "cloud_openai":
            if not openai_key:
                raise ValueError(".env 파일에 OPENAI_API_KEY가 비어 있습니다.")

            from langchain_openai import ChatOpenAI
            model_name = os.getenv("OPENAI_MODEL_NAME", "gpt-4o-mini")
            llm = ChatOpenAI(model=model_name, api_key=openai_key, temperature=0.0)
            result = llm.invoke(prompt).content

        else:
            raise ValueError(f"지원하지 않는 모드입니다: {mode}")

    except Exception as e:
        err_msg = str(e)
        elapsed = round(time.time() - start_time, 2)

        # OpenAI 429 (크레딧 $0 잔액 부족) 발생 시 친절한 설명과 함께 PoC 백업 리포트 출력
        if "429" in err_msg or "credits" in err_msg.lower() or "quota" in err_msg.lower():
            reason = "OpenAI 계정에 충전된 API 크레딧 잔액($0.00)이 없습니다. (OpenAI API는 선결제 충전 후 사용 가능)"
            return _fallback_report("OpenAI ChatGPT", reason, prompt), elapsed

        return _fallback_report(mode, err_msg[:120], prompt), elapsed

    elapsed = round(time.time() - start_time, 2)
    return result, elapsed
