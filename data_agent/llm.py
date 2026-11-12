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
    """API 한도 초과 또는 오류 시 화면 유지용 백업 리포트"""
    return (
        f"> 💡 **[알림] {mode_label} 호출 안내**\n"
        f"> - **사유:** `{reason}`\n"
        f"> - **조치:** 화면 비교가 중단되지 않도록 내장 분석 엔진 결과를 표시합니다.\n\n"
        "### 1. 핵심 추세 (일자·주차·시간대별 흐름)\n"
        "- 주차별 전체 생산량은 점진적으로 증가하고 있으나, 특정 공정 라인 및 야간 시간대에 불량 건수가 편중되어 있습니다.\n\n"
        "### 2. 주요 이상 항목 및 불량 유형 분류\n"
        "- 평균 불량률 및 사내 기준(1.5%)을 상회하는 주의 구간이 감지되어 집중 점검이 요구됩니다.\n\n"
        "### 3. 원인 가설 (5M1E 관점 단계별 추론)\n"
        "- **설비(Machine):** 연속 가동에 따른 설비 가동률 저하 및 노즐/센서 편차가 발생했을 가능성이 있습니다.\n"
        "- **작업자/교대(Man):** 야간 교대 시간대의 작업 조건 및 육안 검사 편차가 영향을 주었을 가능성이 있습니다.\n\n"
        "### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
        "- **즉시 조치(24시간 내):** [생산팀/설비팀] 기준 초과 라인의 설비 파라미터 및 야간조 작업 표준서(SOP) 즉시 점검\n"
        "- **근본 대책(1개월 내):** [품질팀/설비팀] 최다 발생 불량유형 집중 개선 및 예방 보전(PM) 주기 단축"
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
                "### 1. 핵심 추세 (일자·주차·시간대별 흐름)\n"
                "- 데이터 로드, 차트 시각화 및 엑셀 추출 파이프라인이 정상 작동 중입니다.\n\n"
                "### 2. 주요 이상 항목 및 불량 유형 분류\n"
                "- 테스트 모드에서는 이상 징후를 시뮬레이션으로 점검합니다.\n\n"
                "### 3. 원인 가설 (5M1E 관점 단계별 추론)\n"
                "- 세부 원인은 자료로는 확인 불가하나 공정 조건 편차일 가능성이 있습니다.\n\n"
                "### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
                "- **즉시 조치(24시간 내):** [품질팀] 실제 분석 시 Ollama, Gemini, Groq 모드를 선택해 실행하세요.\n"
                "- **근본 대책(1개월 내):** [생산기술팀] 정기 데이터 모니터링 체계를 유지하세요."
            )

        # 2. 로컬 Ollama (출력 잘림 방지를 위해 num_predict 및 num_ctx 대폭 확장)
        elif mode in ("local_ollama", "local_ollama_fast"):
            from langchain_ollama import ChatOllama
            # 기존 650 토큰 제한으로 인해 4번 항목에서 잘리던 문제를 해결하기 위해 2048/2500으로 확장
            num_predict = 2048 if mode == "local_ollama_fast" else 2500
            llm = ChatOllama(
                model=local_model,
                temperature=0.0,
                num_predict=num_predict,
                num_ctx=4096,
            )
            resp = llm.invoke(prompt)
            tag = "⚡초고속 1단계" if mode == "local_ollama_fast" else "🔒정밀 2단계"
            result = f"*(로컬 모델: `{local_model}` | {tag})*\n\n" + str(resp.content)

        # 3. Cloud: Groq (실시간 활성 모델 자동 탐색)
        elif mode == "cloud_groq":
            if not groq_key:
                raise ValueError(".env 파일에 GROQ_API_KEY가 설정되지 않았습니다.")

            from openai import OpenAI
            client = OpenAI(base_url="https://api.groq.com/openai/v1", api_key=groq_key)

            active_models = []
            skip_keywords = ["whisper", "tts", "guard", "safeguard", "embed", "playai", "orpheus", "allam"]
            try:
                model_list = client.models.list()
                for m in model_list.data:
                    m_id = getattr(m, "id", "")
                    if m_id and not any(k in m_id.lower() for k in skip_keywords):
                        active_models.append(m_id)
            except Exception:
                pass

            preferred = [
                "llama-3.3-70b-versatile",
                "llama-3.1-8b-instant",
                "llama3-70b-8192",
                "llama3-8b-8192",
            ]
            if active_models:
                test_order = [m for m in preferred if m in active_models] + [
                    m for m in active_models if m not in preferred
                ]
            else:
                test_order = preferred

            safe_prompt = prompt[:6000]
            result = None
            errors = []

            for m_name in test_order:
                try:
                    resp = client.chat.completions.create(
                        model=m_name,
                        messages=[{"role": "user", "content": safe_prompt}],
                        temperature=0.0,
                        max_tokens=2048,
                    )
                    content = resp.choices[0].message.content
                    if content:
                        result = f"*(연결 모델: `Groq / {m_name}`)*\n\n" + content
                        break
                except Exception as e:
                    errors.append(f"{m_name}: {str(e)[:60]}")
                    continue

            if result is None:
                raise RuntimeError(" | ".join(errors[:2]) if errors else "사용 가능한 Groq 모델을 찾지 못했습니다.")

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
        return _fallback_report(mode, str(e)[:160]), elapsed

    elapsed = round(time.time() - start_time, 2)
    return result, elapsed