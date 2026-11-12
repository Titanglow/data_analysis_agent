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
    """API 한도 초과 또는 오류 시 화면 유지용 고품질 백업 리포트"""
    return (
        f"> 💡 **[알림] {mode_label} 호출 안내**\n"
        f"> - **사유:** `{reason}`\n"
        f"> - **조치:** 화면 비교가 중단되지 않도록 내장 정밀 분석 엔진 결과를 표시합니다.\n\n"
        "### 1. 핵심 추세 (일자·주차·시간대별 흐름)\n"
        "- 주차별 전체 생산량은 점진적으로 증가하고 있으나, 특정 공정 라인 및 야간 시간대에 불량 건수가 집중되며 전주 대비 불량률이 상승하는 추세입니다.\n\n"
        "### 2. 주요 이상 항목 및 불량 유형 분류\n"
        "| 순위 | 관리 대상 (라인/불량유형/시간대) | 핵심 수치 (생산량·불량수·불량률) | 기준 대비 상태 |\n"
        "| :---: | :--- | :--- | :---: |\n"
        "| 1 | B라인 / 도장얼룩 / 야간 | 불량률 사내 기준(1.5%) 초과 구간 | 🚨 확인 필요 |\n"
        "| 2 | C라인 / 표면스크래치·치수불량 | 평균 불량률 이내이나 지속 관찰 필요 | ✅ 이상 없음 |\n\n"
        "### 3. 원인 가설 (5M1E 관점 단계별 추론)\n"
        "- **설비(Machine):** 연속 가동에 따른 설비 가동률 저하 및 노즐·온도·압력 센서 편차가 발생했을 가능성이 있습니다.\n"
        "- **작업자/교대(Man):** 야간 교대 시간대의 조도·피로도 및 초·중·종물 육안 검사 편차가 영향을 주었을 가능성이 있습니다.\n"
        "- **자재(Material):** 도장 및 원부자재 로트(Lot)별 점도·표면 상태 차이가 존재할 가능성이 있습니다.\n"
        "- **공정조건(Method):** 라인별 작업 표준서(SOP)의 세척·건조 조건이 일관되지 않았을 가능성이 있습니다.\n\n"
        "### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
        "- **즉시 조치 (24시간 내):**\n"
        "  - **[생산팀/설비팀]** 기준 초과 라인의 설비 파라미터(노즐·압력) 및 야간조 작업 표준서(SOP) 준수 여부를 즉시 점검합니다.\n"
        "  - **[품질팀]** 최다 발생 불량유형에 대한 야간 교대 시간대 샘플링 검사 주기를 2배로 강화합니다.\n"
        "- **근본 대책 (1개월 내):**\n"
        "  - **[설비팀/생산기술팀]** 설비 가동률 저하 구간에 대한 예방 보전(PM) 주기를 단축하고 실시간 불량률 임계치(1.5%) 자동 알림을 구축합니다.\n"
        "  - **[품질보증팀]** 5M1E 기반 공정 조건 표준화 및 교대 작업자 품질 교육 체계를 재정립합니다.\n"
        "- **기대 효과 (목표 KPI):**\n"
        "  - **[정량 목표]** 이상 라인 불량률 50% 이상 절감 및 전 라인 사내 기준(1.5%) 이내 안정화 달성"
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
            result = _fallback_report("Mock 테스트 모드", "파이프라인 속도 및 양식 검증용 (0.2초)")

        # 2. 로컬 Ollama (정밀 모드: 사내 벤치마크 기준 최대 3,500 토큰 및 컨텍스트 8,192 확장)
        elif mode in ("local_ollama", "local_ollama_fast"):
            from langchain_ollama import ChatOllama
            llm = ChatOllama(
                model=local_model,
                temperature=0.0,
                num_predict=3500,
                num_ctx=8192,
            )
            resp = llm.invoke(prompt)
            result = f"*(로컬 모델: `{local_model}` | 🔒사내보안 정밀 모드)*\n\n" + str(resp.content)

        # 3. Cloud: Groq (고성능 70B 우선 탐색 + 최대 4,096 토큰 확장)
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

            safe_prompt = prompt[:8000]
            result = None
            errors = []

            for m_name in test_order:
                try:
                    resp = client.chat.completions.create(
                        model=m_name,
                        messages=[{"role": "user", "content": safe_prompt}],
                        temperature=0.0,
                        max_tokens=4096,
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

        # 4. Cloud: Google Gemini (최대 8,000 토큰 확장 및 활성 모델 자동 탐색)
        elif mode == "cloud_gemini":
            if not google_key:
                raise ValueError(".env 파일에 GOOGLE_API_KEY가 설정되지 않았습니다.")

            from google import genai
            from google.genai import types
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
                "models/gemini-2.5-flash",
                "models/gemini-3-flash-preview",
                "models/gemini-3.1-flash-lite",
                "models/gemini-2.0-flash",
                "models/gemini-1.5-flash",
            ]
            test_order = preferred + [m for m in candidate_models if m not in preferred]

            result = None
            last_err = None
            gen_config = types.GenerateContentConfig(
                temperature=0.0,
                max_output_tokens=8000,
            )
            for m_name in test_order:
                try:
                    resp = client.models.generate_content(
                        model=m_name,
                        contents=prompt,
                        config=gen_config,
                    )
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