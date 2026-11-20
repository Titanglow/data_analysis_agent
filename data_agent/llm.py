# data_agent/llm.py
import os
import time
import urllib.request
import json
from pathlib import Path
from dotenv import load_dotenv

PROJECT_ROOT = Path(__file__).resolve().parent.parent
ENV_PATH = PROJECT_ROOT / ".env"
load_dotenv(dotenv_path=ENV_PATH, override=True)


def get_installed_ollama_models() -> list[str]:
    """내 PC의 로컬 Ollama 서버에서 설치된 모델 목록을 실시간 조회합니다."""
    base_url = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
    default_model = os.getenv("LOCAL_MODEL_NAME", "qwen2.5-coder:7b")
    try:
        req = urllib.request.Request(f"{base_url}/api/tags", method="GET")
        with urllib.request.urlopen(req, timeout=2.0) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            models = [m["name"] for m in data.get("models", []) if "name" in m]
            if models:
                if default_model in models:
                    models.remove(default_model)
                    models.insert(0, default_model)
                return models
    except Exception:
        pass
    return [default_model, "qwen3:8b", "llama3.1:8b", "qwen2.5:3b"]


def _get_active_groq_models(api_key: str) -> list[str]:
    """
    Groq 서버에서 현재 계정이 사용 가능한 최신 활성 텍스트 모델 목록을 조회합니다.
    (약관 동의가 필요한 모델이나 오디오/비전 모델은 제외하고 안정적인 텍스트 모델만 반환)
    """
    preferred = [
        "llama-3.3-70b-versatile",
        "llama-3.1-8b-instant",
    ]
    try:
        req = urllib.request.Request(
            "https://api.groq.com/openai/v1/models",
            headers={"Authorization": f"Bearer {api_key}", "User-Agent": "Mozilla/5.0"},
        )
        with urllib.request.urlopen(req, timeout=3.0) as resp:
            data = json.loads(resp.read().decode("utf-8"))
            avail = [
                m["id"]
                for m in data.get("data", [])
                if "id" in m and not any(w in m["id"].lower() for w in ["whisper", "tts", "guard", "vision", "audio", "canopylabs", "tool-use"])
            ]
            ordered = [m for m in preferred if m in avail]
            for m in avail:
                if m not in ordered:
                    ordered.append(m)
            return ordered if ordered else preferred
    except Exception:
        return preferred


def call_llm(prompt: str, mode: str = "local_ollama") -> tuple[str, float]:
    """선택된 모드로 LLM을 호출하고 (응답 텍스트, 소요 시간 초)를 반환합니다."""
    load_dotenv(dotenv_path=ENV_PATH, override=True)
    t0 = time.perf_counter()

    if mode == "mock":
        time.sleep(0.15)
        sample = (
            "### 1. 전체 검증 데이터 요약 및 사내 기준 판정 현황\n"
            "- 파이썬 정밀 집계 결과 전체 데이터 건수 및 주요 분류 항목이 정상 확인되었습니다.\n\n"
            "### 2. 실제 분류 컬럼별 분포 & 주요 시나리오 그룹핑 심층 분석\n"
            "- 상위 집중 분류 항목과 핵심 시나리오 그룹의 비중(%)이 사내 관리 기준에 맞춰 분석되었습니다.\n\n"
            "### 3. 데이터 품질 리스크 및 취약 구간 진단\n"
            "- 미기입(공란/N/A) 항목 및 특정 모듈 편중 구간에 대한 보완이 필요합니다.\n\n"
            "### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
            "- **즉시 조치 (24시간 내 / 담당팀)**: 기준 초과 미기입 항목 및 주의 대상 우선 점검\n"
            "- **근본 대책 (1개월 내 / 품질관리 파트)**: 취약 구간 시나리오 확충 및 템플릿 표준화"
        )
        return sample, round(time.perf_counter() - t0, 2)

    if mode == "cloud_groq":
        api_key = os.getenv("GROQ_API_KEY", "").strip()
        if not api_key:
            return "⚠️ `.env` 파일에 `GROQ_API_KEY`가 설정되지 않았습니다.", round(time.perf_counter() - t0, 2)
        try:
            from openai import OpenAI
            client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")
            candidates = _get_active_groq_models(api_key)
            last_err = None

            # [다이어트 추가 안전장치] 만약 파이썬에서 줄이는 걸 깜빡했더라도 여기서 물리적으로 잘라버림
            safe_prompt = prompt if len(prompt) < 10000 else prompt[:10000] + "\n...[데이터 길이 초과로 요약됨]"

            for m_name in candidates[:3]:
                try:
                    resp = client.chat.completions.create(
                        model=m_name,
                        messages=[
                            {
                                "role": "system",
                                "content": "당신은 사내 데이터 분석 수석 전문가입니다. 원본 샘플 행 데이터를 통째로 나열하지 말고, 요청된 질문에 대해 간결하고 명확하게 답변하십시오.",
                            },
                            {"role": "user", "content": safe_prompt},
                        ],
                        temperature=0.2,
                        max_tokens=900,  # 출력 토큰 제한 하향
                    )
                    txt = resp.choices[0].message.content or ""
                    if txt.strip():
                        header = f"*(연결 모델: Groq `{m_name}`)*\n\n"
                        return header + txt.strip(), round(time.perf_counter() - t0, 2)
                except Exception as e:
                    last_err = e
                    # 413, 429 에러(토큰 초과) 발생 시 즉시 중단 및 안내 메시지 반환
                    if any(err_code in str(e) for err_code in ["413", "429"]):
                        return f"⚠️ Groq 토큰 한도 초과 에러: 입력 데이터가 너무 많습니다. 조회 조건을 더 좁혀서(예: '상위 5건') 검색해 주세요.\n({last_err})", round(time.perf_counter() - t0, 2)
                    continue
            return f"⚠️ Groq API 호출 오류: {last_err}", round(time.perf_counter() - t0, 2)
        except Exception as e:
            return f"⚠️ Groq 초기화 오류: {e}", round(time.perf_counter() - t0, 2)

    if mode == "cloud_gemini":
        api_key = os.getenv("GOOGLE_API_KEY", "").strip() or os.getenv("GEMINI_API_KEY", "").strip()
        if not api_key:
            return "⚠️ `.env` 파일에 `GOOGLE_API_KEY`가 설정되지 않았습니다.", round(time.perf_counter() - t0, 2)
        try:
            from google import genai
            client = genai.Client(api_key=api_key)
            preferred_gemini = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash"]
            avail_gemini = []
            try:
                for m in client.models.list():
                    m_id = getattr(m, "name", "").replace("models/", "")
                    if "gemini" in m_id and "flash" in m_id and "vision" not in m_id:
                        avail_gemini.append(m_id)
            except Exception:
                pass
            candidates = [m for m in preferred_gemini if m in avail_gemini] + [
                m for m in avail_gemini if m not in preferred_gemini
            ]
            if not candidates:
                candidates = preferred_gemini

            last_err = None
            for m_name in candidates[:4]:
                try:
                    resp = client.models.generate_content(model=m_name, contents=prompt)
                    txt = getattr(resp, "text", "") or ""
                    if txt.strip():
                        header = f"*(연결 모델: Google `{m_name}`)*\n\n"
                        return header + txt.strip(), round(time.perf_counter() - t0, 2)
                except Exception as e:
                    last_err = e
                    continue
            return f"⚠️ Gemini API 호출 오류: {last_err}", round(time.perf_counter() - t0, 2)
        except Exception as e:
            return f"⚠️ Gemini 초기화 오류: {e}", round(time.perf_counter() - t0, 2)

    base_url = os.getenv("OLLAMA_BASE_URL", "http://localhost:11434")
    local_model = os.getenv("LOCAL_MODEL_NAME", "qwen2.5-coder:7b")
    try:
        payload = json.dumps({
            "model": local_model,
            "prompt": prompt,
            "stream": False,
            "options": {
                "temperature": 0.2,
                "num_predict": 2200,
                "num_ctx": 4096,
            },
        }).encode("utf-8")
        req = urllib.request.Request(
            f"{base_url}/api/generate",
            data=payload,
            headers={"Content-Type": "application/json"},
            method="POST",
        )
        with urllib.request.urlopen(req, timeout=180.0) as resp:
            res_json = json.loads(resp.read().decode("utf-8"))
            txt = res_json.get("response", "").strip()
            header = f"*(로컬 모델: Ollama `{local_model}`)*\n\n"
            return header + txt, round(time.perf_counter() - t0, 2)
    except Exception as e:
        return f"⚠️ 로컬 Ollama 호출 오류 ({local_model}): {e}", round(time.perf_counter() - t0, 2)