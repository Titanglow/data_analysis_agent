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
    """Groq 서버에서 안정적인 텍스트 모델 목록을 조회합니다."""
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
                if "id" in m and not any(w in m["id"].lower() for w in ["whisper", "tts", "guard", "vision", "audio", "canopylabs", "tool-use", "llama-guard"])
            ]
            ordered = [m for m in preferred if m in avail]
            for m in avail:
                if m not in ordered:
                    ordered.append(m)
            return ordered if ordered else preferred
    except Exception:
        return preferred


# [RAG 추가] knowledge_context 파라미터를 추가로 받습니다.
def call_llm(prompt: str, mode: str = "local_ollama", knowledge_context: str = "") -> tuple[str, float]:
    """선택된 모드로 LLM을 호출하고 (응답 텍스트, 소요 시간 초)를 반환합니다."""
    load_dotenv(dotenv_path=ENV_PATH, override=True)
    t0 = time.perf_counter()

    if mode == "mock":
        time.sleep(0.15)
        return "테스트 응답입니다.", round(time.perf_counter() - t0, 2)

    # [RAG 핵심] 시스템 프롬프트에 사내 지식 문서를 강제 주입합니다.
    sys_prompt = "당신은 사내 데이터 분석 수석 전문가입니다. 원본 샘플 행 데이터를 통째로 나열하지 말고, 요청된 질문에 대해 간결하고 명확하게 답변하십시오."
    if knowledge_context:
        sys_prompt += (
            "\n\n[사내 규정 및 참고 지식 (RAG)]\n"
            "아래 제공된 사내 지식 문서를 반드시 참고하여 조치 제안이나 원인을 작성하십시오.\n"
            f"{knowledge_context}"
        )

    if mode == "cloud_groq":
        api_key = os.getenv("GROQ_API_KEY", "").strip()
        if not api_key:
            return "⚠️ `.env` 파일에 `GROQ_API_KEY`가 설정되지 않았습니다.", round(time.perf_counter() - t0, 2)
        try:
            from openai import OpenAI
            client = OpenAI(api_key=api_key, base_url="https://api.groq.com/openai/v1")
            candidates = _get_active_groq_models(api_key)
            last_err = None

            # 토큰 방어선
            safe_prompt = prompt if len(prompt) < 10000 else prompt[:10000] + "\n\n...[데이터 생략]"

            for m_name in candidates[:3]:
                try:
                    resp = client.chat.completions.create(
                        model=m_name,
                        messages=[
                            {"role": "system", "content": sys_prompt},
                            {"role": "user", "content": safe_prompt},
                        ],
                        temperature=0.2,
                        max_tokens=900,
                    )
                    txt = resp.choices[0].message.content or ""
                    if txt.strip():
                        header = f"*(연결 모델: Groq `{m_name}`)*\n\n"
                        return header + txt.strip(), round(time.perf_counter() - t0, 2)
                except Exception as e:
                    last_err = e
                    if any(err_code in str(e) for err_code in ["413", "429"]):
                        return f"⚠️ Groq 토큰 한도 초과: 검색 조건을 좁혀주세요.\n({last_err})", round(time.perf_counter() - t0, 2)
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

            # Gemini 프롬프트 병합
            gemini_prompt = f"{sys_prompt}\n\n[사용자 요청]\n{prompt}"

            for m_name in candidates[:4]:
                try:
                    resp = client.models.generate_content(model=m_name, contents=gemini_prompt)
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
        ollama_prompt = f"{sys_prompt}\n\n[사용자 요청]\n{prompt}"
        payload = json.dumps({
            "model": local_model,
            "prompt": ollama_prompt,
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