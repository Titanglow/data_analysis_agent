import os
import time
from dotenv import load_dotenv

load_dotenv(override=True)  # 환경변수 강제 갱신


class MockLLM:
    def invoke(self, prompt: str) -> str:
        time.sleep(0.4)
        return (
            "### 📊 데이터 분석 요약 리포트 (테스트 모드)\n"
            "- **핵심 추세:** 업로드된 데이터의 수치 변동 폭이 안정적입니다.\n"
            "- **이상 징후:** 특정 구간에서 평균치를 상회하는 수치가 관측되었습니다.\n"
            "- **조치 제안:** 해당 구간의 원인을 점검하고 모니터링을 강화하세요."
        )


def call_llm(prompt: str, mode: str = "mock") -> tuple[str, float]:
    start_time = time.time()

    openai_key = os.getenv("OPENAI_API_KEY")
    google_key = os.getenv("GOOGLE_API_KEY")

    try:
        if mode == "mock":
            llm = MockLLM()
            result = llm.invoke(prompt)

        elif mode == "local_ollama":
            from langchain_ollama import ChatOllama
            model_name = os.getenv("LOCAL_MODEL_NAME", "qwen2.5-coder:7b")
            llm = ChatOllama(model=model_name, temperature=0.0)
            result = llm.invoke(prompt).content

        elif mode == "cloud_gemini":
            if not google_key:
                raise ValueError("GOOGLE_API_KEY가 설정되지 않았습니다. .env 파일을 확인해주세요.")
            from langchain_google_genai import ChatGoogleGenerativeAI
            # AQ 계정 및 표준 API와 호환되는 가장 안정적인 모델 리스트
            gemini_models = ["gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.0-flash"]
            last_err = None
            result = None
            for m_name in gemini_models:
                try:
                    llm = ChatGoogleGenerativeAI(model=m_name, google_api_key=google_key, temperature=0.0)
                    resp = llm.invoke(prompt)
                    result = str(resp.content) if not isinstance(resp.content, list) else "\n".join(str(x) for x in resp.content)
                    break
                except Exception as e:
                    last_err = e
                    continue
            if result is None:
                raise last_err

        elif mode == "cloud_openai":
            if not openai_key:
                raise ValueError("OPENAI_API_KEY가 설정되지 않았습니다. .env 파일을 확인해주세요.")
            from langchain_openai import ChatOpenAI
            model_name = os.getenv("OPENAI_MODEL_NAME", "gpt-4o-mini")
            llm = ChatOpenAI(model=model_name, api_key=openai_key, temperature=0.0)
            result = llm.invoke(prompt).content

        else:
            raise ValueError(f"지원하지 않는 모드입니다: {mode}")

    except Exception as e:
        elapsed = round(time.time() - start_time, 2)
        return f"⚠️ **[{mode}] 연결 에러:**\n`{str(e)}`\n\n*(API 키가 올바르지 않거나 네트워크 연결을 확인해주세요)*", elapsed

    elapsed = round(time.time() - start_time, 2)
    return result, elapsed