import os
import time
from dotenv import load_dotenv

load_dotenv()

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
            from langchain_google_genai import ChatGoogleGenerativeAI
            gemini_models = ["gemini-2.5-flash", "gemini-2.0-flash", "gemini-1.5-flash", "gemini-3-flash-preview"]
            last_err = None
            result = None
            for m_name in gemini_models:
                try:
                    llm = ChatGoogleGenerativeAI(model=m_name, temperature=0.0)
                    resp = llm.invoke(prompt)
                    result = str(resp.content) if not isinstance(resp.content, list) else "\n".join(str(x) for x in resp.content)
                    break
                except Exception as e:
                    last_err = e
                    continue
            if result is None:
                raise last_err

        elif mode == "cloud_openai":
            from langchain_openai import ChatOpenAI
            model_name = os.getenv("OPENAI_MODEL_NAME", "gpt-4o-mini")
            llm = ChatOpenAI(model=model_name, temperature=0.0)
            result = llm.invoke(prompt).content

        else:
            raise ValueError(f"지원하지 않는 모드입니다: {mode}")

    except Exception as e:
        elapsed = round(time.time() - start_time, 2)
        return f"⚠️ **[{mode}] 연결 에러:**\n`{str(e)}`\n\n*(API 키 또는 네트워크 상태를 확인해주세요)*", elapsed

    elapsed = round(time.time() - start_time, 2)
    return result, elapsed