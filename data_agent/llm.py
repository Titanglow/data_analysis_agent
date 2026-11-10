import os
import time
from dotenv import load_dotenv

load_dotenv()

class MockLLM:
    def invoke(self, prompt: str) -> str:
        time.sleep(0.5)
        return (
            "### 📊 데이터 분석 요약 리포트 (테스트 모드)\n"
            "- **추세:** 주요 지표가 주차별로 변동하고 있습니다.\n"
            "- **이상 항목:** 특정 라인에서 평균 대비 높은 불량률이 감지되었습니다.\n"
            "- **조치 제안:** 설비 점검 및 원인 분석이 필요합니다."
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
            # 실습 자료 기준 최신 모델 우선 시도 후 호환 모델 자동 폴백
            gemini_models = [
                os.getenv("GEMINI_MODEL_NAME", "gemini-3-flash-preview"),
                "gemini-2.0-flash",
                "gemini-1.5-flash"
            ]
            last_err = None
            result = None
            for m_name in gemini_models:
                try:
                    llm = ChatGoogleGenerativeAI(model=m_name, temperature=0.0)
                    resp = llm.invoke(prompt)
                    # 응답이 리스트 형태일 경우 텍스트만 추출
                    if isinstance(resp.content, list):
                        result = "\n".join([ item.get("text", str(item)) if isinstance(item, dict) else str(item) for item in resp.content ])
                    else:
                        result = str(resp.content)
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
        # 빨간색 전체 화면 크래시 대신 깔끔한 안내 메시지 반환
        elapsed = round(time.time() - start_time, 2)
        return f"⚠️ **[{mode}] 호출 중 오류 발생:**\n`{str(e)}`", elapsed

    elapsed = round(time.time() - start_time, 2)
    return result, elapsed