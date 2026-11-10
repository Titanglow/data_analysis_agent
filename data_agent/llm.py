import os
import time
from dotenv import load_dotenv

load_dotenv()


class MockLLM:
    """Ollama 미설치 및 API 키 미입력 상태에서 파이프라인을 검증하기 위한 테스트용 모델"""

    def invoke(self, prompt: str) -> str:
        time.sleep(0.8)  # 실제 추론 시간 시뮬레이션
        return (
            "### 📊 데이터 분석 요약 리포트 (테스트 모드)\n"
            "- **핵심 추세:** 업로드된 데이터의 주요 수치 컬럼을 기준으로 집계한 결과, 상위 항목에 실적이 집중되어 있습니다.\n"
            "- **이상 징후:** 평균 대비 편차가 큰 구간이 관측되어 세부 데이터 확인이 필요합니다.\n"
            "- **조치 제안:** 하위 실적 카테고리의 원인을 파악하고 주간 단위로 지표를 모니터링하세요."
        )


def call_llm(prompt: str, mode: str = "mock") -> tuple[str, float]:
    """
    선택된 모드(mock / local_ollama / cloud_gemini / cloud_openai)로 질문을 보내고
    (응답 텍스트, 소요 시간)을 반환합니다.
    """
    start_time = time.time()

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
        llm = ChatGoogleGenerativeAI(model="gemini-2.5-flash", temperature=0.0)
        result = llm.invoke(prompt).content

    elif mode == "cloud_openai":
        from langchain_openai import ChatOpenAI
        llm = ChatOpenAI(model="gpt-4o-mini", temperature=0.0)
        result = llm.invoke(prompt).content

    else:
        raise ValueError(f"지원하지 않는 모드입니다: {mode}")

    elapsed = round(time.time() - start_time, 2)
    return result, elapsed