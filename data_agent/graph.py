from typing import TypedDict, Any
from langgraph.graph import StateGraph, START, END
from data_agent.llm import call_llm

# 1. State 설계: 노드들이 참조하고 작업 결과를 전달할 공통 작업 저장소
class AnalysisState(TypedDict):
    question: str          # 사용자의 자연어 질문
    mode: str              # 실행 모드 (mock, local_ollama, cloud_gemini 등)
    data_summary: str      # 데이터프레임 기본 요약 정보 (스키마, 통계)
    report: str            # 생성된 최종 인사이트 리포트
    elapsed_time: float    # PoC 비교용 소요 시간(초)

# 2. Node 정의: 데이터 요약 및 리포트 생성 작업자
def analyze_and_report_node(state: AnalysisState) -> dict[str, Any]:
    prompt = f"""
    당신은 사내 데이터 분석 전문 에이전트입니다.
    아래 데이터 요약 정보를 바탕으로 사용자의 질문에 대한 분석 리포트를 마크다운 형식으로 작성하세요.

    [데이터 요약 및 통계]
    {state['data_summary']}

    [사용자 질문]
    {state['question']}
    """
    report_text, elapsed = call_llm(prompt, mode=state["mode"])
    return {
        "report": report_text,
        "elapsed_time": elapsed
    }

# 3. Graph 구성 및 컴파일
def build_agent_graph():
    workflow = StateGraph(AnalysisState)
    workflow.add_node("analyzer_reporter", analyze_and_report_node)
    workflow.add_edge(START, "analyzer_reporter")
    workflow.add_edge("analyzer_reporter", END)
    return workflow.compile()