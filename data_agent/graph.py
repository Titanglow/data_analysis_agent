from typing import TypedDict, Any
from langgraph.graph import StateGraph, START, END
from data_agent.nodes.analyzer_node import analyzer_node
from data_agent.nodes.reporter_node import reporter_node


# 1. State 정의: 분석가와 보고자 노드가 공유하는 작업 저장소
class AnalysisState(TypedDict):
    question: str  # 사용자의 자연어 질문
    mode: str  # 실행 모드 (mock, local_ollama 등)
    data_summary: str  # 데이터프레임 요약 정보
    analysis_result: dict[str, Any]  # Analyzer 노드가 추출한 JSON 결과
    report: str  # Reporter 노드가 작성한 최종 마크다운 리포트
    elapsed_time: float  # 총 소요 시간(초)


# 2. Graph 구성: START -> analyzer -> reporter -> END
def build_agent_graph():
    workflow = StateGraph(AnalysisState)

    # 작업자 노드 2개 등록
    workflow.add_node("analyzer", analyzer_node)
    workflow.add_node("reporter", reporter_node)

    # 엣지(화살표) 연결
    workflow.add_edge(START, "analyzer")
    workflow.add_edge("analyzer", "reporter")
    workflow.add_edge("reporter", END)

    return workflow.compile()