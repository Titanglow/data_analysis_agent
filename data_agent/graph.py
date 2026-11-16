from langgraph.graph import StateGraph, START, END
from data_agent.state import AnalysisState
from data_agent.nodes.analyzer_node import analyzer_node
from data_agent.nodes.reporter_node import reporter_node


def build_agent_graph():
    """
    사내 데이터 분석 에이전트 워크플로우 생성:
    START -> analyzer (파이썬 고속 집계 노드) -> reporter (AI 보고서 작성 노드) -> END
    """
    workflow = StateGraph(AnalysisState)

    # 1. 노드 등록
    workflow.add_node("analyzer", analyzer_node)
    workflow.add_node("reporter", reporter_node)

    # 2. 엣지(실행 순서 화살표) 연결
    workflow.add_edge(START, "analyzer")
    workflow.add_edge("analyzer", "reporter")
    workflow.add_edge("reporter", END)

    return workflow.compile()