# data_agent/graph.py
from langgraph.graph import StateGraph, START, END
from data_agent.state import AgentState
from data_agent.nodes.analyzer_node import analyzer_node
from data_agent.nodes.reporter_node import reporter_node


def build_agent_graph():
    """START -> analyzer -> reporter -> END 순차 파이프라인 그래프를 생성합니다."""
    workflow = StateGraph(AgentState)
    workflow.add_node("analyzer", analyzer_node)
    workflow.add_node("reporter", reporter_node)

    workflow.add_edge(START, "analyzer")
    workflow.add_edge("analyzer", "reporter")
    workflow.add_edge("reporter", END)

    return workflow.compile()