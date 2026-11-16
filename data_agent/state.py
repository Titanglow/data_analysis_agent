from typing import TypedDict, Any


class AnalysisState(TypedDict):
    """
    LangGraph 워크플로우에서 노드(START -> analyzer -> reporter -> END)들이
    순서대로 데이터를 주고받을 때 사용하는 공용 상태(State) 바구니입니다.
    """
    question: str                  # 사용자가 입력한 분석 질문 (또는 AI 비서 자동 질문)
    mode: str                      # 선택한 AI 모델 코드 (local_ollama, cloud_groq, cloud_gemini, mock)
    data_summary: str              # 파이썬이 집계한 시트 요약 및 상세 데이터 텍스트
    analysis_result: dict[str, Any]  # 1단계(analyzer_node)가 추출한 핵심 요약·이상치·트렌드
    report: str                    # 2단계(reporter_node)가 완성한 최종 마크다운 보고서
    elapsed_time: float            # 총 분석 소요 시간(초)