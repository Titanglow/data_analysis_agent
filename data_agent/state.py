# data_agent/state.py
from typing import TypedDict, Any

class AgentState(TypedDict, total=False):
    question: str                   # 사용자 또는 AI 비서 생성 질문
    mode: str                       # 실행 모델 모드 (cloud_groq, cloud_gemini, local_ollama, mock)
    data_summary: str               # 파이썬이 선계산한 정밀 기초통계 및 데이터 요약
    defect_threshold: float         # 사내 품질 기준값 (기본 1.5%)
    analysis_result: dict[str, Any] # 1단계 Analyzer 노드 분석 결과
    report: str                     # 2단계 Reporter 노드 최종 마크다운 리포트
    elapsed_time: float             # 총 소요 시간(초)


