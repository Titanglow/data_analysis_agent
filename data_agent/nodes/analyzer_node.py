import json
from typing import Any
from data_agent.llm import call_llm


def analyzer_node(state: dict[str, Any]) -> dict[str, Any]:
    """데이터를 분석하여 주요 패턴과 이상 징후를 JSON 구조로 추출하는 노드"""
    prompt = f"""당신은 데이터 분석 전문가입니다. 주어진 데이터를 분석하여 주요 패턴과 이상 징후를 식별하세요.

[데이터 요약]
{state.get('data_summary', '')}

[사용자 질문]
{state.get('question', '')}

출력 형식은 반드시 아래 키를 포함한 JSON 형식이어야 합니다:
{{
    "summary": "데이터 핵심 요약",
    "anomalies": ["이상치 또는 주의 항목 1", "이상치 2"],
    "trends": ["주요 트렌드 1", "주요 트렌드 2"]
}}"""

    raw_response, elapsed = call_llm(prompt, mode=state.get("mode", "mock"))

    # JSON 파싱 시도 (테스트 모드이거나 일반 텍스트일 경우 안전하게 폴백 처리)
    try:
        cleaned = raw_response.strip().removeprefix("```json").removesuffix("```").strip()
        analysis_result = json.loads(cleaned)
    except Exception:
        analysis_result = {
            "summary": raw_response,
            "anomalies": ["자동 감지된 특이사항 없음 (원문 참조)"],
            "trends": ["상세 리포트 참조"]
        }

    return {
        "analysis_result": analysis_result,
        "elapsed_time": state.get("elapsed_time", 0.0) + elapsed
    }