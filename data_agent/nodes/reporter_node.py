import json
from typing import Any
from data_agent.llm import call_llm


def reporter_node(state: dict[str, Any]) -> dict[str, Any]:
    """분석 결과를 바탕으로 경영진 보고용 요약 리포트를 작성하는 노드"""
    analysis_str = json.dumps(state.get("analysis_result", {}), ensure_ascii=False, indent=2)

    prompt = f"""당신은 품질 및 데이터 보고서 작성자입니다.
분석 결과를 바탕으로 경영진 보고용 요약 리포트를 마크다운 형식으로 작성하세요.

[분석 결과 (JSON)]
{analysis_str}

[요구사항]
1. 핵심 추세와 문제점을 명확히 기술할 것
2. 구체적인 실무 조치 방안을 제안할 것
3. 전문적이고 간결한 톤을 유지할 것"""

    report_text, elapsed = call_llm(prompt, mode=state.get("mode", "mock"))

    return {
        "report": report_text,
        "elapsed_time": round(state.get("elapsed_time", 0.0) + elapsed, 2)
    }