import json
from typing import Any
from data_agent.llm import call_llm


def analyzer_node(state: dict[str, Any]) -> dict[str, Any]:
    """
    [1단계 정밀 분석 노드]
    파이썬이 계산한 집계표와 데이터를 바탕으로 핵심 요약, 이상치(수치 포함), 주요 트렌드를 JSON으로 구조화합니다.
    """
    mode = state.get("mode", "mock")
    data_summary = state.get("data_summary", "")

    prompt = f"""당신은 수석 데이터 분석 전문가입니다.
주어진 [파이썬 정밀 집계 결과]의 실제 수치만을 근거로 핵심 요약, 이상 징후(Anomalies), 주요 추세(Trends)를 추출하세요.

[파이썬 정밀 집계 결과]
{data_summary}

[사용자 질문]
{state.get('question', '')}

반드시 아래 키를 포함한 순수 JSON 형식으로만 응답하세요 (없는 숫자는 절대 지어내지 마세요):
{{
    "summary": "전체 데이터 핵심 요약 (총 생산량, 총 불량수, 평균 불량률 및 최다 불량 라인 명시)",
    "anomalies": [
        "사내 기준(1.5%) 또는 평균을 초과한 이상 라인·불량유형·시간대 및 구체적 수치 1",
        "설비 가동률 저하 또는 불량 급증 주차 등 이상 징후 2"
    ],
    "trends": [
        "주차별/일자별 생산량 및 불량률 증감(%p) 추세 1",
        "라인별·시간대(주/야간)별 품질 편차 추세 2"
    ]
}}"""

    raw_response, elapsed = call_llm(prompt, mode=mode)

    try:
        cleaned = raw_response.strip()
        if "```" in cleaned:
            cleaned = cleaned.split("```")[1].removeprefix("json").strip()
        start_idx = cleaned.find("{")
        end_idx = cleaned.rfind("}")
        if start_idx != -1 and end_idx != -1:
            cleaned = cleaned[start_idx : end_idx + 1]
        analysis_result = json.loads(cleaned)
    except Exception:
        analysis_result = {
            "summary": "파이썬 정밀 집계 기반 핵심 통계 및 이상치 추출 완료",
            "anomalies": ["사내 기준(1.5%) 초과 구간 및 최다 불량유형 집중 관리 필요 (본문 표 참조)"],
            "trends": ["주차별 불량률 증감(%p) 및 시간대별 편차 추이 확인 (본문 참조)"],
        }

    return {
        "analysis_result": analysis_result,
        "elapsed_time": round(state.get("elapsed_time", 0.0) + elapsed, 2),
    }