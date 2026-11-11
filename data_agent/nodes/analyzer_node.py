import json
from typing import Any
from data_agent.llm import call_llm


def analyzer_node(state: dict[str, Any]) -> dict[str, Any]:
    """데이터를 분석하여 주요 패턴과 이상 징후를 JSON 구조로 추출하는 노드"""
    mode = state.get("mode", "mock")
    data_summary = state.get("data_summary", "")

    # ⚡ 빠른 모드(local_ollama_fast)는 1단계 분석을 파이썬 규칙 엔진으로 즉시 처리(0.01초)하여 LLM 중복 호출 방지
    if mode == "local_ollama_fast":
        lines = [line.strip() for line in data_summary.splitlines() if line.strip()]
        return {
            "analysis_result": {
                "summary": "파이썬 고속 집계 엔진을 통한 핵심 통계 추출 완료",
                "anomalies": [l for l in lines if any(k in l for k in ["불량", "이상", "최대", "주의", "B라인"])][:3]
                or ["데이터 내 평균 대비 편차 구간 집중 점검 필요"],
                "trends": lines[:3] or ["주차별 생산량 및 품질 지표 추이 모니터링"],
            },
            "elapsed_time": state.get("elapsed_time", 0.0),
        }

    prompt = f"""당신은 데이터 분석 전문가입니다. 주어진 데이터를 분석하여 주요 패턴과 이상 징후를 식별하세요.

[데이터 요약]
{data_summary}

[사용자 질문]
{state.get('question', '')}

출력 형식은 반드시 아래 키를 포함한 JSON 형식이어야 합니다:
{{
    "summary": "데이터 핵심 요약 (한 문장)",
    "anomalies": ["이상치 또는 주의 항목 1", "이상치 2"],
    "trends": ["주요 트렌드 1", "주요 트렌드 2"]
}}"""

    raw_response, elapsed = call_llm(prompt, mode=mode)

    try:
        cleaned = raw_response.strip()
        if "```" in cleaned:
            cleaned = cleaned.split("```")[1].removeprefix("json").strip()
        analysis_result = json.loads(cleaned)
    except Exception:
        analysis_result = {
            "summary": raw_response[:200],
            "anomalies": ["리포트 본문 내 이상 징후 분석 참조"],
            "trends": ["리포트 본문 내 주요 추세 분석 참조"],
        }

    return {
        "analysis_result": analysis_result,
        "elapsed_time": round(state.get("elapsed_time", 0.0) + elapsed, 2),
    }