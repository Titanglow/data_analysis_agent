import re
from typing import Any


def analyzer_node(state: dict[str, Any]) -> dict[str, Any]:
    """
    [코드 집계 노드 (03차시 표준 원칙)]
    LLM을 중복 호출해 시간을 낭비하거나 숫자를 지어내지 않도록,
    파이썬 코드가 직접 집계표에서 핵심 요약, 색상 문맥, 이상 징후를 0.01초 만에 정확히 추출합니다.
    """
    data_summary = state.get("data_summary", "")
    lines = [l.strip() for l in data_summary.splitlines() if l.strip()]

    anomalies = []
    trends = []

    for line in lines:
        if any(k in line for k in ["🎨빨강강조", "기준(1.5%) 초과", "실행 실패", "미실행", "Not Run", "중요도/심각도"]):
            if len(anomalies) < 5:
                anomalies.append(line[:140])
        elif any(k in line for k in ["비중(%)", "평균 불량률", "출처 시트명", "총 생산량", "전체 요구사항", "커버리지"]):
            if len(trends) < 5:
                trends.append(line[:140])

    if not anomalies:
        anomalies = ["집계표 내 강조 표시 항목 및 주요 관리 대상 확인 필요"]
    if not trends:
        trends = ["시트별·분류별 주요 분포 및 지표 현황 본문 참조"]

    return {
        "analysis_result": {
            "summary": "파이썬 정밀 코드 집계 및 셀 색상 문맥 분석 완료",
            "anomalies": anomalies,
            "trends": trends,
        },
        "elapsed_time": state.get("elapsed_time", 0.0),
    }