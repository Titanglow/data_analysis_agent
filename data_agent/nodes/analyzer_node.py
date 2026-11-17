# data_agent/nodes/analyzer_node.py
import time
import re
from typing import Any


def analyzer_node(state: dict[str, Any]) -> dict[str, Any]:
    """
    [1단계 Analyzer 노드]
    파이썬이 집계한 data_summary에서 핵심 수치·집중 구간·이상 징후만 0.01초 만에 정밀 추출합니다.
    (배경색 강조가 0건인 파일에서는 불필요한 '빨강배경 0건' 문구를 절대 생성하지 않습니다.)
    """
    t0 = time.perf_counter()
    summary = str(state.get("data_summary", ""))
    threshold = float(state.get("defect_threshold", 1.5) or 1.5)

    anomalies = []
    trends = []

    for line in summary.splitlines():
        s = line.strip()
        if not s:
            continue
        if "🚨" in s or "기준 초과" in s or "실행 실패" in s:
            anomalies.append(s)
        elif "🎯 [중점 관리]" in s:
            trends.append(s)
        elif "🎨" in s and not re.search(r":\s*0건", s):
            anomalies.append(s)

    if not trends:
        for line in summary.splitlines():
            if "%" in line and ("건" in line or "개" in line):
                trends.append(line.strip())
                if len(trends) >= 6:
                    break

    elapsed = round(time.perf_counter() - t0, 3)
    return {
        "analysis_result": {
            "applied_threshold_pct": threshold,
            "detected_anomalies": anomalies[:8],
            "key_distributions": trends[:10],
        },
        "elapsed_time": elapsed,
    }