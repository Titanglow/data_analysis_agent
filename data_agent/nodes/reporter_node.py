import json
import re
from typing import Any
from data_agent.llm import call_llm


def _extract_worst_rate(data_summary: str) -> float:
    """파이썬 집계 텍스트에서 최고 불량률(%) 숫자를 추출합니다."""
    matches = re.findall(r"(\d+\.\d+)%", data_summary)
    if not matches:
        return 0.0
    return max(float(m) for m in matches)


def reporter_node(state: dict[str, Any]) -> dict[str, Any]:
    """사내 표준 4단락(추세·이상항목·5M1E원인가설·R&R조치제안) 및 IF/ELSE 판정 헤더를 생성하는 노드"""
    data_summary = state.get("data_summary", "")
    question = state.get("question", "핵심 추세와 이상 징후, 원인 가설 및 조치 방안을 정리해 줘.")
    analysis_str = json.dumps(state.get("analysis_result", {}), ensure_ascii=False, indent=2)

    # 사내 기준(1.5%) 초과 여부 자동 판정 (IF/ELSE 워크플로우 규칙)
    worst_rate = _extract_worst_rate(data_summary)
    threshold = 1.5
    if worst_rate > threshold:
        status_banner = (
            f"> 🚨 **[확인 필요] 불량률이 사내 기준({threshold}%)을 초과한 항목(최고 `{worst_rate:.2f}%`)이 있습니다. "
            "주간회의 안건으로 올리세요.**\n\n"
        )
    else:
        status_banner = (
            f"> ✅ **[이상 없음] 모든 구간의 불량률이 사내 기준({threshold}%) 이내(최고 `{worst_rate:.2f}%`)입니다. "
            "정기 보고로만 처리하세요.**\n\n"
        )

    prompt = f"""당신은 15년 경력의 수석 품질·데이터 분석 전문가입니다.
아래 [파이썬 코드 집계 결과]의 실제 숫자만을 근거로, 팀장과 경영진이 주간회의에서 즉시 의사결정할 수 있는 보고서를 작성하세요.

[사용자 질문]
{question}

[1차 핵심 요약 (JSON)]
{analysis_str}

[파이썬 코드 집계 결과 (실제 데이터)]
{data_summary}

[작성 형식 및 필수 포함 항목 (반드시 아래 순서를 지킬 것)]
### 1. 핵심 추세 (일자·주차·시간대별 흐름)
- 주차별/일자별 생산량 및 불량률의 변화 방향(상승/하락)과 **전주 대비 증감 폭(%p)**을 정확한 숫자로 서술하세요.

### 2. 주요 이상 항목 및 불량 유형 분류
- 사내 기준(1.5%) 또는 평균을 초과한 **라인·주차·시간대(주/야간)·불량유형**을 특정하고 아래 마크다운 표로 정리하세요.
| 순위 | 관리 대상 (라인/불량유형/시간대) | 핵심 수치 (생산량·불량수·불량률) | 기준 대비 상태 |
| :---: | :--- | :--- | :---: |

### 3. 원인 가설 (5M1E 관점 단계별 추론)
- 설비(Machine), 작업자/교대(Man), 자재(Material), 공정조건(Method) 관점에서 원인을 추론하되, 단정하지 말고 **반드시 '~할 가능성이 있습니다'** 형태로 2~3개 제시하세요. (자료로 알 수 없는 세부 원인은 '세부 원인은 자료로는 확인 불가하나'라고 명시할 것)

### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)
- 담당 부서(생산팀·품질팀·설비팀)가 바로 실행할 수 있도록 **즉시 조치(24시간 내)**와 **근본 대책(1개월 내)**으로 나누어 구체적인 행동 지침을 제시하세요.

[절대 준수 규칙]
- 위 [파이썬 코드 집계 결과]에 없는 숫자는 절대 새로 지어내지 마세요 (환각 금지).
- 추상적이고 뻔한 교과서적 표현("전반적으로 양호함", "지속 모니터링 필요") 대신 데이터 안의 실제 라인명, 불량유형명, 시간대, 수치를 직접 인용하세요."""

    report_text, elapsed = call_llm(prompt, mode=state.get("mode", "mock"))
    final_report = status_banner + report_text

    return {
        "report": final_report,
        "elapsed_time": round(state.get("elapsed_time", 0.0) + elapsed, 2),
    }