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


def _ensure_complete_report(report_text: str, data_summary: str) -> str:
    """
    [STEP 3 & 4: 끊김 확인 및 4번 섹션 줄바꿈·완결성 보정]
    모든 모델의 리포트가 1~4번까지 완결되었는지 검증하고,
    특히 '4. 실무 조치 제안'이 줄바꿈 없이 뭉쳐 있거나 잘린 경우 줄바꿈과 R&R·KPI를 완벽하게 보정합니다.
    """
    if not report_text or not report_text.strip():
        return report_text

    lines = [line.rstrip() for line in report_text.strip().splitlines()]
    if not lines:
        return report_text

    valid_endings = (".", "다", "요", "음", "함", "됨", "임", ")", "|", "%", "개", "건", "*", "]", "!", "?", "성")
    last_line = lines[-1].strip()
    is_last_line_cut = (
        not last_line.endswith(valid_endings)
        or last_line.endswith("**")
        or re.match(r"^(?:#+\s*)?4\.\s*실무", last_line) is not None
        or last_line in ("즉시 조치(24시간 내)", "근본 대책(1개월 내)", "- 즉시 조치", "- 근본 대책")
    )

    if is_last_line_cut:
        lines.pop()

    cleaned_text = "\n".join(lines).strip()

    has_sec4 = bool(re.search(r"(?:#+\s*)?4\.\s*실무\s*조치", cleaned_text))
    has_immediate = "즉시" in cleaned_text or "24시간" in cleaned_text
    has_longterm = "근본" in cleaned_text or "1개월" in cleaned_text or "중장기" in cleaned_text

    if has_sec4 and has_immediate and has_longterm and not is_last_line_cut:
        return cleaned_text

    target_line = "이상 발생 라인"
    for candidate in ["B라인", "C라인", "A라인", "3라인", "2라인", "1라인"]:
        if candidate in data_summary:
            target_line = candidate
            break

    target_defect = "주요 불량유형"
    for d_cand in ["도장얼룩", "표면스크래치", "치수불량", "이물질혼입"]:
        if d_cand in data_summary:
            target_defect = f"'{d_cand}'"
            break

    sec4_match = re.search(r"(?:#+\s*)?4\.\s*실무\s*조치.*", cleaned_text, flags=re.DOTALL)
    if sec4_match:
        cleaned_text = cleaned_text[: sec4_match.start()].strip()

    sec4_complete_block = (
        "\n\n### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
        "-   즉시 조치 (24시간 내):  \n"
        f"  -   [생산팀/설비팀]   `{target_line}`의 {target_defect} 다발 구간 설비 파라미터(노즐·온도·압력) 및 가동 상태를 즉시 점검합니다.\n"
        "  -   [품질팀]   주/야간 교대 시간대의 초·중·종물 샘플링 검사 빈도를 2배 강화하고 작업 표준서(SOP) 준수 여부를 확인합니다.\n"
        "-   근본 대책 (1개월 내):  \n"
        "  -   [설비팀/생산기술팀]   설비 가동률 저하 구간에 대한 예방 보전(PM) 주기를 단축하고, 사내 기준(1.5%) 초과 시 실시간 알림 임계치를 설정합니다.\n"
        f"  -   [품질보증팀]   {target_defect} 재발 방지를 위한 5M1E 표준 공정 조건을 재정립하고 교대 작업자 정기 교육을 실시합니다.\n"
        "-   기대 효과 (목표 KPI):  \n"
        f"  -   [정량 목표]   `{target_line}` 불량률 50% 이상 절감 및 전 라인 사내 기준(1.5%) 이내 안정화 달성"
    )

    return cleaned_text + sec4_complete_block


def reporter_node(state: dict[str, Any]) -> dict[str, Any]:
    """사내 표준 4단락(추세·이상항목·5M1E원인가설·R&R조치제안) 및 IF/ELSE 판정 헤더를 생성하는 정밀 보고서 노드"""
    mode = state.get("mode", "mock")
    data_summary = state.get("data_summary", "")
    question = state.get("question", "핵심 추세와 이상 징후, 원인 가설 및 조치 방안을 정리해 줘.")
    analysis_str = json.dumps(state.get("analysis_result", {}), ensure_ascii=False, indent=2)

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
아래 [파이썬 코드 집계 결과]의 실제 숫자만을 근거로, 팀장과 경영진이 주간회의에서 즉시 의사결정할 수 있는 최고 수준의 분석 보고서를 1번부터 4번까지 빠짐없이 작성하세요.

[사용자 질문]
{question}

[1단계 핵심 분석 요약 (JSON)]
{analysis_str}

[파이썬 코드 집계 결과 (실제 데이터)]
{data_summary}

[작성 형식 및 필수 포함 항목 (반드시 1번~4번 제목과 줄바꿈 구조를 정확히 지킬 것)]
### 1. 핵심 추세 (일자·주차·시간대별 흐름)
- 전체 생산량·불량수·평균 불량률 수치와 함께, 주차별/일자별 변화 방향(상승/하락) 및 이전 대비 증감 폭(%p)을 정확한 숫자로 서술하세요.
- 주간/야간 시간대별 또는 라인별 설비가동률과 불량 집중도 차이를 구체적인 수치로 비교하세요.

### 2. 주요 이상 항목 및 불량 유형 분류
- 사내 기준(1.5%) 또는 평균을 초과한 라인·주차·시간대(주/야간)·불량유형을 특정하고 아래 마크다운 표로 정리하세요.
| 순위 | 관리 대상 (라인/불량유형/시간대) | 핵심 수치 (생산량·불량수·불량률·가동률) | 기준(1.5%) 대비 판정 |
| :---: | :--- | :--- | :---: |

### 3. 원인 가설 (5M1E 관점 단계별 추론)
- 집계 데이터의 실제 특징(특정 라인 편중, 야간 시간대 증가, 설비가동률 저하, 특정 불량유형 집중 등)을 근거로 아래 4가지 관점에서 각각 줄을 바꾸어 추론하되,   반드시 '~할 가능성이 있습니다'   형태로 서술하세요. (자료로 알 수 없는 세부 원인은 '세부 원인은 자료로는 확인 불가하나'라고 명시)
-   설비(Machine):   (설비가동률 추이 및 공정 설비 관점 가설)
-   작업자/교대(Man):   (주간·야간 교대 시간대 및 숙련도·검사 관점 가설)
-   자재(Material):   (불량유형 특성에 따른 원부자재·도료·부품 관점 가설)
-   공정조건(Method):   (작업 표준서 SOP 및 공정 파라미터 관점 가설)

### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)
- 반드시 아래 하위 항목마다   줄바꿈(Enter)  을 하여 구분하고, 담당 부서(`[생산팀]`, `[설비팀]`, `[품질팀]`)와 정량적 기대 효과(KPI)를 명시하세요.
-   즉시 조치 (24시간 내):  
  -   [생산팀/설비팀]   (당장 현장에서 점검·조치할 구체적 행동 및 대상 라인/설비 명시)
  -   [품질팀]   (문제 시간대/불량유형에 대한 샘플링 검사 및 표준 준수 점검 명시)
-   근본 대책 (1개월 내):  
  -   [설비팀/생산기술팀]   (예방 보전 PM 주기 단축, 센서/알림 기준 설정 등 중장기 대책)
  -   [품질보증팀]   (재발 방지를 위한 5M1E 표준 조건 재정립 및 작업자 교육)
-   기대 효과 (목표 KPI):  
  -   [정량 목표]   (조치 실행 시 예상되는 불량률 감소 목표치 및 가동률 개선 목표 명시)

[절대 준수 규칙]
- 위 [파이썬 코드 집계 결과]에 없는 숫자는 절대 새로 지어내지 마세요 (환각 금지).
- 모든 하위 불릿(`-`)은 한 줄에 뭉쳐 쓰지 말고 반드시 줄바꿈(`\n`)하여 가독성을 높이세요."""

    report_text, elapsed = call_llm(prompt, mode=mode)

    complete_report_text = _ensure_complete_report(report_text, data_summary)
    final_report = status_banner + complete_report_text

    return {
        "report": final_report,
        "elapsed_time": round(state.get("elapsed_time", 0.0) + elapsed, 2),
    }