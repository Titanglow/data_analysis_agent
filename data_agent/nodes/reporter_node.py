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
    [STEP 3 & 4: 끊김 확인 및 형식 완결성 검사]
    토큰 한도나 로컬 모델 특성으로 인해 4번 항목('4. 실무 조치 제안')이 중간에 잘렸거나 누락된 경우,
    잘린 미완성 문장을 정리하고 데이터 기반의 시점별 조치 로드맵(즉시 조치 24h / 근본 대책 1개월)을 자동 보완합니다.
    """
    if not report_text or not report_text.strip():
        return report_text

    lines = [line.rstrip() for line in report_text.strip().splitlines()]
    if not lines:
        return report_text

    # 마지막 줄이 완성된 문장/표가 아니라 단어 중간('4. 실무 조', '- **생산' 등)에서 끊겼는지 점검
    valid_endings = (".", "다", "요", "음", "함", "됨", "임", ")", "|", "%", "개", "건", "*", "]", "!", "?")
    last_line = lines[-1].strip()
    is_last_line_cut = (
        not last_line.endswith(valid_endings)
        or last_line.endswith("**")
        or re.match(r"^(?:#+\s*)?4\.\s*실무", last_line) is not None
        or last_line in ("즉시 조치(24시간 내)", "근본 대책(1개월 내)", "- 즉시 조치", "- 근본 대책")
    )

    if is_last_line_cut:
        lines.pop()  # 중간에 잘린 불완전한 마지막 줄 제거

    cleaned_text = "\n".join(lines).strip()

    # 4번 섹션('4. 실무 조치 제안')과 핵심 하위 항목('즉시 조치', '근본 대책')이 모두 정상 포함되어 있는지 확인
    has_sec4 = bool(re.search(r"(?:#+\s*)?4\.\s*실무\s*조치", cleaned_text))
    has_immediate = "즉시" in cleaned_text or "24시간" in cleaned_text
    has_longterm = "근본" in cleaned_text or "1개월" in cleaned_text or "중장기" in cleaned_text

    if has_sec4 and has_immediate and has_longterm and not is_last_line_cut:
        return cleaned_text

    # 데이터 요약에서 주요 관리 대상 라인 및 불량유형 키워드 추출
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

    # 만약 4번 제목만 찍히고 본문이 잘렸다면 4번 제목 이후를 깔끔하게 재작성
    sec4_match = re.search(r"(?:#+\s*)?4\.\s*실무\s*조치.*", cleaned_text, flags=re.DOTALL)
    if sec4_match:
        cleaned_text = cleaned_text[: sec4_match.start()].strip()

    sec4_complete_block = (
        "\n\n### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
        f"- **즉시 조치 (24시간 내):**\n"
        f"  - **[생산팀/설비팀]** `{target_line}`의 {target_defect} 다발 구간 설비 파라미터(노즐·온도·압력) 및 가동 상태를 즉시 점검합니다.\n"
        f"  - **[품질팀]** 주/야간 교대 시간대의 초·중·종물 샘플링 검사 빈도를 강화하고 작업 표준서(SOP) 준수 여부를 확인합니다.\n"
        f"- **근본 대책 (1개월 내):**\n"
        f"  - **[설비팀/생산기술팀]** 설비 가동률 저하 구간에 대한 예방 보전(PM) 주기를 단축하고, 사내 기준(1.5%) 초과 시 실시간 알림 임계치를 설정합니다.\n"
        f"  - **[품질보증팀]** {target_defect} 재발 방지를 위한 5M1E 표준 공정 조건을 재정립하고 교대 작업자 정기 교육을 실시합니다."
    )

    return cleaned_text + sec4_complete_block


def reporter_node(state: dict[str, Any]) -> dict[str, Any]:
    """사내 표준 4단락(추세·이상항목·5M1E원인가설·R&R조치제안) 및 IF/ELSE 판정 헤더를 생성하는 노드"""
    mode = state.get("mode", "mock")
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

    # 빠른 모드일 때는 각 문장을 핵심 개조식(1~2문장)으로 간결하게 쓰도록 지시해 속도 향상 및 끊김 방지
    concise_rule = (
        "- **[빠른 모드 필수 규칙]** 불필요한 서론이나 장황한 일반론을 생략하고, 각 하위 불릿은 **반드시 핵심 숫자가 포함된 1문장 개조식**으로 간결하게 작성하여 1번부터 4번 끝까지 빠르게 완주하세요."
        if mode == "local_ollama_fast"
        else "- 각 항목은 집계된 실제 수치를 인용하여 구체적이고 명확하게 작성하세요."
    )

    prompt = f"""당신은 15년 경력의 수석 품질·데이터 분석 전문가입니다.
아래 [파이썬 코드 집계 결과]의 실제 숫자만을 근거로, 팀장과 경영진이 주간회의에서 즉시 의사결정할 수 있는 보고서를 1번부터 4번까지 빠짐없이 작성하세요.

[사용자 질문]
{question}

[1차 핵심 요약 (JSON)]
{analysis_str}

[파이썬 코드 집계 결과 (실제 데이터)]
{data_summary}

[작성 형식 및 필수 포함 항목 (반드시 1번~4번까지 모두 완성할 것)]
### 1. 핵심 추세 (일자·주차·시간대별 흐름)
- 주차별/일자별 생산량 및 불량률의 변화 방향(상승/하락)과 **이전 대비 증감 폭(%p)**을 정확한 숫자로 2문장 이내로 서술하세요.

### 2. 주요 이상 항목 및 불량 유형 분류
- 사내 기준(1.5%) 또는 평균을 초과한 **라인·주차·시간대(주/야간)·불량유형**을 특정하고 아래 마크다운 표로 정리하세요.
| 순위 | 관리 대상 (라인/불량유형/시간대) | 핵심 수치 (생산량·불량수·불량률) | 기준 대비 상태 |
| :---: | :--- | :--- | :---: |

### 3. 원인 가설 (5M1E 관점 단계별 추론)
- 집계 데이터의 실제 특징(특정 라인 편중, 야간 시간대 증가, 설비가동률 저하, 특정 불량유형 집중 등)을 근거로 설비(Machine), 작업자/교대(Man), 자재(Material), 공정조건(Method) 관점에서 원인을 추론하되, **반드시 '~할 가능성이 있습니다'** 형태로 각 1문장씩 제시하세요. (자료로 알 수 없는 세부 원인은 '세부 원인은 자료로는 확인 불가하나'라고 명시)

### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)
- **즉시 조치 (24시간 내):** [생산팀/설비팀/품질팀]이 당장 현장에서 확인할 구체적 행동 2가지
- **근본 대책 (1개월 내):** [설비팀/품질보증팀]이 재발 방지를 위해 수립할 대책 2가지

[절대 준수 규칙]
- 위 [파이썬 코드 집계 결과]에 없는 숫자는 절대 새로 지어내지 마세요 (환각 금지).
- {concise_rule}"""

    report_text, elapsed = call_llm(prompt, mode=mode)

    # STEP 3 & 4: 출력 끊김 확인 및 4번 섹션 자동 완결 보장
    complete_report_text = _ensure_complete_report(report_text, data_summary)
    final_report = status_banner + complete_report_text

    return {
        "report": final_report,
        "elapsed_time": round(state.get("elapsed_time", 0.0) + elapsed, 2),
    }