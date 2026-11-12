import json
import re
from typing import Any
from data_agent.llm import call_llm


def _detect_data_domain(data_summary: str) -> str:
    """업로드된 데이터가 제조/생산 데이터인지, SW/시스템 테스트케이스(TC)인지, 일반 문서인지 판별합니다."""
    s_lower = data_summary.lower()
    if any(k in s_lower for k in ["pass/fail", "precondition", "procedure", "actual result", "domain", "tc-"]):
        return "tc_verification"
    if "생산량" in s_lower and "불량수" in s_lower:
        return "manufacturing"
    return "general"


def _clean_empty_bullets(text: str) -> str:
    """LLM 출력 결과에서 글자 없이 '•', '-', '*'만 찍힌 빈 줄을 완벽히 제거합니다."""
    cleaned_lines = []
    for line in text.splitlines():
        stripped = line.strip()
        if re.match(r"^[-*•▪▶\s]+$", stripped) and stripped != "":
            continue
        cleaned_lines.append(line.rstrip())
    return "\n".join(cleaned_lines).strip()


def reporter_node(state: dict[str, Any]) -> dict[str, Any]:
    """데이터 도메인(제조 공정 vs TC 검증 시트 vs 일반 데이터)을 자동 인식해 맞춤형 4단락 보고서를 작성하는 노드"""
    mode = state.get("mode", "mock")
    data_summary = state.get("data_summary", "")
    question = state.get("question", "핵심 현황과 이상·주의 항목, 원인 가설 및 실무 조치 방안을 정리해 줘.")
    analysis_str = json.dumps(state.get("analysis_result", {}), ensure_ascii=False, indent=2)

    domain_type = _detect_data_domain(data_summary)

    # 도메인별 IF/ELSE 자동 판정 배너 생성
    if domain_type == "manufacturing":
        matches = re.findall(r"평균 불량률\(%\)\s*\|\s*(\d+\.\d+)%|불량률\(%\)\s*(\d+\.\d+)%", data_summary)
        rates = [float(m[0] or m[1]) for m in matches if (m[0] or m[1])]
        if not rates:
            all_pcts = [float(x) for x in re.findall(r"(\d+\.\d+)%", data_summary) if float(x) < 50.0]
            worst_rate = max(all_pcts) if all_pcts else 0.0
        else:
            worst_rate = max(rates)

        if worst_rate > 1.5:
            status_banner = f"> 🚨 **[확인 필요] 불량률이 사내 기준(1.5%)을 초과한 항목(최고 `{worst_rate:.2f}%`)이 있습니다. 주간회의 안건으로 올리세요.**\n\n"
        else:
            status_banner = f"> ✅ **[이상 없음] 모든 구간의 불량률이 사내 기준(1.5%) 이내입니다. 정기 보고로만 처리하세요.**\n\n"
    else:
        # TC 검증 시트 또는 일반 엑셀일 때: 빨강 배경색 / 불합격 / FAIL / 미기재 항목 감지
        has_red_or_fail = any(k in data_summary for k in ["🔴빨강", "불합격", "FAIL", "NG", "오류", "주의"])
        if has_red_or_fail:
            status_banner = "> 🚨 **[확인 필요] 데이터 내 빨간색 강조 표시(불합격/주의) 또는 미기재 점검 항목이 감지되었습니다. 담당 부서 확인이 필요합니다.**\n\n"
        else:
            status_banner = "> ✅ **[이상 없음] 특이 불합격 색상 표시 없이 정상 범위로 집계되었습니다.**\n\n"

    if domain_type == "tc_verification":
        role_guide = """당신은 수석 QA 및 시스템 검증(Test Case) 분석 전문가입니다.
업로드된 데이터는 시트별 테스트케이스(Domain, App, ID, Title, Procedure, Pass/Fail Criteria, Actual Result 등) 및 셀 배경색(빨강=불합격/주의 등)이 포함된 검증 데이터입니다.
엉뚱한 공장 설비나 작업자 교대(Man) 이야기를 쓰지 말고, 반드시 현재 데이터의 **시트명, App/Domain 분류, TC ID, Actual Result(빨강 배경색 표시 및 미기재 현황)**만을 근거로 작성하세요."""
    elif domain_type == "manufacturing":
        role_guide = """당신은 15년 경력의 수석 생산·품질 데이터 분석 전문가입니다.
아래 [파이썬 코드 집계 결과]의 실제 라인, 주차, 시간대, 불량유형, 불량률(%) 숫자만을 근거로 작성하세요."""
    else:
        role_guide = """당신은 수석 기업 데이터 분석 전문가입니다.
아래 [파이썬 코드 집계 결과]에 나타난 실제 시트명, 컬럼명, 항목별 건수, 작성률(%), 색상 강조 표시 항목만을 근거로 작성하세요."""

    prompt = f"""{role_guide}

[사용자 질문]
{question}

[1단계 핵심 분석 요약 (JSON)]
{analysis_str}

[파이썬 코드 집계 결과 (실제 데이터)]
{data_summary}

[작성 형식 및 필수 포함 항목 (반드시 아래 1번~4번 제목을 그대로 사용하고 빈 줄/빈 불릿을 만들지 말 것)]
### 1. 핵심 추세 및 데이터 현황 요약
- 전체 조회 건수, 시트별/분류별(Domain·App·라인 등) 분포 비중(%)과 작성 완료율 또는 핵심 지표를 정확한 숫자로 서술하세요.

### 2. 주요 이상·주의 항목 분류 (색상 강조 및 미기재/기준 초과 항목)
- 데이터에서 감지된 **빨간색 배경 표시(`🔴빨강`), 불합격(FAIL), 미기재(`-`) 항목 또는 기준 초과 대상**을 특정하고 아래 마크다운 표로 정리하세요.
| 순위 | 관리 대상 (시트/분류/ID/라인) | 핵심 내용 및 감지 상태 (건수·비율·배경색 표시) | 판정 상태 |
| :---: | :--- | :--- | :---: |

### 3. 원인 가설 (데이터 근거 기반 추론)
- 발견된 주의 항목(예: 특정 TC/기능의 빨강 배경 표시, Actual Result 미기재, 또는 특정 불량 집중)이 발생한 원인을 데이터의 실제 컬럼 내용과 절차(Procedure/조건)를 근거로 3가지 이내로 추론하되, **반드시 '~할 가능성이 있습니다'** 형태로 작성하세요.

### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)
- **즉시 조치 (24시간 내):**
  - **[담당 검증/현업팀]** (빨간색 배경 표시된 주의 항목 또는 기준 초과 건에 대해 당장 재검증·확인할 구체적 행동)
  - **[품질/QA팀]** (미기재 항목 보완 및 판정 기준 점검 행동)
- **근본 대책 (1개월 내):**
  - **[개발/기술팀]** (반복 이슈 방지를 위한 표준 양식·자동화·프로세스 개선 대책)
- **기대 효과 (목표 KPI):**
  - **[정량 목표]** (조치 완료 시 달성할 검증 완료율 100% 또는 불량 감소 목표치 명시)

[절대 준수 규칙]
- 글자 없이 불릿 기호(`•`, `-`, `*`)만 있는 빈 줄은 절대 출력하지 마세요.
- 위 [파이썬 코드 집계 결과]에 없는 숫자나 데이터와 무관한 일반론은 절대 지어내지 마세요."""

    report_text, elapsed = call_llm(prompt, mode=mode)
    clean_report = _clean_empty_bullets(report_text)
    final_report = status_banner + clean_report

    return {
        "report": final_report,
        "elapsed_time": round(state.get("elapsed_time", 0.0) + elapsed, 2),
    }