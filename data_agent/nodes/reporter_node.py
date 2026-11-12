import json
import re
from typing import Any
from data_agent.llm import call_llm


def _detect_data_domain(data_summary: str) -> str:
    s_lower = data_summary.lower()
    if any(k in s_lower for k in ["pass/fail", "precondition", "procedure", "actual result", "not run", "tc-"]):
        return "tc_verification"
    if "생산량" in s_lower and "불량수" in s_lower:
        return "manufacturing"
    return "general"


def _convert_json_if_needed(raw_text: str) -> str:
    """만약 LLM이 마크다운 대신 JSON({\"summary\": ...})을 출력했을 경우 자동으로 4단락 마크다운 보고서로 변환합니다."""
    stripped = raw_text.strip()
    model_header = ""
    if stripped.startswith("*(연결 모델:") or stripped.startswith("*(로컬 모델:"):
        first_nl = stripped.find("\n")
        if first_nl != -1:
            model_header = stripped[:first_nl].strip() + "\n\n"
            stripped = stripped[first_nl:].strip()

    if stripped.startswith("{") and stripped.endswith("}"):
        try:
            data = json.loads(stripped)
            summary = data.get("summary", "데이터 핵심 요약 분석 완료")
            anomalies = data.get("anomalies", [])
            trends = data.get("trends", [])
            causes = data.get("causes", data.get("hypotheses", []))
            actions = data.get("actions", data.get("recommendations", []))

            md_parts = [
                model_header + "### 1. 시트별 종합 현황 및 핵심 추세",
                f"- {summary}",
            ]
            for t in trends:
                md_parts.append(f"- {t}")

            md_parts.append("\n### 2. 주요 이상·강조 항목 및 문맥 분석")
            for a in anomalies:
                md_parts.append(f"- {a}")

            md_parts.append("\n### 3. 원인 가설 및 배경 분석 (데이터 근거 기반)")
            if causes:
                for c in causes:
                    md_parts.append(f"- {c}")
            else:
                md_parts.append("- **중점 관리 항목 설정 가능성:** 색상 강조 표시된 항목들은 중요도(High) 또는 경험 기반(Error Guessing) 등 중점 검증 시나리오로 지정되었을 가능성이 있습니다.")
                md_parts.append("- **실행 단계 대기 가능성:** 미실행(Not Run) 비중이 높은 경우 실제 기능 결함이 아니라 테스트 설계 완료 후 본 실행(Execution) 전 단계일 가능성이 있습니다.")

            md_parts.append("\n### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)")
            if actions:
                for act in actions:
                    md_parts.append(f"- {act}")
            else:
                md_parts.append("- **즉시 조치 (24시간 내):**\n  - **[담당 검증/QA팀]** 색상 강조(중요도 High·특수 기법 등)가 표시된 핵심 항목부터 우선 테스트 수행 일정을 확정합니다.")
                md_parts.append("- **근본 대책 (1개월 내):**\n  - **[개발/품질팀]** 테스트 실행 후 실제 Pass/Fail 결과를 Summary 시트와 자동 연동하여 미실행(Not Run) 잔여 건수를 0건으로 관리합니다.")

            return "\n".join(md_parts)
        except Exception:
            pass
    return raw_text


def _fix_collapsed_markdown_tables(text: str) -> str:
    """Gemini 등이 마크다운 표를 줄바꿈 없이 한 줄('| ... | | :---: | || 1 |')로 출력했을 때 줄바꿈을 자동 복원합니다."""
    # '|| ' 또는 '| |' 패턴으로 이어붙은 표 행 분리
    fixed = re.sub(r"\|\s*\|\s*(?=(?::?-+:?|\d+\s*\||TC-|REQ-|[A-Za-z가-힣]))", "|\n| ", text)
    # 문장 끝에 바로 '| 순위 |' 헤더가 붙은 경우 줄바꿈 삽입
    fixed = re.sub(r"([.다요음함)])\s*(\|\s*순위\s*\|)", r"\1\n\n\2", fixed)
    return fixed


def _clean_empty_bullets(text: str) -> str:
    cleaned_lines = []
    for line in text.splitlines():
        stripped = line.strip()
        if re.match(r"^[-*•▪▶\s]+$", stripped) and stripped != "":
            continue
        cleaned_lines.append(line.rstrip())
    return "\n".join(cleaned_lines).strip()


def reporter_node(state: dict[str, Any]) -> dict[str, Any]:
    """데이터 도메인과 시트 구성, 셀 색상 문맥을 정확히 반영하여 4단락 마크다운 보고서를 생성하는 노드"""
    mode = state.get("mode", "mock")
    data_summary = state.get("data_summary", "")
    question = state.get("question", "시트별 종합 현황과 강조 표시 항목의 실제 의미, 원인 분석 및 실무 조치 방안을 정리해 줘.")

    domain_type = _detect_data_domain(data_summary)

    # 도메인 및 실제 상태 기반 IF/ELSE 판정 배너
    if domain_type == "manufacturing":
        matches = re.findall(r"평균 불량률\(%\)\s*\|\s*(\d+\.\d+)%|불량률\(%\)\s*(\d+\.\d+)%", data_summary)
        rates = [float(m[0] or m[1]) for m in matches if (m[0] or m[1])]
        worst_rate = max(rates) if rates else 0.0
        if worst_rate > 1.5:
            status_banner = f"> 🚨 **[확인 필요] 불량률이 사내 기준(1.5%)을 초과한 라인(최고 `{worst_rate:.2f}%`)이 있습니다. 주간회의 안건으로 올리세요.**\n\n"
        else:
            status_banner = "> ✅ **[이상 없음] 모든 라인의 불량률이 사내 기준(1.5%) 이내입니다. 정기 보고로만 처리하세요.**\n\n"
    else:
        # 실제 FAIL이 있는지, 아니면 Not Run(미실행) 상태에서 중요도/기법 색상 강조만 있는지 정확히 구분!
        has_real_fail = bool(re.search(r"실행 실패|불합격\(FAIL\)|\bFAIL\b", data_summary))
        has_not_run = "Not Run" in data_summary or "미실행" in data_summary
        if has_real_fail:
            status_banner = "> 🚨 **[확인 필요] 실행 결과 불합격(FAIL) 또는 결함 항목이 감지되었습니다. 원인 분석 및 재검증이 필요합니다.**\n\n"
        elif has_not_run:
            status_banner = "> 📌 **[점검 안내] 현재 테스트 실행 전(`Not Run` 미실행) 설계 단계이며, 중요도(High) 또는 특정 기법 항목에 색상 강조 표시가 포함되어 있습니다.**\n\n"
        else:
            status_banner = "> ✅ **[이상 없음] 전체 시트 데이터가 정상 범위로 집계되었습니다.**\n\n"

    prompt = f"""당신은 수석 데이터 분석 및 품질·검증 전문가입니다.
아래 [파이썬 정밀 집계 및 시트별 통합 요약]의 실제 내용만을 근거로, 경영진과 실무 팀장이 한눈에 파악할 수 있는 마크다운 보고서를 작성하세요.

[사용자 질문]
{question}

[파이썬 정밀 집계 및 시트별 통합 요약 (실제 데이터)]
{data_summary}

[중요 해석 지침 (반드시 지킬 것)]
1. **다중 시트 종합 반영:** 보조 시트(`Summary`, `Environment` 등)의 요약 지표·환경 정보와 메인 시트(`Test Cases` 등)의 집계 건수를 구분하여 종합적으로 설명하세요.
2. **색상 강조(`🎨빨강강조` 등)의 정확한 문맥 해석:** 빨간색 배경 표시가 있다고 해서 무조건 '불합격(FAIL)'이라고 단정하지 마세요! 위 집계표의 **[셀 배경색 위치 및 문맥별 실제 의미 분석]**을 확인하여, 그것이 **중요도(High) 표시**인지, **테스트 기법(Error Guessing 등) 강조**인지, 아니면 **미실행(Not Run) 상태의 중점 관리 표시**인지 실제 컬럼과 셀 값 그대로 정확히 설명하세요.
3. **절대 JSON(`{{"summary": ...}}`)으로 출력하지 말고**, 반드시 아래 1번~4번 마크다운 제목(`### 1.` ~ `### 4.`) 형식으로만 작성하세요.

### 1. 시트별 종합 현황 및 핵심 추세 요약
- 메인 데이터 시트의 정확한 건수와 보조 시트(`Summary`, `Environment` 등)의 핵심 지표·환경 구성을 종합하여 2~3문장으로 명확히 요약하세요.

### 2. 주요 강조 표시(색상) 및 관리 대상 항목의 문맥 분석
- 데이터 내 색상 강조(`🎨빨강강조` 등)가 표시된 항목이나 주요 관리 대상을 아래 마크다운 표로 정리하세요. (반드시 표의 각 행마다 줄바꿈 Enter를 넣으세요!)
| 순위 | 식별 ID / 분류 | 색상 강조된 컬럼 및 실제 값 | 실행/판정 상태 | 문맥상 실제 의미 (중요도/기법/불합격 등) |
| :---: | :--- | :--- | :---: | :--- |

### 3. 배경 및 원인 가설 (데이터 근거 기반 추론)
- 특정 항목에 색상 강조가 지정된 이유(예: 안드로이드 화면 회전·강제 종료·상태 전이 등 고위험 시나리오, 중요도 High, 또는 특정 라인 불량 집중)를 데이터의 실제 내용(`Title`, `Category`, `Procedure` 등)을 근거로 2~3가지 추론하되, **반드시 '~할 가능성이 있습니다'** 형태로 작성하세요.

### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)
- **즉시 조치 (24시간 내):**
  - **[담당 검증/현업팀]** (색상 강조된 중요 항목 또는 우선순위 대상에 대한 구체적 실행·점검 행동)
  - **[품질/QA팀]** (미실행 항목 수행 일정 확정 또는 기준 점검 행동)
- **근본 대책 (1개월 내):**
  - **[개발/기술팀]** (고위험 시나리오 자동화 검증 또는 공정 표준화 대책)
- **기대 효과 (목표 KPI):**
  - **[정량 목표]** (조치 완료 시 달성할 커버리지·실행 완료율 또는 품질 개선 목표치 명시)"""

    raw_report, elapsed = call_llm(prompt, mode=mode)

    # 후처리 1: 혹시 JSON으로 응답한 모델이 있으면 마크다운 4단락으로 자동 변환
    md_report = _convert_json_if_needed(raw_report)
    # 후처리 2: 한 줄로 붙은 마크다운 표(||) 줄바꿈 자동 복원
    md_report = _fix_collapsed_markdown_tables(md_report)
    # 후처리 3: 빈 불릿 제거
    md_report = _clean_empty_bullets(md_report)

    final_report = status_banner + md_report

    return {
        "report": final_report,
        "elapsed_time": round(state.get("elapsed_time", 0.0) + elapsed, 2),
    }