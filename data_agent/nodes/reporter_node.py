import json
import re
from typing import Any
from data_agent.llm import call_llm


def _detect_data_domain(data_summary: str) -> str:
    """
    업로드된 데이터의 실제 컬럼과 집계 내용을 분석하여 도메인 성격을 판별합니다:
    - 'manufacturing': 생산량·불량수 중심 제조 공정 데이터
    - 'tc_procedure_spec': 실행 결과(Pass/Fail)나 색상 강조 없이 Sub System, Domain, Automated, Procedure 등으로 구성된 TC 절차서/명세서
    - 'tc_execution_result': 실제 Pass/Fail 실행 결과나 색상 강조가 포함된 검증 결과 시트
    - 'general_doc': 워드(.docx) 문단이나 일반 업무 문서
    """
    s_lower = data_summary.lower()
    if "생산량" in s_lower and "불량수" in s_lower:
        return "manufacturing"
    if "문단id" in s_lower or "문서 섹션(목차)" in s_lower:
        return "general_doc"
    if any(k in s_lower for k in ["sub system", "automated", "pass/fail criteria", "pre-condition", "precondition", "procedure", "test type"]):
        has_color_table = "셀 배경색(강조 표시) 위치 및 문맥별 실제 의미 분석" in data_summary
        has_exec_status = "실행 및 판정 상태(pass / fail / not run) 정밀 집계" in s_lower
        if has_color_table or has_exec_status:
            return "tc_execution_result"
        return "tc_procedure_spec"
    return "general_doc"


def _convert_json_if_needed(raw_text: str) -> str:
    """만약 LLM이 마크다운 대신 JSON으로 응답했을 경우 4단락 마크다운 보고서로 자동 변환합니다."""
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
                model_header + "### 1. Executive Snapshot (전체 핵심 지표 요약 및 종합 현황)",
                f"- {summary}",
            ]
            for t in trends:
                md_parts.append(f"- {t}")

            md_parts.append("\n### 2. 도메인 / 서브시스템 & 주요 주제별 심층 분포")
            for a in anomalies:
                md_parts.append(f"- {a}")

            md_parts.append("\n### 3. 기술적 리스크 및 원인 가설 (데이터 근거 기반 추론)")
            if causes:
                for c in causes:
                    md_parts.append(f"- {c}")
            else:
                md_parts.append("- 집계된 상위 도메인·서브시스템 및 수동(Manual) 테스트 비중을 고려할 때 인터페이스 경계 조건에서 리스크가 발생할 가능성이 있습니다.")

            md_parts.append("\n### 4. 실무 조치 제안 (시점별 로드맵 및 부서별 R&R)")
            if actions:
                for act in actions:
                    md_parts.append(f"- {act}")
            else:
                md_parts.append("- **즉시 조치 (24시간 내):**\n  - **[담당 검증/QA팀]** 핵심 서브시스템 및 수동(Manual) 검증 항목의 우선순위 점검")
                md_parts.append("- **근본 대책 (1개월 내):**\n  - **[개발/자동화팀]** 반복 기능 테스트의 자동화(Auto) 전환 및 표준 검증 절차 고도화")

            return "\n".join(md_parts)
        except Exception:
            pass
    return raw_text


def _fix_collapsed_markdown_tables(text: str) -> str:
    """마크다운 표가 줄바꿈 없이 한 줄로 이어붙었을 때 줄바꿈을 자동 복원합니다."""
    fixed = re.sub(r"\|\s*\|\s*(?=(?::?-+:?|\d+\s*\||TC-|REQ-|Auto_|•|[A-Za-z가-힣]))", "|\n| ", text)
    fixed = re.sub(r"([.다요음함)])\s*(\|\s*(?:순위|항목|구분|분류|테스트)\s*\|)", r"\1\n\n\2", fixed)
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
    """
    데이터의 실제 성격(생산 공정 vs TC 절차서 명세 vs TC 실행 결과 vs 일반 문서)과
    실제 존재하는 컬럼만을 반영하여 Groq 스타일의 Executive Snapshot 보고서를 생성합니다.
    (데이터에 없는 색상 강조, 중요도 High, 가짜 PASS/FAIL, 없는 Summary 시트 언급을 원천 차단!)
    """
    mode = state.get("mode", "mock")
    data_summary = state.get("data_summary", "")
    question = state.get("question", "핵심 지표 현황과 분류별 비중, 리스크 가설 및 부서별 실행 계획을 정리해 줘.")

    domain_type = _detect_data_domain(data_summary)
    has_color_cells = "셀 배경색(강조 표시) 위치 및 문맥별 실제 의미 분석" in data_summary
    has_aux_sheets = "[보조 시트/내장표 목록]: 없음" not in data_summary

    if domain_type == "manufacturing":
        matches = re.findall(r"평균 불량률\(%\)\s*\|\s*(\d+\.\d+)%|불량률\(%\)\s*(\d+\.\d+)%", data_summary)
        rates = [float(m[0] or m[1]) for m in matches if (m[0] or m[1])]
        worst_rate = max(rates) if rates else 0.0
        if worst_rate > 1.5:
            status_banner = f"> 🚨 **[확인 필요] 불량률이 사내 기준(1.5%)을 초과한 라인(최고 `{worst_rate:.2f}%`)이 있습니다. 주간회의 안건으로 올리세요.**\n\n"
        else:
            status_banner = "> ✅ **[이상 없음] 모든 라인의 불량률이 사내 기준(1.5%) 이내입니다. 정기 보고로만 처리하세요.**\n\n"
    elif domain_type == "tc_procedure_spec":
        status_banner = (
            "> 📋 **[TC 검증 절차서/명세 분석 완료] 본 파일은 테스트 실행 전 단계인 표준 검증 절차서(Specification)입니다. "
            "도메인·서브시스템 분포, 자동화(Auto/Manual) 현황 및 중점 관리 시나리오를 중심으로 보고합니다.**\n\n"
        )
    elif domain_type == "tc_execution_result":
        has_real_fail = bool(re.search(r"실행 실패|불합격\(FAIL\)|\bFAIL\b", data_summary))
        if has_real_fail:
            status_banner = "> 🚨 **[확인 필요] 테스트 실행 결과 불합격(FAIL) 또는 주의 항목이 감지되었습니다. 원인 분석 및 재검증이 필요합니다.**\n\n"
        else:
            status_banner = "> 📌 **[점검 안내] 테스트케이스 상태 및 강조 표시 항목에 대한 문맥 분석 결과입니다.**\n\n"
    else:
        status_banner = "> 📊 **[문서·데이터 종합 분석 완료] 업로드된 데이터의 핵심 분류와 주요 내용을 기반으로 작성된 보고서입니다.**\n\n"

    # 데이터에 없는 내용을 절대 지어내지 못하도록 동적 금지 규칙 생성
    strict_negative_rules = [
        "1. **[절대 환각 금지]** 아래 [1. 분석 대상 정밀 집계표]에 실제로 적혀 있는 숫자와 건수, 비율(%), 컬럼명, ID만 사용하세요. 표에 없는 숫자를 절대 임의로 지어내지 마세요."
    ]
    if not has_color_cells:
        strict_negative_rules.append(
            "2. **[색상/중요도 날조 절대 금지]** 현재 데이터에는 배경색 강조 셀이나 중요도(High) 표시가 **전혀 없습니다(0건)**. 보고서에 '색상 강조', '빨간색 배경', '중요도(High)'라는 단어를 **단 한 번도 쓰지 마세요**."
        )
    if domain_type == "tc_procedure_spec":
        strict_negative_rules.append(
            "3. **[가짜 PASS/FAIL 판정 날조 절대 금지]** 현재 파일은 테스트 실행 전 '절차서'이므로 실제 성공(PASS)이나 실패(FAIL) 결과가 없습니다. 특정 TC가 PASS 또는 FAIL 했다거나 '실패 원인은 Checksum...' 같은 가짜 실행 결과를 절대 꾸며내지 마세요. 반드시 실제 존재하는 `Domain`, `Sub System`, `Test Type`, `Automated(Auto/Manual)`, `Title 주제 분류`만 분석하세요."
        )
    if not has_aux_sheets:
        strict_negative_rules.append(
            "4. **[없는 보조 시트 언급 금지]** 현재 파일에는 `Summary`나 `Environment` 보조 시트가 없습니다. 존재하지 않는 보조 시트 이야기를 절대 쓰지 마세요."
        )

    negative_rules_str = "\n".join(strict_negative_rules)

    prompt = f"""당신은 수석 데이터 분석 및 QA·품질 전략 전문가입니다.
아래 [파이썬 정밀 집계표 및 실제 데이터]만을 근거로, 경영진과 팀장이 한눈에 의사결정할 수 있는 **Executive Snapshot 스타일의 보고서**를 한국어로 작성하세요.

[사용자 질문]
{question}

[절대 준수 규칙 (위반 시 보고서 실패 처리)]
{negative_rules_str}
- JSON(`{{"summary": ...}}`)으로 출력하지 말고, 반드시 아래 1번~4번 마크다운 제목(`### 1.` ~ `### 4.`)과 마크다운 표 형식을 정확히 지키세요.

[파이썬 정밀 집계표 및 실제 데이터]
{data_summary}

---
[보고서 필수 작성 양식 (반드시 아래 4개 섹션과 표 구조를 그대로 작성할 것)]

### 1. Executive Snapshot (전체 핵심 지표 요약표 및 종합 현황)
- 전체 데이터 건수와 주요 분류 항목별(예: `Domain`, `Sub System`, `Test Type`, `Automated`, `주요 기술 키워드 그룹` 또는 `라인`, `불량유형`, `주차` 등 실제 집계표에 있는 항목) 핵심 건수·비중(%)을 아래 마크다운 표로 일목요연하게 정리하세요.
| 분류 기준 (항목) | 세부 구분 | 건수 (건) | 비중 (%) | 핵심 요약 및 비고 |
| :--- | :--- | :---: | :---: | :--- |
- 표 아래에 전체 데이터의 핵심 구조적 특징을 2~3개 불릿으로 명쾌하게 요약하세요.

### 2. 도메인 / 서브시스템 & 주요 주제별 심층 분포 분석
- 집계표의 상위 비중 영역(예: `MULTIMEDIA`, `SFI CDD`, `Linux`, `SFI` 등 서브시스템/도메인 분포, `Auto vs Manual` 자동화 현황, `Title` 핵심 기술 키워드 묶음, 또는 공정 라인별 현황)을 교차 분석하여 아래 마크다운 표로 정리하세요.
| 구분 / 관리 영역 | 건수 및 비중(%) | 주요 시나리오 / 대표 항목 (실제 ID·Title 인용) | 핵심 관리 포인트 |
| :--- | :---: | :--- | :--- |
- 표 아래에 주요 영역 집중도 및 자동화율(`Auto` vs `Manual`)이 실무적으로 의미하는 바를 2~3개 불릿으로 설명하세요.

### 3. 자동화 여부 · 테스트 유형별 중점 관리 시나리오 및 리스크 가설
- 실제 데이터에 나타난 특징(예: `Auto` 자동화 테스트와 `Manual` 수동 테스트의 역할 차이, `Cybersecurity Test` 등 특수 테스트 유형, 하드웨어/인터페이스/부팅/코덱/멀티미디어 등 실제 절차 내용, 또는 특정 라인 불량 집중)을 표나 불릿으로 정리하고, 테스트 수행 시 유의해야 할 기술적 리스크 가설을 3가지 제시하세요. (단정하지 말고 **'~할 가능성이 있습니다'** 또는 **'~에 대한 중점 검증이 필요합니다'** 형태로 실제 데이터의 ID나 키워드를 인용해 서술)

### 4. 실무 조치 제안 (시점별 로드맵 및 부서별 R&R)
- **즉시 조치 (24시간 내 / 단기 실행):**
  - **[담당 검증/QA팀]** (데이터 내 수동 `Manual` 항목, 고비중 서브시스템 등 우선순위가 높은 시나리오에 대한 구체적 점검 행동)
  - **[개발/엔지니어링팀]** (핵심 도메인·인터페이스·사전조건 `Pre-condition` 환경 세팅 및 리스크 예방 행동)
- **근본 대책 (1개월 내 / 중장기 고도화):**
  - **[테스트 자동화/기술팀]** (수동 테스트의 자동화 스크립트 전환, 커버리지 확대 또는 공정 표준화 대책)
- **기대 효과 (목표 KPI):**
  - **[정량 목표]** (조치 완료 시 달성할 자동화율 목표, 검증 커버리지 또는 품질 안정화 목표 수치 명시)"""

    raw_report, elapsed = call_llm(prompt, mode=mode)

    md_report = _convert_json_if_needed(raw_report)
    md_report = _fix_collapsed_markdown_tables(md_report)
    md_report = _clean_empty_bullets(md_report)

    final_report = status_banner + md_report

    return {
        "report": final_report,
        "elapsed_time": round(state.get("elapsed_time", 0.0) + elapsed, 2),
    }