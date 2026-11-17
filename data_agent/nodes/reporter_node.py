# data_agent/nodes/reporter_node.py
import re
from typing import Any
from data_agent.llm import call_llm


def _is_tc_dataset(data_summary: str, question: str) -> bool:
    combined = f"{data_summary} {question}".lower()
    tc_keywords = [
        "procedure", "pass/fail", "criteria", "actual result",
        "precondition", "pre-condition", "sub system", "test type",
        "auto_rtc", "rtc-", "domain", "mfc dd", "tc", "_polarion",
    ]
    return any(k in combined for k in tc_keywords)


def _convert_markdown_table_lines(text: str) -> str:
    out_lines = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("|") and stripped.endswith("|") and stripped.count("|") >= 3:
            cells = [c.strip() for c in stripped.strip("|").split("|") if c.strip()]
            if not cells or all(re.match(r"^[:\-]+$", c) for c in cells):
                continue
            header_joined = " ".join(cells)
            if "구분" in header_joined and ("건수" in header_joined or "비중" in header_joined):
                continue
            if len(cells) >= 3 and re.match(r"^\d+$", cells[1]) and "%" in cells[2]:
                merged_head = f"- **{cells[0]}**: {cells[1]}건 ({cells[2]})"
                rest = " — ".join(cells[3:])
                out_lines.append(f"{merged_head} | {rest}" if rest else merged_head)
            else:
                out_lines.append("- " + " | ".join(cells))
        else:
            out_lines.append(line)
    return "\n".join(out_lines)


def _sanitize_report_format(text: str) -> str:
    if not text:
        return ""

    cleaned = _convert_markdown_table_lines(text)
    cleaned = re.sub(r"<br\s*/?>\s*-\s*", ", ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"<br\s*/?>", ", ", cleaned, flags=re.IGNORECASE)

    cleaned = re.sub(
        r"-\s*구분\s*/\s*관리\s*영역\s*건수\s*및\s*비중\(%\)\s*주요\s*시나리오\s*/\s*대표\s*항목.*?핵심\s*관리\s*포인트\s*",
        "- ",
        cleaned,
    )

    cleaned = re.sub(
        r"([A-Za-z가-힣_]+?)(\d{1,4})(100|\d{1,2}(?:\.\d+)?)%",
        r"**\1**: \2건 (\3%) — ",
        cleaned,
    )

    cleaned_lines = []
    for line in cleaned.splitlines():
        s = line.strip()
        if s in ["•", "-", "*", "▶", "• ▶"]:
            continue
        if re.search(r"manual\s*:?\s*0건?\s*\(?0%?\)?.*수동\s*테스트는\s*없습니다", s, flags=re.IGNORECASE):
            continue
        if s in ["- 수동 테스트는 없습니다.", "수동 테스트는 없습니다."]:
            continue
        cleaned_lines.append(line)

    return "\n".join(cleaned_lines).strip()


def _ensure_complete_report(report_text: str, is_tc: bool = False) -> str:
    cleaned = _sanitize_report_format(report_text)
    if not cleaned:
        return "⚠️ 리포트 생성 결과가 비어 있습니다."

    if not cleaned.endswith((".", "다", "요", "함", "음", ")", "]", "%")):
        last_newline = cleaned.rfind("\n")
        if last_newline > 0 and len(cleaned) - last_newline < 45:
            cleaned = cleaned[:last_newline].strip()

    if "4." not in cleaned and "조치 제안" not in cleaned:
        if is_tc:
            cleaned += (
                "\n\n### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
                "- **즉시 조치 (24시간 내 / QA·검증팀)**: 판정 미기입 항목 및 주의(🔴빨강/🟡노랑 배경) 표시된 TC의 재수행 로그를 확인하고 결과를 업데이트하십시오.\n"
                "- **단기 대책 (1주 내 / 서브시스템·드라이버 개발팀)**: 주요 도메인 및 서브시스템(Sub System / IP) 집중 구간의 예외·경계값 테스트 케이스를 보강하십시오.\n"
                "- **중장기 대책 (1개월 내 / 품질관리 파트)**: 수동 절차 검증과 자동화 수행 대상의 분류 기준을 표준화하고 중점 모듈의 회귀 테스트 커버리지를 확대하십시오."
            )
        else:
            cleaned += (
                "\n\n### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
                "- **즉시 조치 (24시간 내 / 현장관리자)**: 기준치 초과 및 이상 징후 발생 구간의 설비 파라미터와 작업 이력을 즉시 점검하십시오.\n"
                "- **근본 대책 (1개월 내 / 품질·기술팀)**: 반복 발생 유형에 대한 표준 작업서(SOP) 개정 및 예방 점검 주기를 단축하십시오."
            )
    return cleaned


def reporter_node(state: dict[str, Any]) -> dict[str, Any]:
    question = state.get("question", "현재 데이터의 핵심 현황과 이상 항목, 실무 조치 제안을 정리해 줘.")
    mode = state.get("mode", "local_ollama")
    data_summary = state.get("data_summary", "")
    analysis_result = state.get("analysis_result", {})
    defect_threshold = float(state.get("defect_threshold", 1.5) or 1.5)
    prev_elapsed = float(state.get("elapsed_time", 0.0) or 0.0)

    is_tc = _is_tc_dataset(data_summary, question)
    has_automated_col = "'automated'" in data_summary.lower() or "[automated]" in data_summary.lower()

    if is_tc:
        auto_rule = (
            "데이터에 'Automated' 또는 'Test Type' 집계표가 있으므로 파이썬이 집계한 실제 건수와 비중(%)을 그대로 인용하여 분석하십시오."
            if has_automated_col
            else "본 시트에서 TC ID의 'Auto_RTC'는 차량용(Automotive) 프로젝트 식별자이므로 절대 '수동 테스트는 없습니다(Manual 0%)'라고 단정하지 마십시오."
        )
        prompt = f"""당신은 수석 QA 및 검증 데이터 분석 전문가입니다.
아래 [파이썬 정밀 집계 및 데이터 요약]과 [1차 분석 인사이트]의 실제 수치만 사용하여 경영진·실무자용 고밀도 검증 리포트를 작성하십시오.

[절대 준수 규칙]
1. 마크다운 표(|)나 HTML 줄바꿈 태그(<br>)를 절대 사용하지 말고, 오직 불릿 리스트(- )로만 깔끔하게 작성하십시오.
2. 건수와 비중(%)을 표기할 때는 반드시 '47건 (47.0%)'처럼 단위와 공백을 띄어 쓰고, 절대 '525%'처럼 숫자를 붙여 쓰지 마십시오.
3. {auto_rule}
4. 똑같은 대표 TC 목록을 여러 항목에 중복해서 복사·붙여넣기하지 말고, 각 도메인·서브시스템·시나리오에 맞는 대표 ID·Title을 간결히 인용하십시오.

[파이썬 정밀 집계 및 데이터 요약]
{data_summary}

[1차 분석 인사이트]
{analysis_result}

[사용자 질문]
{question}

[필수 출력 목차 (아래 4개 번호와 제목을 정확히 유지)]
### 1. 전체 검증 데이터 요약 및 판정 현황
- 전체 조회 TC 건수, 검증 대상 도메인(Domain)·서브시스템(Sub System/App), 판정 및 수행 현황 핵심 요약

### 2. 도메인 / 서브시스템 & 주요 기능 시나리오별 심층 분포 분석
- **도메인(Domain) 분포**: 각 영역별 건수 (비중%) | 대표 TC ID·Title | 핵심 검증 포인트
- **서브시스템(Sub System / App / IP) 분포**: 상위 모듈별 건수 (비중%) | 핵심 관리 포인트
- **주요 시나리오 및 테스트 유형(Test Type / Automated 등) 분포**: 파이썬 집계표에 기반한 기능 시나리오 묶음별 건수·비중 및 실무적 의미

### 3. 품질 리스크 및 취약 구간 진단 (커버리지·판정 결과 관점)
- **결과 판정 및 미기입 리스크**: 불합격/주의(🔴빨강/🟡노랑 배경셀) 또는 미기입 항목 점검
- **모듈 편중 및 예외 검증 관점**: 특정 도메인·서브시스템 편중 현상과 경계값·스트레스·예외 처리 테스트 보완 필요성

### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)
- **즉시 조치 (24시간 내 / QA·검증 엔지니어)**: 주의 표시 및 미기입 TC 재확인 조치
- **근본 대책 (1개월 내 / 개발·품질관리 파트)**: 취약 서브시스템 커버리지 확충 및 TC 자산 표준화 방안
"""
    else:
        prompt = f"""당신은 사내 데이터 분석 전문 AI 에이전트입니다.
아래 [파이썬 정밀 집계 및 데이터 요약](사내 품질 기준값: {defect_threshold}%)과 [1차 분석 인사이트]의 실제 수치만 사용하여 상세 분석 리포트를 작성하십시오.
절대 마크다운 표(|)나 <br> 태그를 사용하지 말고 불릿(-) 리스트와 볼드체(**)로만 작성하십시오. 건수와 비중은 '5건 (25.0%)'처럼 명확히 띄어 쓰십시오.

[파이썬 정밀 집계 및 데이터 요약]
{data_summary}

[1차 분석 인사이트]
{analysis_result}

[사용자 질문]
{question}

[필수 출력 목차]
### 1. 핵심 추세 및 전체 요약 (사내 기준 {defect_threshold}% 판정 포함)
### 2. 주요 항목별 분포 및 이상 징후 심층 분석 (건수 및 비중% 명시)
### 3. 데이터 기반 원인 가설 (다각도 관점 추론)
### 4. 실무 조치 제안 (즉시 조치 24h / 근본 대책 1개월 R&R 및 기대 효과)
"""

    raw_report, elapsed = call_llm(prompt, mode=mode)
    final_report = _ensure_complete_report(raw_report, is_tc=is_tc)

    return {
        "report": final_report,
        "elapsed_time": round(prev_elapsed + elapsed, 2),
    }