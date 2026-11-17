# data_agent/nodes/reporter_node.py
import re
from typing import Any
from data_agent.llm import call_llm


def _is_tc_dataset(data_summary: str, question: str) -> bool:
    """데이터 요약문과 질문을 기반으로 TC(테스트 케이스/검증 절차서) 데이터 여부 판별"""
    combined = f"{data_summary} {question}".lower()
    tc_keywords = [
        "procedure", "pass/fail", "criteria", "actual result",
        "precondition", "auto_rtc", "rtc-", "domain", "mfc dd", "tc",
    ]
    return any(k in combined for k in tc_keywords)


def _convert_markdown_table_lines(text: str) -> str:
    """
    LLM이 마크다운 표(| 컬럼1 | 컬럼2 |)를 출력했을 때,
    파이프(|)를 무작정 지워 'MULTIMEDIA525%'처럼 숫자가 붙는 현상을 방지하고
    깔끔한 불릿('- 항목: 값 — 설명') 형태로 변환합니다.
    """
    out_lines = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith("|") and stripped.endswith("|") and stripped.count("|") >= 3:
            cells = [c.strip() for c in stripped.strip("|").split("|") if c.strip()]
            # 구분선(|---|---|)이나 표 머리글 행은 건너뜀
            if not cells or all(re.match(r"^[:\-]+$", c) for c in cells):
                continue
            header_joined = " ".join(cells)
            if "구분" in header_joined and ("건수" in header_joined or "비중" in header_joined):
                continue
            # 건수(숫자)와 비중(%) 셀이 따로 나뉘어 있으면 '5건 (25%)'로 예쁘게 결합
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
    """
    리포트 후처리:
    1) <br>, <br/> 태그 제거 및 자연스러운 구분자(', ') 치환
    2) 'MULTIMEDIA525%', 'Manual00%'처럼 건수와 비중(%)이 붙어버린 패턴 자동 분리
    3) '수동 테스트는 없습니다' 오판 문구 제거 및 교정
    """
    if not text:
        return ""

    # 1. 마크다운 표 줄 안전 변환
    cleaned = _convert_markdown_table_lines(text)

    # 2. 노출된 <br> 태그 제거
    cleaned = re.sub(r"<br\s*/?>\s*-\s*", ", ", cleaned, flags=re.IGNORECASE)
    cleaned = re.sub(r"<br\s*/?>", ", ", cleaned, flags=re.IGNORECASE)

    # 3. 표 머리글이 본문 줄 앞에 달라붙은 패턴 제거
    cleaned = re.sub(
        r"-\s*구분\s*/\s*관리\s*영역\s*건수\s*및\s*비중\(%\)\s*주요\s*시나리오\s*/\s*대표\s*항목.*?핵심\s*관리\s*포인트\s*",
        "- ",
        cleaned,
    )

    # 4. 'MULTIMEDIA525%' 또는 'Video525%'처럼 문자 뒤에 건수+비중(%)이 공백 없이 붙은 경우 분리
    cleaned = re.sub(
        r"([A-Za-z가-힣_]+?)(\d{1,3})(100|\d{1,2}(?:\.\d+)?)%",
        r"**\1**: \2건 (\3%) — ",
        cleaned,
    )

    # 5. Auto_RTC 오판으로 인한 'Manual 0% 수동 테스트는 없습니다' 줄 제거/교정
    cleaned_lines = []
    for line in cleaned.splitlines():
        s = line.strip()
        if s in ["•", "-", "*", "▶", "• ▶"]:
            continue
        # 'Manual ... 수동 테스트는 없습니다' 단독 불릿은 제거
        if re.search(r"manual.*수동\s*테스트는\s*없습니다", s, flags=re.IGNORECASE):
            continue
        if s in ["- 수동 테스트는 없습니다.", "수동 테스트는 없습니다."]:
            continue
        # Automated 항목이 잘못 생성된 경우 '기능 시나리오 검증(Auto_RTC 시리즈)'로 명칭 보정
        line = re.sub(
            r"\*\*Automated\*\*",
            "**Auto_RTC 프로젝트 검증 시나리오**",
            line,
        )
        cleaned_lines.append(line)

    return "\n".join(cleaned_lines).strip()


def _ensure_complete_report(report_text: str, is_tc: bool = False) -> str:
    """문장 잘림 방지 및 4번 실무 조치 제안 누락 시 자동 보완"""
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
                "- **즉시 조치 (24시간 내 / QA·검증팀)**: Actual Result 내 미기입 항목 및 주의(🔴빨강/🟡노랑 배경) 표시된 TC의 재수행 로그를 확인하고 판정 결과를 업데이트하십시오.\n"
                "- **단기 대책 (1주 내 / 멀티미디어·드라이버 개발팀)**: 비중이 집중된 Video Encode/Decode(VP9, VP8, H.265, H.264) 및 Multiple Instances 시나리오의 예외 스트림·경계값 테스트를 보강하십시오.\n"
                "- **중장기 대책 (1개월 내 / 품질관리 파트)**: 절차서(Procedure) 기반 수동 검증과 자동화 스크립트 수행 대상을 구분할 수 있도록 시트 내 '수행방식(Execution Type)' 표준 컬럼을 추가 관리하십시오."
            )
        else:
            cleaned += (
                "\n\n### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
                "- **즉시 조치 (24시간 내 / 현장관리자)**: 기준치 초과 및 이상 징후 발생 구간의 설비 파라미터와 작업 이력을 즉시 점검하십시오.\n"
                "- **근본 대책 (1개월 내 / 품질·기술팀)**: 반복 발생 유형에 대한 표준 작업서(SOP) 개정 및 예방 점검 주기를 단축하십시오."
            )
    return cleaned


def reporter_node(state: dict[str, Any]) -> dict[str, Any]:
    """
    LangGraph Reporter 노드:
    analyzer_node의 분석 결과(analysis_result)와 파이썬 정밀 집계표(data_summary)를 바탕으로
    최종 마크다운 리포트를 생성합니다.
    """
    question = state.get("question", "현재 데이터의 핵심 현황과 이상 항목, 실무 조치 제안을 정리해 줘.")
    mode = state.get("mode", "local_ollama")
    data_summary = state.get("data_summary", "")
    analysis_result = state.get("analysis_result", {})
    prev_elapsed = float(state.get("elapsed_time", 0.0) or 0.0)

    is_tc = _is_tc_dataset(data_summary, question)

    if is_tc:
        prompt = f"""당신은 수석 QA 및 검증 데이터 분석 전문가입니다.
아래 [파이썬 정밀 집계 및 데이터 요약]과 [1차 분석 인사이트]의 실제 수치만 사용하여 경영진·실무자용 고밀도 검증 리포트를 작성하십시오.

[절대 준수 규칙]
1. 마크다운 표(|)나 HTML 줄바꿈 태그(<br>)를 절대 사용하지 말고, 오직 불릿 리스트(- )로만 깔끔하게 작성하십시오.
2. 건수와 비중(%)을 표기할 때는 반드시 '5건 (25.0%)'처럼 단위와 공백을 띄어 쓰고, 절대 '525%'처럼 숫자를 붙여 쓰지 마십시오.
3. TC ID에 포함된 'Auto_RTC'의 'Auto'는 차량용(Automotive) 프로젝트 식별자일 뿐 자동화 테스트를 뜻하지 않습니다. 또한 본 문서는 상세 수행 절차(Procedure)와 판정 기준(Pass/Fail Criteria)이 적힌 검증 절차서이므로 절대 "Automated vs Manual"로 나누거나 "수동 테스트는 없습니다(Manual 0%)"라고 쓰지 마십시오.
4. 똑같은 6개 TC 목록을 여러 항목에 중복해서 복사·붙여넣기하지 말고, 각 시나리오 유형(인코딩 코덱별, 디코딩, 멀티 인스턴스 등)에 맞는 대표 ID·Title만 2~3개씩 간결히 인용하십시오.

[파이썬 정밀 집계 및 데이터 요약]
{data_summary}

[1차 분석 인사이트]
{analysis_result}

[사용자 질문]
{question}

[필수 출력 목차 (아래 4개 번호와 제목을 정확히 유지)]
### 1. 전체 검증 데이터 요약 및 판정 현황
- 전체 조회 TC 건수, 검증 대상 도메인/모듈, Actual Result 판정 현황(정상/🔴빨강배경·주의/미기입 공란) 핵심 요약

### 2. 도메인 / 서브시스템 & 주요 기능 시나리오별 심층 분포 분석
- **도메인(Domain) 분포**: [영역명] — [건수 (비중%)] | 대표 TC ID·Title | 핵심 검증 포인트
- **서브시스템(App/Module) 분포**: [모듈명(예: Video, MFC DD 등)] — [건수 (비중%)] | 대표 TC ID·Title | 핵심 관리 포인트
- **기능 시나리오 유형별 분포**:
  - 비디오 인코딩(Encode — VP9, VP8, H.265, H.264 등 코덱별 기능 검증) 건수·비중 및 대표 TC
  - 비디오 디코딩(Decode — 단일/복수 스트림 디코딩 검증) 건수·비중 및 대표 TC
  - 다중 인스턴스(Multiple Instances) 및 기타 기능 검증 건수·비중 및 대표 TC
- **검증 집중도 및 절차서(Procedure) 구성의 실무적 의미**: 특정 멀티미디어 코덱/드라이버 집중도의 의미와 절차 기반 검증 시트 관리 포인트

### 3. 품질 리스크 및 취약 구간 진단 (커버리지·판정 결과 관점)
- **결과 판정 및 미기입 리스크**: Actual Result 내 🔴빨강/🟡노랑 배경 표시 항목 또는 결과 미기입 항목 점검
- **시나리오 편중 관점**: 기본 기능(Functionality) 테스트 외에 해상도/프레임레이트 경계값·오류 스트림 주입·장시간 스트레스 테스트 보완 필요성
- **테스트 자산 표준화 관점**: 절차서 내 수동 검증과 자동화 스크립트 연동을 구분하기 위한 메타데이터 관리 제안

### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)
- **즉시 조치 (24시간 내 / QA·검증 엔지니어)**: 주의 표시(배경색) 및 미기입 TC 재확인 조치
- **근본 대책 (1개월 내 / 개발·품질관리 파트)**: 코덱 예외 시나리오 확충 및 TC 관리 템플릿 고도화 방안
"""
    else:
        prompt = f"""당신은 사내 데이터 분석 전문 AI 에이전트입니다.
아래 [파이썬 정밀 집계 및 데이터 요약]과 [1차 분석 인사이트]의 실제 수치만 사용하여 상세 분석 리포트를 작성하십시오.
절대 마크다운 표(|)나 <br> 태그를 사용하지 말고 불릿(-) 리스트와 볼드체(**)로만 작성하십시오. 건수와 비중은 '5건 (25.0%)'처럼 명확히 띄어 쓰십시오.

[파이썬 정밀 집계 및 데이터 요약]
{data_summary}

[1차 분석 인사이트]
{analysis_result}

[사용자 질문]
{question}

[필수 출력 목차]
### 1. 핵심 추세 및 전체 요약 (사내 기준 판정 포함)
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