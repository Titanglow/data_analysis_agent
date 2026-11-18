# data_agent/nodes/reporter_node.py
import re
from typing import Any
from data_agent.llm import call_llm


def _is_tc_dataset(data_summary: str, question: str) -> bool:
    combined = f"{data_summary} {question}".lower()
    tc_keywords = [
        "procedure", "pass/fail", "criteria", "actual result",
        "precondition", "pre-condition", "sub system", "test type",
        "priority", "test technique", "execution status", "not run",
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


def _strip_dumped_sample_rows(text: str) -> str:
    """
    Groq 등이 4번 섹션 뒤에 '5. 상세 행 샘플 데이터', '6. 상세 행 샘플 데이터 (상위 40건)' 등을 붙이며
    TC-001, TC-002 원본 행을 줄줄이 복사하다가 끊기는 현상을 원천 제거합니다.
    """
    lines = text.splitlines()
    kept_lines = []
    in_sec_4 = False

    for line in lines:
        s = line.strip()
        if re.match(r"^(?:#+\s*)?4\.\s*", s):
            in_sec_4 = True

        # 5번 이상 섹션 제목이나 '상세 행 샘플 데이터' 제목이 등장하면 그 즉시 뒤쪽을 모두 잘라냄!
        if re.match(r"^(?:#+\s*)?[5-9]\.\s*", s) or "상세 행 샘플 데이터" in s or "상세 행 데이터" in s:
            break

        # 4번 섹션 이후에 '- TC-001' 처럼 개별 원본 행을 덤프하기 시작하면 중단
        if in_sec_4 and re.match(r"^[-*•]\s*TC-\d{3,}", s):
            break

        kept_lines.append(line)

    return "\n".join(kept_lines).strip()


def _sanitize_report_format(text: str, has_color_cells: bool = False) -> str:
    if not text:
        return ""

    # 1. 불필요한 5~6번 샘플 데이터 덤프 구간 먼저 제거
    cleaned = _strip_dumped_sample_rows(text)
    cleaned = _convert_markdown_table_lines(cleaned)
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
        if not has_color_cells:
            line = re.sub(r"\(?🔴\s*빨강(?:/🟡\s*노랑)?\s*배경(?:셀)?\)?", "", line)
            line = re.sub(r"주의\s*표시\(🔴빨강\)로\s*분류된\s*항목이\s*있으므로,?\s*", "", line)
            if re.search(r"(?:빨강|노랑|배경셀|색상\s*강조).*0\s*건", line):
                continue
        cleaned_lines.append(line)

    return "\n".join(cleaned_lines).strip()


def _ensure_complete_report(
    report_text: str,
    verdict_banner: str,
    is_tc: bool = False,
    has_color_cells: bool = False,
) -> str:
    cleaned = _sanitize_report_format(report_text, has_color_cells=has_color_cells)
    if not cleaned:
        cleaned = "⚠️ 리포트 생성 결과가 비어 있습니다."

    # 마지막 줄이 문장 중간에 끊겼으면 깔끔하게 마침표 줄까지만 유지
    if not cleaned.endswith((".", "다", "요", "함", "음", ")", "]", "%", "*")):
        last_newline = cleaned.rfind("\n")
        if last_newline > 0 and len(cleaned) - last_newline < 65:
            cleaned = cleaned[:last_newline].strip()

    if "4." not in cleaned and "조치 제안" not in cleaned:
        if is_tc:
            cleaned += (
                "\n\n### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
                "- **즉시 조치 (24시간 내 / QA·검증팀)**: 미실행(Not Run) 및 결과 미기입(공란/N/A) 항목의 우선순위를 확인하고 중점 카테고리부터 검증을 수행하십시오.\n"
                "- **단기 대책 (1주 내 / 개발·검증 파트)**: 비중이 집중된 주요 기능 구간의 경계값·예외 처리 테스트 케이스를 보강하십시오.\n"
                "- **중장기 대책 (1개월 내 / 품질관리 파트)**: 중점 모듈의 회귀 테스트 자동화율을 높이고 부서별 TC 양식 메타데이터를 표준화하십시오."
            )
        else:
            cleaned += (
                "\n\n### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
                "- **즉시 조치 (24시간 내 / 현장관리자)**: 사내 기준 초과 및 이상 징후 발생 구간의 설비 파라미터와 작업 이력을 즉시 점검하십시오.\n"
                "- **근본 대책 (1개월 내 / 품질·기술팀)**: 반복 발생 유형에 대한 표준 작업서(SOP) 개정 및 예방 점검 주기를 단축하십시오."
            )

    if verdict_banner:
        cleaned = f"> **{verdict_banner}**\n\n{cleaned}"
    return cleaned


def reporter_node(state: dict[str, Any]) -> dict[str, Any]:
    question = state.get("question", "현재 데이터의 핵심 현황과 이상 항목, 실무 조치 제안을 정리해 줘.")
    mode = state.get("mode", "local_ollama")
    data_summary = state.get("data_summary", "")
    analysis_result = state.get("analysis_result", {})
    defect_threshold = float(state.get("defect_threshold", 1.5) or 1.5)
    prev_elapsed = float(state.get("elapsed_time", 0.0) or 0.0)

    verdict_banner = ""
    for line in data_summary.splitlines():
        if line.startswith("[사내 기준 자동 판정]:"):
            verdict_banner = line.replace("[사내 기준 자동 판정]:", "").strip()
            break

    is_tc = _is_tc_dataset(data_summary, question)
    has_color_cells = "셀 배경색(강조 표시) 위치 및 문맥별 실제 의미 분석" in data_summary

    if is_tc:
        color_instruction = (
            "- [셀 배경색 위치 및 문맥별 실제 의미 분석] 표의 내용을 바탕으로 색상이 강조된 이유를 설명하십시오."
            if has_color_cells
            else "- 본 엑셀 파일에는 셀 배경색(빨강/노랑 등)이 칠해진 셀이 단 1건도 없으므로, 절대 '빨강 배경', '노랑 배경', '주의 표시(🔴)'라는 단어를 언급하지 마십시오."
        )

        prompt = f"""당신은 수석 QA 및 검증 데이터 분석 전문가입니다.
반드시 아래 [파이썬 정밀 집계 및 데이터 요약]에 적힌 **실제 컬럼명과 실제 계산된 건수·비중(%)**만 그대로 사용하여 고품질 분석 리포트를 작성하십시오.

[절대 준수 규칙 — 위반 시 심각한 오류]
1. **[환각 절대 금지]** [파이썬 정밀 집계 및 데이터 요약]에 없는 분류 이름이나 없는 숫자를 절대 스스로 지어내지 마십시오.
2. **[원본 샘플 행 나열 절대 금지]** 절대 `5.`, `6.` 같은 추가 목차를 만들거나 `[대표 행 참고 샘플]`의 개별 TC 행(TC-001, TC-002 등)을 줄줄이 복사·나열하지 마십시오. 오직 아래 **1번~4번 필수 목차까지만 작성하고 즉시 종료**하십시오.
3. 마크다운 표(|)나 HTML 줄바꿈 태그(<br>)를 절대 사용하지 말고, 불릿 리스트(- )로만 깔끔하게 작성하십시오.
4. 건수와 비중(%)을 표기할 때는 반드시 '48건 (76.2%)'처럼 단위와 공백을 띄어 쓰십시오.
5. {color_instruction}
6. 현재 사용자가 설정한 사내 관리 기준값은 **{defect_threshold}%**입니다. 파이썬 집계표에서 `사내 기준({defect_threshold}%) 판정`이 초과(🚨)이거나 기준 이상(🎯)으로 표시된 항목을 중심으로 깊이 있게 설명하십시오.

[파이썬 정밀 집계 및 데이터 요약]
{data_summary}

[1차 분석 인사이트]
{analysis_result}

[사용자 질문]
{question}

[필수 출력 목차 (반드시 아래 1~4번만 작성하고 5번 이후는 절대 만들지 마십시오)]
### 1. 전체 검증 데이터 요약 및 사내 기준({defect_threshold}%) 판정 현황
- 전체 조회 건수, 실제 존재하는 주요 컬럼 구성, 사내 기준({defect_threshold}%) 대비 미기입(공란/N/A) 또는 판정 상태 핵심 요약

### 2. 실제 분류 컬럼별 분포 & 주요 시나리오 그룹핑 심층 분석
- **실제 분류 컬럼별 분포**: 파이썬 집계표에 나온 실제 분류 열(예: Category, Priority, Test Technique, Domain, Sub System 등)의 항목별 건수 (비중%) 및 사내 기준({defect_threshold}%) 이상 집중 항목 정리
- **시나리오 키워드 그룹핑 분포**: [유사 시나리오·키워드 그룹핑 집계] 표에 나온 상위 시나리오 묶음별 건수 (비중%)와 대표 사례 1~2건만 짧게 인용

### 3. 데이터 품질 리스크 및 취약 구간 진단 (사내 기준 {defect_threshold}% 관점)
- **미기입(공란) 및 미실행(Not Run / N/A) 점검**: [전체 컬럼별 데이터 작성 현황] 표에서 미기입률이 사내 기준({defect_threshold}%)을 초과한 컬럼의 구체적 건수·비중 지적
- **검증 편중 및 예외 테스트 보완 관점**: 특정 분류/기법 편중 현상과 경계값·예외 처리 테스트 보강 필요성

### 4. 실무 조치 제안 (시점별 로드맵 및 R&R)
- **즉시 조치 (24시간 내 / QA·검증 엔지니어)**: 사내 기준({defect_threshold}%) 초과 미기입/미실행 항목 정비 및 고위험 항목 우선 검증
- **근본 대책 (1개월 내 / 개발·품질관리 파트)**: 취약 구간 테스트 케이스 확충 및 양식 표준화 방안
"""
    else:
        prompt = f"""당신은 사내 데이터 분석 전문 AI 에이전트입니다.
아래 [파이썬 정밀 집계 및 데이터 요약](적용된 사내 품질 기준값: {defect_threshold}%)의 실제 수치만 사용하여 상세 분석 리포트를 작성하십시오.
집계표에 없는 숫자를 절대 지어내지 마시고, 원본 샘플 행을 나열하지 마십시오. 반드시 아래 1~4번 목차만 불릿(-) 리스트로 작성하고 종료하십시오.

[파이썬 정밀 집계 및 데이터 요약]
{data_summary}

[1차 분석 인사이트]
{analysis_result}

[사용자 질문]
{question}

[필수 출력 목차]
### 1. 핵심 추세 및 사내 기준({defect_threshold}%) 판정 요약
### 2. 주요 항목별 분포 및 기준({defect_threshold}%) 초과 구간 분석 (건수 및 비중% 명시)
### 3. 데이터 기반 원인 가설 (다각도 관점 추론)
### 4. 실무 조치 제안 (즉시 조치 24h / 근본 대책 1개월 R&R 및 기대 효과)
"""

    raw_report, elapsed = call_llm(prompt, mode=mode)
    final_report = _ensure_complete_report(
        raw_report, verdict_banner=verdict_banner, is_tc=is_tc, has_color_cells=has_color_cells
    )

    return {
        "report": final_report,
        "elapsed_time": round(prev_elapsed + elapsed, 2),
    }