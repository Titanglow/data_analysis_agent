# data_agent/nodes/reporter_node.py
import re
import pandas as pd
from data_agent.llm import invoke_llm


def _detect_dataset_type(df: pd.DataFrame, summary_text: str = "") -> str:
    """데이터프레임 컬럼과 요약문을 기반으로 데이터 유형(tc / manufacturing / general) 판별"""
    cols_str = " ".join([str(c).lower() for c in df.columns]) if df is not None and not df.empty else ""
    combined = f"{cols_str} {summary_text.lower()}"

    tc_keywords = ["procedure", "pass/fail", "criteria", "actual result", "precondition", "tc", "test", "domain", "rtc"]
    mfg_keywords = ["생산량", "불량수", "불량률", "생산라인", "설비가동률", "교대"]

    if any(k in combined for k in tc_keywords):
        return "tc"
    if any(k in combined for k in mfg_keywords):
        return "manufacturing"
    return "general"


def _build_tc_structured_context(df: pd.DataFrame) -> str:
    """
    TC 절차 엑셀 데이터를 파이썬 규칙 기반으로 사전 집계합니다.
    - Auto_RTC 등 ID 접두사를 자동화(Automated)로 오판하지 않도록 실제 컬럼 기반으로만 집계
    - 건수와 비중을 'N건 (XX.X%)' 형식으로 분리하여 525% 같은 숫자 붙음 방지
    - 그룹별 실제 대표 TC를 중복 없이 추출
    """
    if df is None or df.empty:
        return "조회된 TC 데이터가 없습니다."

    total_cnt = len(df)
    lines = [f"[파이썬 정밀 집계 요약 — 총 {total_cnt}건]"]

    # 컬럼 매핑 탐색
    col_map = {str(c).strip().lower(): c for c in df.columns}
    domain_col = next((col_map[k] for k in col_map if "domain" in k or "도메인" in k or "영역" in k), None)
    app_col = next((col_map[k] for k in col_map if k in ["app", "module", "subsystem", "앱", "모듈", "서브시스템"] or "app" in k), None)
    id_col = next((col_map[k] for k in col_map if k in ["id", "tc id", "tc_id", "test id"] or "id" in k), None)
    title_col = next((col_map[k] for k in col_map if "title" in k or "제목" in k or "시나리오" in k or "name" in k), None)
    actual_col = next((col_map[k] for k in col_map if "actual" in k or "결과" in k or "판정" in k or "result" in k), None)
    auto_col = next((col_map[k] for k in col_map if k in ["automation", "auto/manual", "execution type", "자동화여부", "수행방식"]), None)

    def _get_rep_items(sub_df: pd.DataFrame, max_items: int = 3) -> str:
        reps = []
        for _, row in sub_df.head(max_items).iterrows():
            tid = str(row[id_col]).strip() if id_col and pd.notna(row[id_col]) else ""
            ttitle = str(row[title_col]).strip() if title_col and pd.notna(row[title_col]) else ""
            item_str = f"{tid} {ttitle}".strip()
            if item_str and item_str != "-":
                reps.append(item_str[:70])
        return " / ".join(reps) if reps else "대표 항목 없음"

    # 1) 도메인별 집계
    if domain_col:
        lines.append("\n1. 도메인(Domain)별 분포:")
        vc = df[domain_col].fillna("미분류").astype(str).str.strip().value_counts()
        for val, cnt in vc.head(6).items():
            pct = (cnt / total_cnt) * 100
            sub = df[df[domain_col].fillna("미분류").astype(str).str.strip() == val]
            lines.append(f"  - {val}: {cnt}건 ({pct:.1f}%) | 대표 TC: {_get_rep_items(sub, 3)}")

    # 2) 서브시스템 / App별 집계
    if app_col and app_col != domain_col:
        lines.append("\n2. 서브시스템(App/Module)별 분포:")
        vc = df[app_col].fillna("미분류").astype(str).str.strip().value_counts()
        for val, cnt in vc.head(6).items():
            pct = (cnt / total_cnt) * 100
            sub = df[df[app_col].fillna("미분류").astype(str).str.strip() == val]
            lines.append(f"  - {val}: {cnt}건 ({pct:.1f}%) | 대표 TC: {_get_rep_items(sub, 3)}")

    # 3) 기능 시나리오 유형별 분류 (Title/Procedure 키워드 기반)
    text_Series = pd.Series([""] * total_cnt, index=df.index)
    for c in [title_col, id_col]:
        if c:
            text_Series = text_Series + " " + df[c].fillna("").astype(str)

    scenario_buckets = {
        "Video Encode (VP9/VP8/H.264/H.265 등 인코딩)": text_Series.str.contains("encode", case=False, na=False),
        "Video Decode (디코딩 및 재생)": text_Series.str.contains("decode", case=False, na=False),
        "Multi-Instance / 동시성 및 스트레스": text_Series.str.contains("multiple|multi|concurrent|stress", case=False, na=False),
    }
    matched_any = pd.Series([False] * total_cnt, index=df.index)
    lines.append("\n3. 주요 기능 시나리오 유형별 분포:")
    for label, mask in scenario_buckets.items():
        cnt = int(mask.sum())
        if cnt > 0:
            matched_any = matched_any | mask
            pct = (cnt / total_cnt) * 100
            lines.append(f"  - {label}: {cnt}건 ({pct:.1f}%) | 대표 TC: {_get_rep_items(df[mask], 2)}")
    other_cnt = int((~matched_any).sum())
    if other_cnt > 0:
        pct = (other_cnt / total_cnt) * 100
        lines.append(f"  - 기타 일반 기능/인터페이스 검증: {other_cnt}건 ({pct:.1f}%) | 대표 TC: {_get_rep_items(df[~matched_any], 2)}")

    # 4) Actual Result / 판정 및 수행 현황
    if actual_col:
        s_act = df[actual_col].fillna("-").astype(str).str.strip()
        fail_mask = s_act.str.contains("빨강|불합격|fail|ng|주의|error", case=False, na=False)
        pass_mask = s_act.str.contains("초록|정상|pass|ok", case=False, na=False) & (~fail_mask)
        warn_mask = s_act.str.contains("노랑|확인필요|hold|pending|inconc", case=False, na=False)
        empty_mask = s_act.isin(["-", "", "None", "nan", "N/A"]) | (~(fail_mask | pass_mask | warn_mask))

        f_cnt, p_cnt, w_cnt, e_cnt = int(fail_mask.sum()), int(pass_mask.sum()), int(warn_mask.sum()), int(empty_mask.sum())
        lines.append("\n4. Actual Result(실제 판정 및 배경색) 현황:")
        lines.append(f"  - 정상(PASS/초록): {p_cnt}건 ({(p_cnt/total_cnt)*100:.1f}%)")
        lines.append(f"  - 불합격/주의(FAIL/빨강 배경): {f_cnt}건 ({(f_cnt/total_cnt)*100:.1f}%)")
        lines.append(f"  - 확인필요/보류(노랑 배경): {w_cnt}건 ({(w_cnt/total_cnt)*100:.1f}%)")
        lines.append(f"  - 미수행/결과 미기입(공란): {e_cnt}건 ({(e_cnt/total_cnt)*100:.1f}%)")
        if f_cnt > 0:
            lines.append(f"  - [주의/불합격 대표 항목]: {_get_rep_items(df[fail_mask], 3)}")

    # 5) 수행 방식(수동 절차 vs 자동화) 안내
    if auto_col:
        lines.append("\n5. 수행 방식(자동/수동) 분포:")
        vc = df[auto_col].fillna("미지정").astype(str).value_counts()
        for val, cnt in vc.items():
            lines.append(f"  - {val}: {cnt}건 ({(cnt/total_cnt)*100:.1f}%)")
    else:
        lines.append(
            "\n5. 테스트 수행 성격(중요 가이드):\n"
            "  - 본 문서는 Precondition, Procedure(수행 절차), Pass/Fail Criteria, Actual Result로 구성된 'TC 검증 절차 레포트'입니다.\n"
            "  - TC ID에 포함된 'Auto_RTC'의 'Auto'는 차량용(Automotive) 프로젝트/식별자 접두사일 뿐 자동화 스크립트 여부를 뜻하지 않습니다.\n"
            "  - 별도의 자동화 전용 컬럼이 없으므로 절대 '수동 테스트가 0건(0%)이다'라고 단정하지 말고, "
            "'표준 절차(Procedure) 기반 기능 검증 및 결과 기록 시트' 관점에서 분석하십시오."
        )

    return "\n".join(lines)


def _sanitize_report_format(text: str) -> str:
    """
    LLM 출력 후처리:
    1) <br>, <br/> 태그가 화면에 노출되지 않도록 제거/공백 치환
    2) 건수와 비중이 '525%', '00%'처럼 붙어버린 패턴 및 마크다운 표 깨짐 교정
    3) '수동 테스트는 없습니다' 오판 문구 자동 교정
    """
    if not text:
        return ""

    # 1. 노출된 HTML 줄바꿈 태그 정리
    cleaned = re.sub(r"<br\s*/?>\s*-\s*", ", ", text, flags=re.IGNORECASE)
    cleaned = re.sub(r"<br\s*/?>", " / ", cleaned, flags=re.IGNORECASE)

    # 2. 깨진 표 머리글이 한 줄로 붙은 현상 제거
    cleaned = re.sub(
        r"-\s*구분\s*/\s*관리\s*영역\s*건수\s*및\s*비중\(%\).*?핵심\s*관리\s*포인트\s*",
        "- ",
        cleaned,
    )

    # 3. 'MULTIMEDIA525%' 또는 'Manual00%'처럼 영문/한글 뒤에 건수+비중이 붙은 패턴 교정
    # 예: MULTIMEDIA525% -> MULTIMEDIA : 5건 (25%)
    cleaned = re.sub(r"([A-Za-z가-힣]+)(\d+)\s*(100|\d{1,2})%", r"\1: \2건 (\3%) - ", cleaned)

    # 4. '수동 테스트는 없습니다' 오판 문구 교정
    cleaned = re.sub(
        r"-\s*Manual\s*:?\s*0건?\s*\(?0%?\)?.*?수동\s*테스트는\s*없습니다\.?",
        "- 수행 방식 특성: 본 시트는 상세 수행 절차(Procedure)와 판정 기준(Pass/Fail Criteria)이 정의된 절차 기반 TC 레포트이며, ID의 'Auto_RTC'는 프로젝트 식별자입니다.",
        cleaned,
    )
    cleaned = cleaned.replace(
        "수동 테스트는 없습니다.",
        "본 데이터는 Procedure(수행 절차) 및 판정 기준이 명시된 절차 기반 검증 시트입니다.",
    )

    # 5. 빈 불릿(•, -) 정리
    lines = []
    for line in cleaned.splitlines():
        stripped = line.strip()
        if stripped in ["•", "-", "*", "▶", "• ▶"]:
            continue
        lines.append(line)

    return "\n".join(lines).strip()


def _ensure_complete_report(report_text: str, dataset_type: str = "tc") -> str:
    """끊김 감지 및 4번 실무 조치 제안 완결성 보장 안전장치"""
    cleaned = _sanitize_report_format(report_text)
    if not cleaned:
        return "⚠️ 리포트 생성 결과가 비어 있습니다."

    # 마지막 줄이 중간에 잘렸는지 확인 후 보정
    if not cleaned.endswith((".", "다", "요", "함", "음", ")", "]")):
        last_newline = cleaned.rfind("\n")
        if last_newline > 0 and len(cleaned) - last_newline < 45:
            cleaned = cleaned[:last_newline].strip()

    if "4." not in cleaned and "조치 제안" not in cleaned:
        if dataset_type == "tc":
            cleaned += (
                "\n\n4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
                "- [즉시 조치 (24시간 내 / QA·검증팀)]: Actual Result 미기입 항목 및 주의(빨강/노랑 배경) 표시된 TC의 재수행 로그를 확보하고 판정 사유를 기입하십시오.\n"
                "- [단기 대책 (1주 내 / 코덱·드라이버 개발팀)]: 다수 비중을 차지하는 Video Encode/Decode(VP9, VP8, H.264, H.265) 및 Multi-Instance 시나리오의 경계값·예외 스트림 테스트를 보강하십시오.\n"
                "- [중장기 대책 (1개월 내 / 품질관리 파트)]: 엑셀 내 수동/자동 수행 구분(Execution Type) 컬럼을 표준화하여 자동화 커버리지와 수동 절차 검증 범위를 명확히 분리 관리하십시오."
            )
        else:
            cleaned += (
                "\n\n4. 실무 조치 제안 (시점별 로드맵 및 R&R)\n"
                "- [즉시 조치 (24시간 내 / 현장관리자)]: 이상치 발생 구간의 설비 파라미터 및 작업 이력을 즉시 점검하십시오.\n"
                "- [근본 대책 (1개월 내 / 품질·기술팀)]: 반복 불량 유형에 대한 표준 작업서(SOP) 개정 및 예방 점검 주기를 단축하십시오."
            )
    return cleaned


def reporter_node(state: dict) -> dict:
    """LangGraph 리포터 노드: 데이터 유형 맞춤형 프롬프트 생성 및 LLM 호출"""
    df = state.get("df")
    summary = state.get("summary", "")
    question = state.get("question", "현재 데이터의 핵심 분포와 이상 징후, 실무 조치 제안을 정리해 줘.")
    model_mode = state.get("model_mode", "local_ollama")

    dataset_type = _detect_dataset_type(df, summary)
    is_fast = "fast" in str(model_mode).lower()

    if dataset_type == "tc":
        tc_context = _build_tc_structured_context(df) if df is not None else summary
        prompt = f"""당신은 수석 QA/검증 데이터 분석 전문가입니다.
아래 [파이썬 정밀 집계 요약]의 수치만 그대로 사용하여 사용자의 질문에 대한 분석 리포트를 작성하십시오.

[중요 규칙]
1. 파이썬이 계산한 건수와 비중 형식(예: '5건 (25.0%)')을 그대로 표기하고, 절대 숫자를 붙여 쓰지 마십시오(예: 525% 금지).
2. 마크다운 표(|)나 <br> 태그를 절대 사용하지 말고, 깔끔한 불릿 리스트(- ) 형식으로 작성하십시오.
3. TC ID의 'Auto_RTC'는 차량용(Automotive) 식별자일 뿐이므로, 절대 "수동 테스트가 없다(Manual 0%)"고 쓰지 마십시오.
4. 각 항목의 대표 TC는 파이썬 집계에 나온 서로 다른 ID·Title을 인용하고 똑같은 목록을 반복 복사하지 마십시오.

[파이썬 정밀 집계 요약]
{tc_context}

[추가 기초 요약]
{summary}

[사용자 분석 질문]
{question}

[필수 출력 목차 (아래 4개 번호와 제목을 정확히 유지)]
1. 전체 검증 데이터 요약 및 수행·판정 현황
- 전체 조회 건수, 주요 검증 대상(도메인/모듈), Actual Result 판정 분포(정상/불합격·빨강배경/미기입)를 수치를 포함해 2~3줄로 요약하십시오.

2. 도메인 / 서브시스템 & 주요 기능 시나리오별 심층 분포 분석
- 도메인별 분포: [영역명] — [건수 (비중%)] | 대표 항목: [ID·Title] | 핵심 검증 포인트
- 서브시스템(App/Module)별 분포: [모듈명] — [건수 (비중%)] | 대표 항목: [ID·Title] | 핵심 검증 포인트
- 기능 시나리오 유형별 분포: 인코딩(VP9/VP8/H.265/H.264), 디코딩, 다중 인스턴스(Multiple Instances) 등 실제 시나리오별 건수(비중%)와 대표 ID·Title 정리
- 검증 집중도 및 절차서 구성의 실무적 의미: 특정 코덱/드라이버(MFC DD 등) 집중도와 절차(Procedure) 기반 검증 시트로서의 관리 포인트 설명

3. 품질 리스크 및 취약 구간 진단 (검증 커버리지·판정 결과 관점)
- 결과 판정 리스크: Actual Result 내 불합격/주의(빨강·노랑 배경셀) 항목 또는 결과 미기입(공란) 항목에 대한 점검 포인트
- 시나리오 편중 및 예외 검증 관점: 정상 기능(Functionality) 위주 테스트 외에 예외/스트레스/경계값 테스트 보완 필요성
- 테스트 자산 관리 관점: 수동 절차 수행과 자동화 스크립트 대상 구분을 위한 메타데이터(실행 방식 열 등) 관리 제안

4. 실무 조치 제안 (시점별 로드맵 및 R&R)
- 즉시 조치 (24시간 내 / 검증·QA 엔지니어): 주의/불합격 셀 및 미기입 TC 확인 조치
- 근본 대책 (1개월 내 / 개발·품질 파트): 코덱/서브시스템 커버리지 확장 및 TC 템플릿 표준화 방안
"""
    else:
        style_guide = "각 항목을 1~2문장 핵심 수치 개조식으로 간결히 작성하십시오." if is_fast else "구체적 수치와 근거를 들어 상세히 작성하십시오."
        prompt = f"""당신은 사내 데이터 분석 AI 에이전트입니다.
아래 집계 수치만 사용하여 {style_guide} 마크다운 표(|)나 <br> 태그 없이 불릿(-)으로 작성하십시오.

[데이터 집계 요약]
{summary}

[사용자 질문]
{question}

[필수 출력 목차]
1. 핵심 추세 및 전체 요약
2. 주요 이상 항목 및 집중 구간 분석 (건수 및 비중% 명시)
3. 원인 가설 (데이터 성격에 맞는 3~4대 관점 추론)
4. 실무 조치 제안 (즉시 조치 24h / 근본 대책 1개월 R&R)
"""

    raw_report = invoke_llm(prompt, mode=model_mode)
    final_report = _ensure_complete_report(raw_report, dataset_type=dataset_type)
    return {"report": final_report}