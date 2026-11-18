# data_agent/excel_exporter.py
import io
import math
import re
import zipfile
from html import escape
from datetime import datetime
from collections import Counter
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def _safe_str(val) -> str:
    if val is None:
        return "-"
    try:
        if pd.isna(val):
            return "-"
    except Exception:
        pass
    s = str(val).strip()
    return "-" if s.lower() in ("", "none", "nan", "null", "<na>") else s


def _clean_tag_str(val) -> str:
    s = _safe_str(val)
    if s == "-":
        return "-"
    cleaned = re.sub(r"\s*\[[^\]]+\]", "", s).strip()
    return cleaned if cleaned else "-"


def shorten_long_label(text, max_len: int = 28) -> str:
    s = _clean_tag_str(text)
    if len(s) <= max_len:
        return s
    return s[: max_len - 2] + ".."


def evaluate_threshold_verdict(df: pd.DataFrame, threshold: float = 1.5) -> str:
    """사내 관리 기준값(threshold%)을 현재 데이터에 적용하여 자동 판정 문구를 생성합니다."""
    if df is None or len(df) == 0:
        return f"✅ [기준({threshold}%) 판정] 조회된 데이터가 없습니다."

    all_cols = df.columns.tolist()
    total_cnt = len(df)

    if "생산량" in all_cols and "불량수" in all_cols:
        line_col = "라인" if "라인" in all_cols else ("생산라인" if "생산라인" in all_cols else None)
        exceed_lines = []
        if line_col:
            for l_name, grp in df.groupby(line_col, sort=False):
                p_sum = int(grp["생산량"].sum())
                d_sum = int(grp["불량수"].sum())
                rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
                if rate > threshold:
                    exceed_lines.append(f"{l_name}({rate:.2f}%)")
        total_p = int(df["생산량"].sum())
        total_d = int(df["불량수"].sum())
        avg_rate = (total_d / total_p * 100) if total_p > 0 else 0.0

        if exceed_lines:
            return (
                f"🚨 [확인 필요 — 사내 기준({threshold}%) 초과] "
                f"기준({threshold}%) 초과 라인: {', '.join(exceed_lines)} (전체 평균 {avg_rate:.2f}%)"
            )
        return (
            f"✅ [이상 없음 — 사내 기준({threshold}%) 이내] "
            f"모든 라인이 허용 기준({threshold}%) 이내입니다 (전체 평균 {avg_rate:.2f}%)"
        )

    exceed_issues = []
    focus_groups = []

    for col in all_cols:
        if str(col).lower() in ("_polarion", "id", "tc id", "tc_id", "번호", "시트명", "파일명"):
            continue
        vals = [_clean_tag_str(v) for v in df[col].tolist()]
        empty_cnt = sum(1 for v in vals if v in ("-", "", "(빈칸)"))
        na_or_notrun_cnt = sum(1 for v in vals if v.lower() in ("n/a", "not run", "미실행", "보류", "fail", "불합격"))

        empty_pct = (empty_cnt / total_cnt * 100) if total_cnt > 0 else 0.0
        na_pct = (na_or_notrun_cnt / total_cnt * 100) if total_cnt > 0 else 0.0

        if empty_pct > threshold:
            exceed_issues.append(f"[{col}] 미기입 {empty_cnt}건({empty_pct:.1f}%)")
        if na_pct > threshold:
            exceed_issues.append(f"[{col}] 미실행/N/A/주의 {na_or_notrun_cnt}건({na_pct:.1f}%)")

        valid_vals = [v for v in vals if v not in ("-", "", "N/A", "n/a")]
        if 2 <= len(set(valid_vals)) <= 15 and len(focus_groups) < 3:
            vc = pd.Series(valid_vals).value_counts()
            for k, cnt in vc.items():
                pct = cnt / total_cnt * 100
                if pct >= threshold:
                    focus_groups.append(f"{col}:{shorten_long_label(k, 14)}({pct:.1f}%)")

    if exceed_issues:
        top_issues = ", ".join(exceed_issues[:3])
        return (
            f"🚨 [확인 필요 — 사내 기준({threshold}%) 초과] "
            f"기준({threshold}%) 초과 항목 감지: {top_issues}"
        )
    if focus_groups:
        return (
            f"✅ [이상 없음 — 사내 기준({threshold}%) 이내] "
            f"미기입/주의 비율이 기준({threshold}%) 이내이며, 기준 이상 주요 분포는 {', '.join(focus_groups[:3])}입니다."
        )
    return f"✅ [이상 없음 — 사내 기준({threshold}%) 이내] 전 항목이 설정 기준({threshold}%) 이내로 양호합니다."


def cluster_text_series_by_topic(series: pd.Series, max_groups: int = 10) -> pd.DataFrame:
    if series is None or len(series) == 0:
        return pd.DataFrame()

    stop_words = {
        "test", "case", "tc", "verify", "check", "the", "and", "for", "with", "from", "that", "when", "what",
        "functionality", "function", "stream", "video", "auto", "rtc", "mfc", "dd", "테스트", "확인", "검증", "기능", "시나리오",
    }
    cleaned_vals = [_clean_tag_str(v) for v in series.tolist() if _clean_tag_str(v) != "-"]
    total = len(cleaned_vals)
    if total == 0:
        return pd.DataFrame()

    if len(set(cleaned_vals)) <= max_groups:
        vc = pd.Series(cleaned_vals).value_counts()
        rows = []
        for k, cnt in vc.items():
            rows.append({
                "주제/분류 그룹": shorten_long_label(k, 26),
                "건수": int(cnt),
                "비중(%)": f"{(cnt / total * 100):.1f}%",
                "대표 항목 예시": shorten_long_label(k, 50),
            })
        return pd.DataFrame(rows)

    token_counter = Counter()
    row_tokens = []
    for text in cleaned_vals:
        tokens = re.findall(r"[A-Za-z0-9_.+-]{2,}|[가-힣]{2,}", text)
        filtered = [t for t in tokens if t.lower() not in stop_words and not t.isdigit()]
        row_tokens.append((text, filtered))
        for t in set(filtered):
            token_counter[t] += 1

    top_keywords = [k for k, _ in token_counter.most_common(max_groups - 1)]
    groups: dict[str, list[str]] = {kw: [] for kw in top_keywords}
    groups["기타 시나리오"] = []

    for text, filtered in row_tokens:
        assigned = None
        for kw in top_keywords:
            if kw in filtered:
                assigned = kw
                break
        if assigned:
            groups[assigned].append(text)
        else:
            groups["기타 시나리오"].append(text)

    rows = []
    for kw, items in groups.items():
        if not items:
            continue
        cnt = len(items)
        label = f"[{kw}] 관련 시나리오" if kw != "기타 시나리오" else "기타 일반 시나리오"
        rows.append({
            "주제/분류 그룹": label,
            "건수": int(cnt),
            "비중(%)": f"{(cnt / total * 100):.1f}%",
            "대표 항목 예시": shorten_long_label(items[0], 55),
        })

    res_df = pd.DataFrame(rows)
    if not res_df.empty:
        res_df = res_df.sort_values(by="건수", ascending=False).reset_index(drop=True)
    return res_df


def analyze_color_cell_contexts(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or len(df) == 0:
        return pd.DataFrame()

    rows = []
    id_col = next((c for c in df.columns if str(c).upper() in ("ID", "TC ID", "TC_ID", "번호", "라인", "설비번호")), df.columns[0])
    status_col = next(
        (c for c in df.columns if any(k in str(c).lower() for k in ["execution", "actual", "status", "result", "결과", "판정", "상태"])),
        None,
    )

    for _, row in df.iterrows():
        row_id = _clean_tag_str(row.get(id_col, "-"))
        row_status = _clean_tag_str(row.get(status_col, "-")) if status_col else "-"

        for col in df.columns:
            val_str = _safe_str(row[col])
            color_match = re.search(r"\[(🎨[^\]]+|빈칸·🎨[^\]]+)\]", val_str)
            if color_match:
                color_tag = color_match.group(1)
                clean_val = _clean_tag_str(val_str)
                if clean_val == "-":
                    clean_val = "(빈칸)"
                col_lower = str(col).lower()

                if any(k in col_lower for k in ["priority", "severity", "중요도", "심각도", "우선순위", "등급"]):
                    meaning = f"중요도/심각도({clean_val}) 강조 표시 (불합격 아님)"
                elif any(k in col_lower for k in ["technique", "category", "type", "기법", "분류", "유형"]):
                    meaning = f"테스트 기법/분류({clean_val}) 중점 항목 강조"
                elif any(k in col_lower for k in ["actual", "result", "status", "판정", "결과"]):
                    if any(w in clean_val.lower() for w in ["fail", "ng", "불합격", "오류", "에러"]):
                        meaning = "실행 결과 불합격(FAIL) 강조"
                    elif "not run" in clean_val.lower() or "미실행" in clean_val.lower() or clean_val == "(빈칸)":
                        meaning = "미실행(Not Run) 또는 결과 입력 대기칸 강조"
                    else:
                        meaning = f"실행 결과({clean_val}) 상태 강조"
                else:
                    meaning = f"[{col}] 항목 시각적 강조 표시"

                rows.append({
                    "식별 ID / 구분": row_id,
                    "색상 표시된 컬럼(열)": str(col),
                    "셀 실제 내용": clean_val,
                    "감지된 색상": color_tag,
                    "현재 실행/판정 상태": row_status,
                    "자동 문맥 해석 (실제 의미)": meaning,
                })

    return pd.DataFrame(rows)


def summarize_aux_sheets(aux_sheets: dict[str, pd.DataFrame] | None) -> pd.DataFrame:
    if not aux_sheets:
        return pd.DataFrame()

    summary_rows = []
    for s_name, s_df in aux_sheets.items():
        if s_df is None or len(s_df) == 0:
            continue
        cols = list(s_df.columns)
        for _, r in s_df.head(25).iterrows():
            non_empty_vals = []
            for c in cols:
                cleaned = _clean_tag_str(r[c])
                if cleaned != "-":
                    non_empty_vals.append(cleaned)
            if not non_empty_vals:
                continue
            key_str = non_empty_vals[0]
            val_str = " | ".join(non_empty_vals[1:]) if len(non_empty_vals[1:]) >= 1 else "(단일 설명 항목)"
            summary_rows.append({
                "출처 시트명": s_name,
                "구분 / 지표명": key_str[:50],
                "상세 내용 및 설정값": val_str[:120],
            })
    return pd.DataFrame(summary_rows)


def build_korean_stat_tables(
    df: pd.DataFrame,
    aux_sheets: dict[str, pd.DataFrame] | None = None,
    threshold: float = 1.5,
) -> dict[str, pd.DataFrame]:
    """
    [핵심 업그레이드]
    전체 행(63건 등 전체 데이터)을 전수 스캔하여 각 컬럼의 미기입(공란) 건수뿐만 아니라
    '미기입된 실제 TC ID (및 제목) 목록'을 표에 직접 기록합니다.
    이를 통해 AI 리포트와 채팅방이 '샘플 외에 있어서 모른다'고 답하는 현상을 100% 원천 차단합니다!
    """
    tables = {}
    if df is None or len(df) == 0:
        return tables

    all_cols = df.columns.tolist()
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    total_cnt = len(df)

    # 식별자(ID) 컬럼 및 제목(Title) 컬럼 자동 탐색
    id_col = next(
        (c for c in all_cols if str(c).strip().lower() in ("id", "tc id", "tc_id", "test id", "_polarion", "번호", "라인", "일자")),
        all_cols[0],
    )
    title_col = next(
        (c for c in all_cols if any(k in str(c).lower() for k in ["title", "test case", "name", "제목", "시나리오", "purpose"])),
        None,
    )

    id_list = [_clean_tag_str(v) for v in df[id_col].tolist()]
    title_list = [_clean_tag_str(v) for v in df[title_col].tolist()] if title_col else [""] * total_cnt

    aux_df = summarize_aux_sheets(aux_sheets)
    if len(aux_df) > 0:
        tables["1. 보조 시트(Summary · Environment 등) 핵심 내용 통합 요약"] = aux_df

    col_check_rows = []
    for col in all_cols:
        clean_list = [_clean_tag_str(v) for v in df[col].tolist()]
        empty_cnt = 0
        na_cnt = 0
        empty_id_items = []
        na_id_items = []

        for idx, v in enumerate(clean_list):
            row_label = f"{id_list[idx]}({shorten_long_label(title_list[idx], 16)})" if title_col and title_list[idx] != "-" else id_list[idx]
            if v in ("-", "", "(빈칸)"):
                empty_cnt += 1
                if len(empty_id_items) < 20:
                    empty_id_items.append(row_label)
            elif v.lower() in ("n/a", "not run", "미실행"):
                na_cnt += 1
                if len(na_id_items) < 12:
                    na_id_items.append(id_list[idx])

        filled_cnt = total_cnt - empty_cnt
        empty_rate = (empty_cnt / total_cnt * 100) if total_cnt > 0 else 0.0
        valid_vals = [v for v in clean_list if v not in ("-", "")]
        top_val = pd.Series(valid_vals).mode().iloc[0] if valid_vals else "(전부 미기입)"

        if empty_rate > threshold:
            col_verdict = f"🚨 [확인 필요] 미기입률({empty_rate:.1f}%) > 기준({threshold}%)"
        else:
            col_verdict = f"✅ [양호] 기준({threshold}%) 이내"

        if empty_cnt == 0:
            missing_id_str = "없음 (전수 기재 완료)"
        elif empty_cnt <= 20:
            missing_id_str = ", ".join(empty_id_items)
        else:
            missing_id_str = f"{', '.join(empty_id_items[:15])} 외 {empty_cnt - 15}건"

        col_check_rows.append({
            "컬럼(열) 이름": str(col),
            "전체 건수": total_cnt,
            "기재 완료(건)": filled_cnt,
            "미기입/공란(건)": empty_cnt,
            "미기입률(%)": f"{empty_rate:.1f}%",
            "미기입(공란) 해당 TC ID 목록": missing_id_str,
            "N/A·미실행(건)": na_cnt,
            "최다 기재값(대표)": shorten_long_label(top_val, 24),
            f"사내 기준({threshold}%) 판정": col_verdict,
        })
    tables[f"{len(tables) + 1}. 전체 컬럼별 데이터 작성 현황 및 미기입 TC ID 전수 명단 (기준 {threshold}%)"] = pd.DataFrame(col_check_rows)

    status_cols = [
        c for c in all_cols
        if any(k in str(c).lower() for k in ["execution status", "status", "actual result", "pass/fail", "판정", "결과", "상태"])
        and "criteria" not in str(c).lower()
    ]
    if status_cols:
        st_rows = []
        for s_col in status_cols[:2]:
            clean_list = [_clean_tag_str(v) for v in df[s_col].tolist()]
            vc = pd.Series(clean_list).value_counts()
            for val_name, cnt in vc.items():
                share = (cnt / total_cnt * 100) if total_cnt > 0 else 0.0
                val_low = str(val_name).lower()
                if "pass" in val_low or "합격" in val_low or "정상" in val_low:
                    interp = "✅ 정상 통과 (PASS)"
                elif "fail" in val_low or "ng" in val_low or "불합격" in val_low:
                    interp = "🚨 실행 실패 (실제 결함/FAIL)"
                elif "not run" in val_low or "미실행" in val_low or val_name in ("-", "", "(빈칸)"):
                    interp = f"⏳ 미실행/미기입 ({'기준 ' + str(threshold) + '% 초과 점검 대상' if share > threshold else '기준 이내'})"
                else:
                    interp = "📌 기타 상태"
                st_rows.append({
                    "판정/상태 기준 열": str(s_col),
                    "상태 값": val_name if val_name != "-" else "(미기입/공란)",
                    "건수(건)": int(cnt),
                    "비중(%)": f"{share:.1f}%",
                    "실무 의미 해석": interp,
                })
        if st_rows:
            tables[f"{len(tables) + 1}. 실행 및 판정 상태 정밀 집계"] = pd.DataFrame(st_rows)

    color_ctx_df = analyze_color_cell_contexts(df)
    if len(color_ctx_df) > 0:
        tables[f"{len(tables) + 1}. 셀 배경색(강조 표시) 위치 및 문맥별 실제 의미 분석"] = color_ctx_df

    line_col = "라인" if "라인" in all_cols else ("생산라인" if "생산라인" in all_cols else None)
    if line_col and "생산량" in all_cols and "불량수" in all_cols:
        line_rows = []
        for line_name, grp in df.groupby(line_col, sort=False):
            p_sum = int(grp["생산량"].sum())
            d_sum = int(grp["불량수"].sum())
            rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
            row_rates = (grp["불량수"] / grp["생산량"] * 100).fillna(0)
            max_rate = float(row_rates.max()) if len(row_rates) > 0 else 0.0
            top_defect = "-"
            if "불량유형" in grp.columns:
                defect_grp = grp.groupby("불량유형")["불량수"].sum()
                if not defect_grp.empty:
                    top_defect = f"{defect_grp.idxmax()} ({int(defect_grp.max())}개)"

            status = f"🚨 [확인 필요] 기준({threshold}%) 초과" if rate > threshold else f"✅ [이상 없음] 기준({threshold}%) 이내"
            line_rows.append({
                "라인 구분": str(line_name),
                "집계 건수": len(grp),
                "총 생산량(개)": p_sum,
                "총 불량수(개)": d_sum,
                "평균 불량률(%)": f"{rate:.2f}%",
                "최고 불량률(%)": f"{max_rate:.2f}%",
                "주요 불량유형(최다)": top_defect,
                f"사내 기준({threshold}%) 판정": status,
            })
        if line_rows:
            tables[f"{len(tables) + 1}. 라인별 생산·품질 및 사내 기준({threshold}%) 판정 집계"] = pd.DataFrame(line_rows)

    meaningful_cats = []
    for c in all_cols:
        if c in status_cols or str(c).lower() in ("_polarion", "id", "tc id", "tc_id", "번호"):
            continue
        if any(k in str(c).lower() for k in ["파일명", "시트명", "sub system", "domain", "category", "priority", "severity", "technique", "test type", "automated", "app", "ip", "불량유형", "시간대", "부서", "분류"]):
            clean_vals = [v for v in (_clean_tag_str(x) for x in df[c].tolist()) if v != "-"]
            if len(set(clean_vals)) >= 2 or len(meaningful_cats) == 0:
                meaningful_cats.append(c)

    if not meaningful_cats:
        for c in all_cols:
            clean_vals = [v for v in (_clean_tag_str(x) for x in df[c].tolist()) if v != "-"]
            if 2 <= len(set(clean_vals)) <= 25:
                meaningful_cats.append(c)

    for c_col in meaningful_cats[:4]:
        clean_list = [_clean_tag_str(v) for v in df[c_col].tolist()]
        vc = pd.Series(clean_list).value_counts()
        c_rows = []
        for val_name, cnt in vc.head(15).items():
            share = (cnt / total_cnt * 100) if total_cnt > 0 else 0.0
            focus_status = f"🎯 [기준({threshold}%) 이상 비중]" if share >= threshold else f"⚪ 기준({threshold}%) 미만 소수"
            c_rows.append({
                f"[{c_col}] 실제 항목명": val_name if val_name != "-" else "(미기입/공란)",
                "건수(건)": int(cnt),
                "비중(%)": f"{share:.1f}%",
                f"사내 기준({threshold}%) 비교": focus_status,
            })
        if c_rows:
            tables[f"{len(tables) + 1}. [{c_col}] 실제 항목별 분포 및 기준({threshold}%) 비교"] = pd.DataFrame(c_rows)

    if title_col:
        topic_df = cluster_text_series_by_topic(df[title_col], max_groups=8)
        if not topic_df.empty:
            tables[f"{len(tables) + 1}. [{title_col}] 유사 시나리오·키워드 그룹핑 집계"] = topic_df

    if num_cols:
        stat_rows = []
        for col in num_cols:
            s = df[col].dropna()
            if len(s) == 0:
                continue
            stat_rows.append({
                "수치 항목명": str(col),
                "데이터 건수": int(len(s)),
                "합계": round(float(s.sum()), 2),
                "평균": round(float(s.mean()), 2),
                "최솟값": round(float(s.min()), 2),
                "중앙값(50%)": round(float(s.median()), 2),
                "최댓값": round(float(s.max()), 2),
            })
        if stat_rows:
            tables[f"{len(tables) + 1}. 수치 항목별 기초 통계 요약"] = pd.DataFrame(stat_rows)

    return tables


def _format_section_text_for_excel(text: str) -> str:
    if not text or not str(text).strip():
        return ""
    cleaned = str(text).replace("**", "").replace("\r\n", "\n").replace("\r", "\n")
    cleaned = re.sub(
        r"(?:^|\n|\s+)(?:[-*•]\s*)?(즉시\s*조치|근본\s*대책|중장기\s*대책|기대\s*효과|설비\(Machine\)|작업자/교대\(Man\)|자재\(Material\)|공정조건\(Method\))",
        r"\n▶ \1",
        cleaned,
    )
    cleaned = re.sub(r"(?:^|\n|\s+)(?:[-*•]\s*)?(\[(?:생산|설비|품질|정량|개발|검증|QA|담당)[^\]]*\])", r"\n  • \1", cleaned)

    formatted_lines = []
    for raw_line in cleaned.splitlines():
        line = raw_line.strip()
        if not line or re.match(r"^[-*•▪▶\s]+$", line):
            continue
        if re.match(r"^\|?[\s\-:|]+\|?$", line):
            continue
        if line.startswith("|") and line.endswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|") if c.strip()]
            if cells:
                line = " ▪ " + " | ".join(cells)
            else:
                continue
        elif line.startswith("- ") or line.startswith("* "):
            body = line[2:].strip()
            if not body:
                continue
            line = "• " + body
        formatted_lines.append(line)

    return "\n".join(formatted_lines).strip()


def parse_report_into_sections(report_text: str) -> list[dict[str, str]]:
    if not report_text or not str(report_text).strip():
        return [{"구분 항목": "분석 리포트", "상세 분석 내용": "내용 없음"}]

    cleaned_lines = []
    model_info = ""
    status_banner = ""

    for raw_line in str(report_text).splitlines():
        line = raw_line.strip()
        if not line or re.match(r"^[-*•▪▶\s]+$", line):
            continue
        if line.startswith("*(연결 모델:") or line.startswith("*(로컬 모델:"):
            model_info = line.strip("*() ")
            continue
        if line.startswith(">") and ("확인 필요" in line or "이상 없음" in line or "사내 기준" in line):
            status_banner = line.lstrip("> ").replace("**", "").strip()
            continue
        cleaned_lines.append(line)

    sections = []
    if model_info:
        sections.append({"구분 항목": "구동 AI 엔진 상세", "상세 분석 내용": model_info})
    if status_banner:
        sections.append({"구분 항목": "자동 기준 판정 요약", "상세 분석 내용": status_banner})

    current_title = "핵심 분석 요약"
    current_body = []

    for line in cleaned_lines:
        header_match = re.match(r"^(?:#+\s*)?(\d+\.\s*.+)$", line)
        if header_match:
            if current_body:
                formatted_body = _format_section_text_for_excel("\n".join(current_body))
                if formatted_body:
                    sections.append({"구분 항목": current_title, "상세 분석 내용": formatted_body})
                current_body = []
            current_title = header_match.group(1).replace("**", "").strip()
        elif line.startswith("### "):
            if current_body:
                formatted_body = _format_section_text_for_excel("\n".join(current_body))
                if formatted_body:
                    sections.append({"구분 항목": current_title, "상세 분석 내용": formatted_body})
                current_body = []
            current_title = line.replace("### ", "").replace("**", "").strip()
        else:
            current_body.append(line)

    if current_body:
        formatted_body = _format_section_text_for_excel("\n".join(current_body))
        if formatted_body:
            sections.append({"구분 항목": current_title, "상세 분석 내용": formatted_body})

    return sections


def _style_worksheet(ws):
    header_fill = PatternFill(start_color="059669", end_color="059669", fill_type="solid")
    sub_title_fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
    key_col_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
    alert_fill = PatternFill(start_color="FEF2F2", end_color="FEF2F2", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    title_font = Font(name="맑은 고딕", size=11, bold=True, color="047857")
    bold_body_font = Font(name="맑은 고딕", size=10, bold=True, color="111827")
    body_font = Font(name="맑은 고딕", size=10, color="1F2937")
    thin_border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )

    col_widths = {}
    for col in ws.columns:
        max_len = 0
        col_idx = col[0].column
        col_letter = get_column_letter(col_idx)
        for cell in col:
            val = str(cell.value or "")
            for line in val.splitlines():
                char_len = sum(2 if ord(c) > 127 else 1 for c in line)
                if char_len > max_len:
                    max_len = char_len
        width = max(min(max_len + 5, 65), 16)
        ws.column_dimensions[col_letter].width = width
        col_widths[col_idx] = width

    for row in ws.iter_rows():
        max_visual_lines = 1
        for cell in row:
            if cell.value is not None and str(cell.value).strip() != "":
                val_str = str(cell.value)
                c_width = max(col_widths.get(cell.column, 30) - 4, 20)

                visual_lines = 0
                for line in val_str.splitlines():
                    line_char_len = sum(2 if ord(c) > 127 else 1 for c in line)
                    visual_lines += max(1, math.ceil(line_char_len / c_width))
                if visual_lines > max_visual_lines:
                    max_visual_lines = visual_lines

                if val_str.startswith("■"):
                    cell.font = title_font
                    cell.fill = sub_title_fill
                    cell.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
                elif cell.row == 1 or (cell.offset(row=-1).value and str(cell.offset(row=-1).value).startswith("■")):
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(vertical="center", horizontal="center", wrap_text=True)
                    cell.border = thin_border
                else:
                    if cell.column == 1:
                        cell.font = bold_body_font
                        cell.fill = key_col_fill
                        cell.alignment = Alignment(vertical="center", horizontal="left", wrap_text=True)
                    else:
                        cell.font = body_font
                        cell.alignment = Alignment(vertical="top", horizontal="left", wrap_text=True)
                        if "빨강" in val_str or "FAIL" in val_str or "확인 필요" in val_str:
                            cell.fill = alert_fill
                    cell.border = thin_border

        if max_visual_lines > 1:
            ws.row_dimensions[row[0].row].height = min(max_visual_lines * 19 + 12, 380)
        else:
            ws.row_dimensions[row[0].row].height = 26


def _write_stat_tables_to_sheet(
    writer,
    df: pd.DataFrame,
    aux_sheets: dict[str, pd.DataFrame] | None = None,
    sheet_name: str = "2_시트통합요약_및_핵심집계",
    threshold: float = 1.5,
):
    tables = build_korean_stat_tables(df, aux_sheets=aux_sheets, threshold=threshold)
    ws = writer.book.create_sheet(title=sheet_name)

    current_row = 1
    for title, t_df in tables.items():
        ws.cell(row=current_row, column=1, value=f"■ {title}")
        current_row += 1

        for col_idx, col_name in enumerate(t_df.columns, start=1):
            ws.cell(row=current_row, column=col_idx, value=col_name)
        current_row += 1

        for _, row_data in t_df.iterrows():
            for col_idx, val in enumerate(row_data, start=1):
                ws.cell(row=current_row, column=col_idx, value=val)
            current_row += 1

        current_row += 2

    _style_worksheet(ws)


def export_analysis_to_excel(
    df: pd.DataFrame,
    summary_df: pd.DataFrame,
    report_text: str,
    elapsed_time: float,
    mode: str,
    aux_sheets: dict[str, pd.DataFrame] | None = None,
    threshold: float = 1.5,
    analyzed_at: str | None = None,
) -> bytes:
    analyzed_at = analyzed_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    verdict_str = evaluate_threshold_verdict(df, threshold=threshold)

    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        parsed_secs = parse_report_into_sections(report_text)
        filtered_secs = [s for s in parsed_secs if s["구분 항목"] != "자동 기준 판정 요약"]

        report_rows = [
            {"구분 항목": "분석 일시", "상세 분석 내용": analyzed_at},
            {"구분 항목": "분석 AI 엔진", "상세 분석 내용": mode},
            {"구분 항목": "총 소요 시간", "상세 분석 내용": f"{elapsed_time}초"},
            {"구분 항목": "분석 대상 건수 및 기준값", "상세 분석 내용": f"총 {len(df):,}건 (적용 사내 기준: {threshold}%)"},
            {"구분 항목": f"자동 기준({threshold}%) 판정 요약", "상세 분석 내용": verdict_str},
        ]
        report_rows.extend(filtered_secs)

        pd.DataFrame(report_rows).to_excel(writer, index=False, sheet_name="1_AI분석리포트")
        _style_worksheet(writer.sheets["1_AI분석리포트"])

        _write_stat_tables_to_sheet(
            writer, df, aux_sheets=aux_sheets, sheet_name="2_시트통합요약_및_핵심집계", threshold=threshold
        )

        df.to_excel(writer, index=False, sheet_name="3_핵심원본데이터")
        _style_worksheet(writer.sheets["3_핵심원본데이터"])

    return output.getvalue()


def export_combined_analysis_to_excel(
    df: pd.DataFrame,
    saved_results: dict,
    model_labels: dict[str, str],
    aux_sheets: dict[str, pd.DataFrame] | None = None,
    threshold: float = 1.5,
) -> bytes:
    verdict_str = evaluate_threshold_verdict(df, threshold=threshold)
    output = io.BytesIO()

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        comp_rows = []
        for m_code, res in saved_results.items():
            analyzed_at = res.get("analyzed_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            sections = parse_report_into_sections(res.get("report", ""))
            sec_map = {s["구분 항목"]: s["상세 분석 내용"] for s in sections}

            trend_val = next((v for k, v in sec_map.items() if "1." in k or "추세" in k or "현황" in k or "요약" in k), "")
            anom_val = next((v for k, v in sec_map.items() if "2." in k or "이상" in k or "주의" in k or "분포" in k), "")
            cause_val = next((v for k, v in sec_map.items() if "3." in k or "원인" in k or "가설" in k or "리스크" in k), "")
            action_val = next((v for k, v in sec_map.items() if "4." in k or "조치" in k), "")

            comp_rows.append({
                "분석 일시": analyzed_at,
                "AI 모델명": model_labels.get(m_code, m_code),
                "소요 시간(초)": float(res.get("elapsed_time", 0.0)),
                f"자동 판정 요약 (기준 {threshold}%)": verdict_str,
                "1. 종합 현황 및 요약": _format_section_text_for_excel(trend_val) or "개별 시트 참조",
                "2. 실제 항목별 심층 분포": _format_section_text_for_excel(anom_val) or "개별 시트 참조",
                "3. 품질 리스크 및 원인 가설": _format_section_text_for_excel(cause_val) or "개별 시트 참조",
                "4. 실무 조치 제안": _format_section_text_for_excel(action_val) or "개별 시트 참조",
            })

        comp_df = pd.DataFrame(comp_rows)
        comp_df.to_excel(writer, index=False, sheet_name="1_모델별_종합비교")
        _style_worksheet(writer.sheets["1_모델별_종합비교"])

        _write_stat_tables_to_sheet(
            writer, df, aux_sheets=aux_sheets, sheet_name="2_시트통합요약_및_핵심집계", threshold=threshold
        )

        df.to_excel(writer, index=False, sheet_name="3_핵심원본데이터")
        _style_worksheet(writer.sheets["3_핵심원본데이터"])

        for m_code, res in saved_results.items():
            analyzed_at = res.get("analyzed_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            safe_sheet = f"리포트_{m_code}"[:31]
            parsed_secs = [s for s in parse_report_into_sections(res.get("report", "")) if s["구분 항목"] != "자동 기준 판정 요약"]
            m_rows = [
                {"구분 항목": "분석 일시", "상세 분석 내용": analyzed_at},
                {"구분 항목": "AI 모델명", "상세 분석 내용": model_labels.get(m_code, m_code)},
                {"구분 항목": "소요 시간", "상세 분석 내용": f"{res.get('elapsed_time', 0.0)}초"},
                {"구분 항목": f"자동 기준({threshold}%) 판정 요약", "상세 분석 내용": verdict_str},
            ]
            m_rows.extend(parsed_secs)
            pd.DataFrame(m_rows).to_excel(writer, index=False, sheet_name=safe_sheet)
            _style_worksheet(writer.sheets[safe_sheet])

    return output.getvalue()


def _docx_p(
    text: str,
    bold: bool = False,
    size_half_pt: int = 20,
    color: str = "1F2937",
    space_before: int = 0,
    space_after: int = 90,
    left_indent: int = 0,
    align: str = "left",
    page_break_before: bool = False,
    keep_with_next: bool = False,
    bottom_border_color: str | None = None,
) -> str:
    t_esc = escape(str(text))
    ppr_parts = []
    if page_break_before:
        ppr_parts.append("<w:pageBreakBefore/>")
    if keep_with_next:
        ppr_parts.append("<w:keepNext/>")
    if bottom_border_color:
        ppr_parts.append(
            f'<w:pBdr><w:bottom w:val="single" w:sz="12" w:space="4" w:color="{bottom_border_color}"/></w:pBdr>'
        )
    ppr_parts.append(f'<w:spacing w:before="{space_before}" w:after="{space_after}" w:line="276" w:lineRule="auto"/>')
    if left_indent > 0:
        ppr_parts.append(f'<w:ind w:left="{left_indent}"/>')
    if align != "left":
        ppr_parts.append(f'<w:jc w:val="{align}"/>')

    b_tag = "<w:b/><w:bCs/>" if bold else ""
    return (
        f'<w:p><w:pPr>{"".join(ppr_parts)}</w:pPr>'
        f'<w:r><w:rPr><w:rFonts w:ascii="맑은 고딕" w:eastAsia="맑은 고딕" w:hAnsi="맑은 고딕" w:cs="맑은 고딕"/>'
        f'{b_tag}<w:color w:val="{color}"/><w:sz w:val="{size_half_pt}"/><w:szCs w:val="{size_half_pt}"/></w:rPr>'
        f'<w:t xml:space="preserve">{t_esc}</w:t></w:r></w:p>'
    )


def _docx_banner(sub_label: str, main_title: str, meta_line: str, page_break_before: bool = False) -> str:
    p1 = _docx_p(
        sub_label,
        bold=True,
        size_half_pt=19,
        color="A7F3D0",
        space_before=0,
        space_after=60,
        page_break_before=page_break_before,
        keep_with_next=True,
    )
    p2 = _docx_p(main_title, bold=True, size_half_pt=28, color="FFFFFF", space_before=0, space_after=70, keep_with_next=True)
    p3 = _docx_p(meta_line, bold=False, size_half_pt=18, color="D1FAE5", space_before=0, space_after=0)

    return (
        '<w:tbl>'
        '<w:tblPr>'
        '<w:tblW w:w="5000" w:type="pct"/>'
        '<w:jc w:val="center"/>'
        '<w:tblCellMar>'
        '<w:top w:w="160" w:type="dxa"/><w:bottom w:w="160" w:type="dxa"/>'
        '<w:left w:w="200" w:type="dxa"/><w:right w:w="200" w:type="dxa"/>'
        '</w:tblCellMar>'
        '</w:tblPr>'
        '<w:tblGrid><w:gridCol w:w="9638"/></w:tblGrid>'
        '<w:tr><w:trPr><w:cantSplit/></w:trPr>'
        '<w:tc><w:tcPr><w:tcW w:w="5000" w:type="pct"/>'
        '<w:shd w:val="clear" w:color="auto" w:fill="064E3B"/></w:tcPr>'
        f'{p1}{p2}{p3}'
        '</w:tc></w:tr></w:tbl>'
    )


def _docx_table(
    headers: list[str],
    rows: list[list[str]],
    col_widths_pct: list[int] | None = None,
    header_bg: str = "059669",
) -> str:
    num_cols = len(headers)
    if not col_widths_pct or len(col_widths_pct) != num_cols:
        even_pct = 100 // max(1, num_cols)
        col_widths_pct = [even_pct] * num_cols
        col_widths_pct[-1] = 100 - sum(col_widths_pct[:-1])

    total_dxa = 9638
    grid_cols = "".join(f'<w:gridCol w:w="{int(total_dxa * (p / 100.0))}"/>' for p in col_widths_pct)

    borders = (
        '<w:tblBorders>'
        '<w:top w:val="single" w:sz="8" w:space="0" w:color="059669"/>'
        '<w:bottom w:val="single" w:sz="8" w:space="0" w:color="059669"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="CBD5E1"/>'
        '<w:left w:val="none"/>'
        '<w:right w:val="none"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>'
        '</w:tblBorders>'
    )
    cell_margins = (
        '<w:tblCellMar>'
        '<w:top w:w="95" w:type="dxa"/>'
        '<w:bottom w:w="95" w:type="dxa"/>'
        '<w:left w:w="135" w:type="dxa"/>'
        '<w:right w:w="135" w:type="dxa"/>'
        '</w:tblCellMar>'
    )

    tbl_xml = [
        f'<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/>{borders}{cell_margins}</w:tblPr>'
        f'<w:tblGrid>{grid_cols}</w:tblGrid>'
    ]

    h_cells = []
    for idx, h in enumerate(headers):
        w_pct_50 = col_widths_pct[idx] * 50
        h_cells.append(
            f'<w:tc><w:tcPr><w:tcW w:w="{w_pct_50}" w:type="pct"/>'
            f'<w:shd w:val="clear" w:color="auto" w:fill="{header_bg}"/>'
            f'<w:vAlign w:val="center"/></w:tcPr>'
            f'{_docx_p(h, bold=True, size_half_pt=19, color="FFFFFF", space_before=0, space_after=0, align="center")}'
            f'</w:tc>'
        )
    tbl_xml.append(f'<w:tr><w:trPr><w:tblHeader/><w:cantSplit/></w:trPr>{"".join(h_cells)}</w:tr>')

    for r_idx, r in enumerate(rows):
        r_cells = []
        row_bg = "FFFFFF" if r_idx % 2 == 0 else "F8FAFC"
        for c_idx, cell_val in enumerate(r):
            w_pct_50 = col_widths_pct[c_idx] * 50
            val_str = str(cell_val or "-").strip()
            is_alert = ("🚨" in val_str) or ("확인 필요" in val_str)

            if c_idx == 0:
                bg = "F0FDF4"
                txt_color = "065F46"
                is_bold = True
            elif is_alert:
                bg = "FEF2F2"
                txt_color = "991B1B"
                is_bold = True
            else:
                bg = row_bg
                txt_color = "1F2937"
                is_bold = False

            lines = [l.strip() for l in val_str.splitlines() if l.strip()] or ["-"]
            p_tags = []
            for l_i, line_txt in enumerate(lines):
                sp_after = 55 if l_i < len(lines) - 1 else 0
                p_tags.append(
                    _docx_p(
                        line_txt,
                        bold=is_bold,
                        size_half_pt=18,
                        color=txt_color,
                        space_before=0,
                        space_after=sp_after,
                        align="center" if (c_idx == 0 and num_cols >= 4) else "left",
                    )
                )
            r_cells.append(
                f'<w:tc><w:tcPr><w:tcW w:w="{w_pct_50}" w:type="pct"/>'
                f'<w:shd w:val="clear" w:color="auto" w:fill="{bg}"/>'
                f'<w:vAlign w:val="center"/></w:tcPr>{"".join(p_tags)}</w:tc>'
            )
        tbl_xml.append(f'<w:tr><w:trPr><w:cantSplit/></w:trPr>{"".join(r_cells)}</w:tr>')

    tbl_xml.append("</w:tbl>")
    return "".join(tbl_xml)


def _pack_docx_xml(body_elements: list[str]) -> bytes:
    sect_pr = (
        '<w:sectPr>'
        '<w:pgSz w:w="11906" w:h="16838"/>'
        '<w:pgMar w:top="1080" w:right="1080" w:bottom="1080" w:left="1080" w:header="720" w:footer="720" w:gutter="0"/>'
        '</w:sectPr>'
    )
    document_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        f'<w:body>{"".join(body_elements)}{sect_pr}</w:body></w:document>'
    )
    content_types_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
        '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
        '<Default Extension="xml" ContentType="application/xml"/>'
        '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>'
        '</Types>'
    )
    rels_xml = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
        '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>'
        '</Relationships>'
    )

    buf = io.BytesIO()
    with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("[Content_Types].xml", content_types_xml)
        zf.writestr("_rels/.rels", rels_xml)
        zf.writestr("word/document.xml", document_xml)
    return buf.getvalue()


def _append_formatted_report_sections(body: list[str], report_text: str):
    parsed = [
        s for s in parse_report_into_sections(report_text)
        if s["구분 항목"] not in ("자동 기준 판정 요약", "구동 AI 엔진 상세")
    ]
    for sec in parsed:
        body.append(
            _docx_p(
                f"■ {sec['구분 항목']}",
                bold=True,
                size_half_pt=22,
                color="047857",
                space_before=160,
                space_after=70,
                keep_with_next=True,
                bottom_border_color="059669",
            )
        )
        lines = [l.strip() for l in str(sec["상세 분석 내용"]).splitlines() if l.strip()]
        for l_idx, clean_l in enumerate(lines):
            is_sub_header = clean_l.startswith("▶")
            body.append(
                _docx_p(
                    clean_l,
                    bold=is_sub_header,
                    size_half_pt=19,
                    color="065F46" if is_sub_header else "1F2937",
                    space_before=60 if is_sub_header else 0,
                    space_after=55 if l_idx < len(lines) - 1 else 100,
                    left_indent=120 if not is_sub_header else 40,
                )
            )


def export_analysis_to_docx(
    df: pd.DataFrame,
    report_text: str,
    elapsed_time: float,
    mode: str,
    threshold: float = 1.5,
    analyzed_at: str | None = None,
) -> bytes:
    analyzed_at = analyzed_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    verdict_str = evaluate_threshold_verdict(df, threshold=threshold)

    body = [
        _docx_banner(
            sub_label="[사내 데이터 분석 AI Agent] 단일 모델 상세 분석 보고서",
            main_title=f"{mode} — 데이터 정밀 분석 리포트",
            meta_line=f"분석 일시: {analyzed_at}  |  대상 데이터: 총 {len(df):,}건  |  소요 시간: {elapsed_time}초",
            page_break_before=False,
        ),
        _docx_p(
            "01. 분석 개요 및 사내 기준 자동 판정",
            bold=True,
            size_half_pt=22,
            color="047857",
            space_before=160,
            space_after=80,
            keep_with_next=True,
            bottom_border_color="059669",
        ),
        _docx_table(
            headers=["구분 항목", "상세 내용"],
            rows=[
                ["분석 일시", analyzed_at],
                ["구동 AI 엔진", f"{mode} (분석 소요 시간: {elapsed_time}초)"],
                ["분석 대상 건수", f"총 {len(df):,}건 (사내 관리 기준값: {threshold}% 적용)"],
                [f"사내 기준({threshold}%) 판정", verdict_str],
            ],
            col_widths_pct=[24, 76],
        ),
    ]

    _append_formatted_report_sections(body, report_text)
    return _pack_docx_xml(body)


def export_combined_analysis_to_docx(
    df: pd.DataFrame,
    saved_results: dict,
    model_labels: dict[str, str],
    threshold: float = 1.5,
) -> bytes:
    verdict_str = evaluate_threshold_verdict(df, threshold=threshold)
    valid_items = [
        (m_code, res)
        for m_code, res in saved_results.items()
        if res and str(res.get("report", "")).strip()
    ]
    first_time = next(
        (r.get("analyzed_at") for _, r in valid_items if r.get("analyzed_at")),
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
    )
    model_names_str = ", ".join(model_labels.get(m, m) for m, _ in valid_items)

    integrated_insights = []
    check_required_points = []
    summary_table_rows = []

    for m_code, res in valid_items:
        m_title = model_labels.get(m_code, m_code)
        short_m = m_title.split("(")[0].replace("⚡", "").replace("☁️", "").replace("🔒", "").strip()
        el_str = f"{res.get('elapsed_time', 0.0)}초"
        sections = parse_report_into_sections(res.get("report", ""))
        sec_map = {s["구분 항목"]: s["상세 분석 내용"] for s in sections}

        s1 = next((v for k, v in sec_map.items() if "1." in k or "요약" in k or "현황" in k or "추세" in k), "")
        s2 = next((v for k, v in sec_map.items() if "2." in k or "분포" in k or "이상" in k), "")
        s3 = next((v for k, v in sec_map.items() if "3." in k or "리스크" in k or "원인" in k), "")
        s4 = next((v for k, v in sec_map.items() if "4." in k or "조치" in k), "")

        s1_lines = [l.lstrip("•-*▶ ").strip() for l in str(s1).splitlines() if l.strip()]
        s2_lines = [l.lstrip("•-*▶ ").strip() for l in str(s2).splitlines() if l.strip()]
        s3_lines = [l.lstrip("•-*▶ ").strip() for l in str(s3).splitlines() if l.strip()]
        s4_lines = [l.lstrip("•-*▶ ").strip() for l in str(s4).splitlines() if l.strip()]

        if s1_lines:
            integrated_insights.append(f"• [{short_m}] {s1_lines[0][:95]}")
        if s3_lines:
            check_required_points.append(f"• [{short_m} 점검] {s3_lines[0][:95]}")
        elif s4_lines:
            check_required_points.append(f"• [{short_m} 조치] {s4_lines[0][:95]}")

        def _compact_lines(lines_list: list[str], max_l: int = 2) -> str:
            if not lines_list:
                return "개별 상세 페이지 참조"
            return "\n".join(f"• {l[:70]}" for l in lines_list[:max_l])

        summary_table_rows.append([
            f"{short_m}\n({el_str})",
            _compact_lines(s1_lines, 2),
            _compact_lines(s2_lines, 2),
            _compact_lines(s4_lines, 2),
        ])

    integrated_insight_str = "\n".join(integrated_insights[:3]) if integrated_insights else "각 모델별 상세 분석 참조"
    check_required_str = (
        f"{verdict_str}\n" + ("\n".join(check_required_points[:3]) if check_required_points else "")
    ).strip()

    body = [
        _docx_banner(
            sub_label="[사내 데이터 분석 AI Agent] 다중 AI 모델 종합 분석 보고서",
            main_title="1페이지: 전체 AI 분석 종합 요약 및 확인 필요 사항",
            meta_line=f"분석 일시: {first_time}  |  비교 모델: {len(valid_items)}종  |  대상 데이터: 총 {len(df):,}건 (기준 {threshold}%)",
            page_break_before=False,
        ),
        _docx_p(
            "01. 전체 분석 종합 요약 및 '확인 필요' 핵심 진단",
            bold=True,
            size_half_pt=22,
            color="047857",
            space_before=140,
            space_after=70,
            keep_with_next=True,
            bottom_border_color="059669",
        ),
        _docx_table(
            headers=["종합 요약 구분", "핵심 내용 및 확인 필요 사항"],
            rows=[
                ["분석 일시 / 대상 건수", f"{first_time}  |  총 {len(df):,}건 (사내 관리 기준값: {threshold}% 적용)"],
                ["참여 AI 모델", f"총 {len(valid_items)}개 모델 동시 분석 ({model_names_str})"],
                [f"확인 필요 사항 (기준 {threshold}%)", check_required_str],
                ["통합 핵심 인사이트 요약", integrated_insight_str],
            ],
            col_widths_pct=[24, 76],
        ),
        _docx_p(
            "02. API(AI) 모델별 핵심 분석 결과 비교 요약표",
            bold=True,
            size_half_pt=22,
            color="047857",
            space_before=140,
            space_after=70,
            keep_with_next=True,
            bottom_border_color="059669",
        ),
        _docx_table(
            headers=["AI 모델 (소요시간)", "1. 핵심 현황 요약", "2. 주요 분포·집중 구간", "4. 실무 조치 핵심 요약"],
            rows=summary_table_rows,
            col_widths_pct=[18, 27, 27, 28],
        ),
        _docx_p(
            "※ 2페이지부터는 각 API(AI) 모델별 상세 분석 리포트 전문이 독립된 페이지에서 순서대로 시작됩니다.",
            bold=True,
            size_half_pt=18,
            color="047857",
            space_before=100,
            space_after=0,
        ),
    ]

    for idx, (m_code, res) in enumerate(valid_items, start=1):
        m_name = model_labels.get(m_code, m_code)
        el = res.get("elapsed_time", 0.0)
        m_time = res.get("analyzed_at") or first_time

        body.append(
            _docx_banner(
                sub_label=f"[개별 AI 상세 분석 섹션 {idx}/{len(valid_items)}] — 독립 페이지 분리 리포트",
                main_title=f"{m_name} 상세 분석 보고서",
                meta_line=f"분석 일시: {m_time}  |  분석 소요 시간: {el}초  |  적용 기준값: {threshold}%",
                page_break_before=True,
            )
        )
        body.append(
            _docx_p(
                "■ 모델 실행 정보 및 사내 기준 판정",
                bold=True,
                size_half_pt=21,
                color="047857",
                space_before=130,
                space_after=65,
                keep_with_next=True,
            )
        )
        body.append(
            _docx_table(
                headers=["항목", "상세 내용"],
                rows=[
                    ["구동 AI 모델", f"{m_name} (총 {el}초 소요)"],
                    ["분석 일시 / 데이터 건수", f"{m_time}  |  총 {len(df):,}건"],
                    [f"자동 기준({threshold}%) 판정 요약", verdict_str],
                ],
                col_widths_pct=[24, 76],
            )
        )

        _append_formatted_report_sections(body, res.get("report", ""))

    return _pack_docx_xml(body)