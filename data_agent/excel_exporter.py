import io
import math
import re
from collections import Counter
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference


# 제목/문장 클러스터링 시 제외할 일반 동사·조사·형용사 불용어
TOPIC_STOPWORDS = {
    "조회", "확인", "이동", "검증", "테스트", "기능", "동작", "표시", "출력", "입력",
    "선택", "클릭", "버튼", "화면", "상태", "경우", "정상", "여부", "시나리오", "항목",
    "있음", "없음", "한다", "된다", "이다", "에서", "으로", "에게", "대한", "통해",
    "후", "전", "시", "및", "등", "수", "값", "창", "란", "키", "키로", "다른", "이전",
}


def cluster_text_series_by_topic(series: pd.Series, max_groups: int = 8) -> pd.DataFrame:
    """
    'Title'처럼 각 행이 1건씩 모두 다른 문장일 때,
    공통으로 반복 등장하는 핵심 키워드(예: '문서', '목차', '검색', '대문', '배달', '쿠폰')를 자동 추출하여
    의미 있는 주제 그룹으로 묶고 건수·비중(%)·포함 항목 예시를 반환합니다.
    """
    clean_vals = [
        re.sub(r"\s*\[[^\]]+\]", "", str(v)).strip()
        for v in series.dropna()
        if str(v).strip() not in ("-", "", "None", "nan")
    ]
    total = len(clean_vals)
    if total == 0:
        return pd.DataFrame()

    vc = pd.Series(clean_vals).value_counts()
    if len(vc) <= max_groups and vc.iloc[0] > 1:
        rows = []
        for k, cnt in vc.items():
            rows.append({
                "주제/분류 그룹": str(k),
                "건수": int(cnt),
                "비중(%)": round(cnt / total * 100, 1),
                "표시라벨": f"{cnt}건 ({cnt / total * 100:.1f}%)",
                "포함된 상세 항목": str(k),
            })
        return pd.DataFrame(rows)

    tokenized_list = []
    word_counter = Counter()
    for text in clean_vals:
        words = re.findall(r"[가-힣A-Za-z0-9]{2,}", text)
        filtered = []
        for w in words:
            w_stem = re.sub(r"(으로|에서|에게|부터|까지|이라|란|로|은|는|이|가|을|를|의|와|과|만|도)$", "", w)
            if len(w_stem) >= 2 and w_stem.lower() not in TOPIC_STOPWORDS:
                filtered.append(w_stem)
        unique_in_row = list(dict.fromkeys(filtered))
        tokenized_list.append((text, unique_in_row))
        for uw in unique_in_row:
            word_counter[uw] += 1

    # 2회 이상 등장한 핵심 키워드를 빈도순으로 선정
    repeated_keywords = [w for w, c in word_counter.most_common(15) if c >= 2]

    grouped_items: dict[str, list[str]] = {}
    for text, tokens in tokenized_list:
        assigned_kw = None
        for kw in repeated_keywords:
            if kw in tokens or kw in text:
                assigned_kw = f"📁 [{kw}] 관련 항목"
                break
        if assigned_kw is None:
            first_word = tokens[0] if tokens else text[:10]
            assigned_kw = f"📌 기타 ({first_word} 등)" if len(repeated_keywords) >= 2 else f"📄 {text[:18]}"
        grouped_items.setdefault(assigned_kw, []).append(text)

    sorted_groups = sorted(grouped_items.items(), key=lambda x: len(x[1]), reverse=True)
    final_rows = []
    etc_items = []

    for idx, (g_name, items) in enumerate(sorted_groups):
        if idx < max_groups - 1 or len(sorted_groups) <= max_groups:
            cnt = len(items)
            sample_preview = ", ".join(items[:4]) + (f" 외 {cnt - 4}건" if cnt > 4 else "")
            final_rows.append({
                "주제/분류 그룹": g_name,
                "건수": cnt,
                "비중(%)": round(cnt / total * 100, 1),
                "표시라벨": f"{cnt}건 ({cnt / total * 100:.1f}%)",
                "포함된 상세 항목": sample_preview,
            })
        else:
            etc_items.extend(items)

    if etc_items:
        cnt = len(etc_items)
        sample_preview = ", ".join(etc_items[:4]) + (f" 외 {cnt - 4}건" if cnt > 4 else "")
        final_rows.append({
            "주제/분류 그룹": "📌 기타 단건 시나리오 묶음",
            "건수": cnt,
            "비중(%)": round(cnt / total * 100, 1),
            "표시라벨": f"{cnt}건 ({cnt / total * 100:.1f}%)",
            "포함된 상세 항목": sample_preview,
        })

    return pd.DataFrame(final_rows)


def shorten_long_label(text: str, max_len: int = 24) -> str:
    """Precondition이나 Expected Result처럼 긴 문장을 차트 축에 표시할 때 핵심만 남기고 깔끔하게 축약합니다."""
    s = re.sub(r"\s*\[[^\]]+\]", "", str(text)).strip()
    s = re.sub(r"https?://[^\s/]+(?:/wiki/)?", "", s).strip()
    s = re.sub(r"^PC 브라우저에서\s*", "", s).strip()
    if len(s) <= max_len:
        return s
    return s[:max_len].rstrip() + "…"


def analyze_color_cell_contexts(df: pd.DataFrame) -> pd.DataFrame:
    """셀 배경색(🎨빨강강조 등)이 칠해진 셀의 컬럼 위치와 실제 의미를 분석합니다."""
    rows = []
    if len(df) == 0:
        return pd.DataFrame()

    id_col = next((c for c in df.columns if c.upper() in ("ID", "TC ID", "TC_ID", "TC-ID", "번호", "라인", "설비번호")), df.columns[0])
    status_col = next(
        (c for c in df.columns if any(k in c.lower() for k in ["execution", "actual", "status", "result", "결과", "판정", "상태"])),
        None,
    )

    for _, row in df.iterrows():
        row_id = str(row.get(id_col, "-")).replace("[🎨빨강강조]", "").strip()
        row_status = str(row.get(status_col, "-")).strip() if status_col else "-"

        for col in df.columns:
            val_str = str(row[col])
            color_match = re.search(r"\[(🎨[^\]]+|빈칸·🎨[^\]]+)\]", val_str)
            if color_match:
                color_tag = color_match.group(1)
                clean_val = re.sub(r"\s*\[[^\]]+\]", "", val_str).strip() or "(빈칸)"
                col_lower = col.lower()

                if any(k in col_lower for k in ["priority", "severity", "중요도", "심각도", "우선순위", "등급"]):
                    meaning = f"중요도/심각도({clean_val}) 강조 표시 (불합격 아님)"
                elif any(k in col_lower for k in ["technique", "category", "type", "기법", "분류", "유형"]):
                    meaning = f"테스트 기법/분류({clean_val}) 중점 항목 강조"
                elif any(k in col_lower for k in ["actual", "result", "status", "판정", "결과"]):
                    if any(w in clean_val.lower() for w in ["fail", "ng", "불합격", "오류", "에러"]):
                        meaning = "실행 결과 불합격(FAIL) 강조"
                    elif "not run" in clean_val.lower() or "미실행" in clean_val.lower() or clean_val == "(빈칸)":
                        meaning = "미실행(Not Run) 또는 결과 입력 대기칸 강조 (실제 실패 아님)"
                    else:
                        meaning = f"실행 결과({clean_val}) 상태 강조"
                else:
                    meaning = f"[{col}] 항목 시각적 강조 표시"

                rows.append({
                    "식별 ID / 구분": row_id,
                    "색상 표시된 컬럼(열)": col,
                    "셀 실제 내용": clean_val[:40],
                    "감지된 색상": color_tag,
                    "현재 실행/판정 상태": row_status[:30],
                    "자동 문맥 해석 (실제 의미)": meaning,
                })

    return pd.DataFrame(rows)


def summarize_aux_sheets(aux_sheets: dict[str, pd.DataFrame] | None) -> pd.DataFrame:
    """Summary, Environment 등 보조 시트의 내용을 한눈에 볼 수 있는 통합 요약표로 변환합니다."""
    if not aux_sheets:
        return pd.DataFrame()

    summary_rows = []
    for s_name, s_df in aux_sheets.items():
        if len(s_df) == 0:
            continue
        cols = list(s_df.columns)
        for _, r in s_df.head(25).iterrows():
            non_empty_vals = [
                re.sub(r"\s*\[🎨[^\]]+\]", "", str(r[c])).strip()
                for c in cols
                if str(r[c]).strip() not in ("-", "", "None", "nan") and not str(r[c]).startswith("[빈칸")
            ]
            if not non_empty_vals:
                continue
            key_str = non_empty_vals[0]
            val_str = " | ".join(non_empty_vals[1:]) if len(non_empty_vals) > 1 else "(단일 설명 항목)"
            summary_rows.append({
                "출처 시트명": s_name,
                "구분 / 지표명": key_str[:50],
                "상세 내용 및 설정값": val_str[:120],
            })
    return pd.DataFrame(summary_rows)


def build_korean_stat_tables(df: pd.DataFrame, aux_sheets: dict[str, pd.DataFrame] | None = None) -> dict[str, pd.DataFrame]:
    """어떤 구조의 파일이든 의미 있는 한글 통계표(보조시트 요약, 주제 그룹핑 집계, 사전조건 분포, 판정/색상 문맥 등)를 생성합니다."""
    tables = {}
    if len(df) == 0:
        return tables

    all_cols = df.columns.tolist()
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    total_cnt = len(df)

    aux_df = summarize_aux_sheets(aux_sheets)
    if len(aux_df) > 0:
        tables["1. 보조 시트(Summary · Environment 등) 핵심 내용 통합 요약"] = aux_df

    title_col = next((c for c in all_cols if any(k in c.lower() for k in ["title", "제목", "타이틀", "시나리오", "기능명", "항목명"])), None)
    if title_col:
        topic_df = cluster_text_series_by_topic(df[title_col], max_groups=10)
        if len(topic_df) > 0:
            tables[f"{len(tables) + 1}. [{title_col}] 유사 주제(키워드)별 자동 그룹핑 집계"] = topic_df[
                ["주제/분류 그룹", "건수", "비중(%)", "포함된 상세 항목"]
            ]

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

            status = "🚨 [확인 필요] 기준(1.5%) 초과" if rate > 1.5 else "✅ [이상 없음] 기준 이내"
            line_rows.append({
                "라인 구분": str(line_name),
                "집계 건수": len(grp),
                "총 생산량(개)": p_sum,
                "총 불량수(개)": d_sum,
                "평균 불량률(%)": f"{rate:.2f}%",
                "최고 불량률(%)": f"{max_rate:.2f}%",
                "주요 불량유형(최다)": top_defect,
                "사내 기준(1.5%) 판정": status,
            })
        if line_rows:
            tables[f"{len(tables) + 1}. 라인별 생산·품질 및 주요 불량 종합 집계"] = pd.DataFrame(line_rows)

    cat_candidates = [
        c for c in all_cols
        if any(k in c.lower() for k in ["category", "priority", "severity", "precondition", "status", "actual", "불량유형", "시간대", "분류"])
        and c != title_col
    ]
    for c_col in cat_candidates[:3]:
        clean_s = df[c_col].astype(str).apply(lambda x: re.sub(r"\s*\[[^\]]+\]", "", x).strip())
        if clean_s.nunique() > 12 and len(df) >= 10:
            grp_df = cluster_text_series_by_topic(clean_s, max_groups=8)
            if len(grp_df) > 0:
                tables[f"{len(tables) + 1}. [{c_col}] 유형별 요약 집계"] = grp_df[
                    ["주제/분류 그룹", "건수", "비중(%)", "포함된 상세 항목"]
                ]
        else:
            vc = clean_s.value_counts()
            c_rows = []
            for val_name, cnt in vc.head(12).items():
                share = (cnt / total_cnt * 100) if total_cnt > 0 else 0.0
                c_rows.append({
                    f"[{c_col}] 구분": val_name,
                    "건수(건)": int(cnt),
                    "비중(%)": f"{share:.1f}%",
                })
            if c_rows:
                tables[f"{len(tables) + 1}. [{c_col}] 항목별 분포 및 비중 집계"] = pd.DataFrame(c_rows)

    if num_cols:
        stat_rows = []
        for col in num_cols:
            s = df[col].dropna()
            if len(s) == 0:
                continue
            stat_rows.append({
                "수치 항목명": col,
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
    if not text or not text.strip():
        return ""
    cleaned = text.replace("**", "").replace("\r\n", "\n").replace("\r", "\n")
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
    if not report_text or not report_text.strip():
        return [{"구분 항목": "분석 리포트", "상세 분석 내용": "내용 없음"}]

    cleaned_lines = []
    model_info = ""
    status_banner = ""

    for raw_line in report_text.splitlines():
        line = raw_line.strip()
        if not line or re.match(r"^[-*•▪▶\s]+$", line):
            continue
        if line.startswith("*(연결 모델:") or line.startswith("*(로컬 모델:"):
            model_info = line.strip("*() ")
            continue
        if "확인 필요" in line or "이상 없음" in line or "점검 안내" in line:
            status_banner = line.lstrip("> ").replace("**", "")
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
                        if "빨강" in val_str or "FAIL" in val_str:
                            cell.fill = alert_fill
                    cell.border = thin_border

        if max_visual_lines > 1:
            ws.row_dimensions[row[0].row].height = min(max_visual_lines * 19 + 12, 380)
        else:
            ws.row_dimensions[row[0].row].height = 26


def _write_stat_tables_and_chart_to_sheet(
    writer,
    df: pd.DataFrame,
    aux_sheets: dict[str, pd.DataFrame] | None = None,
    sheet_name: str = "2_시트통합요약_및_핵심집계",
):
    tables = build_korean_stat_tables(df, aux_sheets=aux_sheets)
    ws = writer.book.create_sheet(title=sheet_name)

    current_row = 1
    chart_target_coords = None

    for title, t_df in tables.items():
        ws.cell(row=current_row, column=1, value=f"■ {title}")
        current_row += 1

        header_row = current_row
        for col_idx, col_name in enumerate(t_df.columns, start=1):
            ws.cell(row=current_row, column=col_idx, value=col_name)
        current_row += 1

        for _, row_data in t_df.iterrows():
            for col_idx, val in enumerate(row_data, start=1):
                ws.cell(row=current_row, column=col_idx, value=val)
            current_row += 1

        end_row = current_row - 1
        if any(k in title for k in ["그룹핑", "라인별", "분포", "판정"]) and chart_target_coords is None:
            chart_target_coords = (header_row, end_row, title)

        current_row += 2

    _style_worksheet(ws)

    if chart_target_coords is not None:
        h_row, e_row, c_title = chart_target_coords
        if e_row > h_row:
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = f"[핵심 집계 차트] {c_title}"
            chart.y_axis.title = "건수"
            chart.x_axis.title = "그룹 / 항목"
            chart.width = 16
            chart.height = 10

            val_col = 2
            for c_idx in range(2, 6):
                h_val = str(ws.cell(row=h_row, column=c_idx).value or "")
                if any(k in h_val for k in ["불량수", "건수"]):
                    val_col = c_idx
                    break

            data_ref = Reference(ws, min_col=val_col, min_row=h_row, max_row=e_row)
            cats_ref = Reference(ws, min_col=1, min_row=h_row + 1, max_row=e_row)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            chart.legend = None
            ws.add_chart(chart, "I2")


def export_analysis_to_excel(
    df: pd.DataFrame,
    summary_df: pd.DataFrame,
    report_text: str,
    elapsed_time: float,
    mode: str,
    aux_sheets: dict[str, pd.DataFrame] | None = None,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        report_rows = [
            {"구분 항목": "분석 AI 엔진", "상세 분석 내용": mode},
            {"구분 항목": "총 소요 시간", "상세 분석 내용": f"{elapsed_time}초"},
            {"구분 항목": "핵심 데이터 건수", "상세 분석 내용": f"총 {len(df)}건"},
        ]
        report_rows.extend(parse_report_into_sections(report_text))

        pd.DataFrame(report_rows).to_excel(writer, index=False, sheet_name="1_AI분석리포트")
        _style_worksheet(writer.sheets["1_AI분석리포트"])

        _write_stat_tables_and_chart_to_sheet(writer, df, aux_sheets=aux_sheets, sheet_name="2_시트통합요약_및_핵심집계")

        df.to_excel(writer, index=False, sheet_name="3_핵심원본데이터")
        _style_worksheet(writer.sheets["3_핵심원본데이터"])

    return output.getvalue()


def export_combined_analysis_to_excel(
    df: pd.DataFrame,
    saved_results: dict,
    model_labels: dict[str, str],
    aux_sheets: dict[str, pd.DataFrame] | None = None,
) -> bytes:
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        comp_rows = []
        for m_code, res in saved_results.items():
            sections = parse_report_into_sections(res.get("report", ""))
            sec_map = {s["구분 항목"]: s["상세 분석 내용"] for s in sections}

            status_val = sec_map.get("자동 기준 판정 요약", "판정 완료")
            trend_val = next((v for k, v in sec_map.items() if "1." in k or "추세" in k or "현황" in k), "")
            anom_val = next((v for k, v in sec_map.items() if "2." in k or "이상" in k or "주의" in k or "색상" in k), "")
            cause_val = next((v for k, v in sec_map.items() if "3." in k or "원인" in k or "가설" in k), "")
            action_val = next((v for k, v in sec_map.items() if "4." in k or "조치" in k), "")

            comp_rows.append({
                "AI 모델명": model_labels.get(m_code, m_code),
                "소요 시간(초)": float(res.get("elapsed_time", 0.0)),
                "자동 판정 요약": _format_section_text_for_excel(status_val),
                "1. 시트 종합 현황 및 추세": _format_section_text_for_excel(trend_val) or "개별 시트 참조",
                "2. 주요 이상·강조 항목 해석": _format_section_text_for_excel(anom_val) or "개별 시트 참조",
                "3. 데이터 근거 원인 가설": _format_section_text_for_excel(cause_val) or "개별 시트 참조",
                "4. 실무 조치 제안": _format_section_text_for_excel(action_val) or "개별 시트 참조",
            })

        comp_df = pd.DataFrame(comp_rows)
        comp_df.to_excel(writer, index=False, sheet_name="1_모델별_종합비교")
        ws_comp = writer.sheets["1_모델별_종합비교"]
        _style_worksheet(ws_comp)

        if len(comp_rows) >= 1:
            speed_chart = BarChart()
            speed_chart.type = "col"
            speed_chart.style = 11
            speed_chart.title = "AI 모델별 분석 소요 시간(초) 비교"
            speed_chart.y_axis.title = "소요 시간(초)"
            speed_chart.x_axis.title = "AI 모델"
            speed_chart.width = 15
            speed_chart.height = 9

            data_ref = Reference(ws_comp, min_col=2, min_row=1, max_row=len(comp_rows) + 1)
            cats_ref = Reference(ws_comp, min_col=1, min_row=2, max_row=len(comp_rows) + 1)
            speed_chart.add_data(data_ref, titles_from_data=True)
            speed_chart.set_categories(cats_ref)
            speed_chart.legend = None
            ws_comp.add_chart(speed_chart, "I2")

        _write_stat_tables_and_chart_to_sheet(writer, df, aux_sheets=aux_sheets, sheet_name="2_시트통합요약_및_핵심집계")

        df.to_excel(writer, index=False, sheet_name="3_핵심원본데이터")
        _style_worksheet(writer.sheets["3_핵심원본데이터"])

        for m_code, res in saved_results.items():
            safe_sheet = f"리포트_{m_code}"[:31]
            m_rows = [
                {"구분 항목": "AI 모델명", "상세 분석 내용": model_labels.get(m_code, m_code)},
                {"구분 항목": "소요 시간", "상세 분석 내용": f"{res.get('elapsed_time', 0.0)}초"},
            ]
            m_rows.extend(parse_report_into_sections(res.get("report", "")))
            pd.DataFrame(m_rows).to_excel(writer, index=False, sheet_name=safe_sheet)
            _style_worksheet(writer.sheets[safe_sheet])

    return output.getvalue()