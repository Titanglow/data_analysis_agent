import io
import math
import re
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference


def _format_section_text_for_excel(text: str) -> str:
    """
    엑셀 셀('1_모델별_종합비교'의 '4. 실무 조치 제안' 등) 안에서 문장들이 한 줄로 뭉치지 않도록
    불릿(-, *, •), 시점 구분(즉시 조치, 근본 대책, 기대 효과), 부서명([생산팀] 등) 앞마다
    명시적인 줄바꿈(\n)을 강제로 삽입하고 마크다운 기호를 깔끔하게 정돈합니다.
    """
    if not text or not text.strip():
        return ""

    cleaned = text.replace("**", "").replace("\r\n", "\n").replace("\r", "\n")

    # 한 줄 안에 이어져 작성된 주요 키워드 및 불릿 앞에 강제 줄바꿈(\n) 삽입
    cleaned = re.sub(r"(?<!\n)\s+(-\s+|\*\s+|•\s+)", r"\n• ", cleaned)
    cleaned = re.sub(
        r"(?<!\n)\s*(즉시\s*조치|근본\s*대책|중장기\s*대책|기대\s*효과|설비\(Machine\)|작업자/교대\(Man\)|자재\(Material\)|공정조건\(Method\))",
        r"\n▶ \1",
        cleaned,
    )
    cleaned = re.sub(r"(?<!\n)\s*(\[(?:생산|설비|품질|정량)[^\]]*\])", r"\n  \1", cleaned)

    formatted_lines = []
    for raw_line in cleaned.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        # 마크다운 표 구분선(|---|---|)은 엑셀 셀 가독성을 위해 제외
        if re.match(r"^\|?[\s\-:|]+\|?$", line):
            continue
        # 마크다운 표 행(| 순위 | 관리대상 | ...)은 보기 좋은 텍스트로 변환
        if line.startswith("|") and line.endswith("|"):
            cells = [c.strip() for c in line.strip("|").split("|") if c.strip()]
            line = " ▪ " + " | ".join(cells)
        elif line.startswith("- ") or line.startswith("* "):
            line = "• " + line[2:].strip()

        formatted_lines.append(line)

    return "\n".join(formatted_lines).strip()


def parse_report_into_sections(report_text: str) -> list[dict[str, str]]:
    """
    리포트를 목차별(판정 배너, 1.추세, 2.이상항목, 3.원인가설, 4.조치제안)로 분류하고,
    각 섹션 내부 텍스트에 엑셀 전용 줄바꿈 포맷팅을 적용합니다.
    """
    if not report_text or not report_text.strip():
        return [{"구분 항목": "분석 리포트", "상세 분석 내용": "내용 없음"}]

    cleaned_lines = []
    model_info = ""
    status_banner = ""

    for raw_line in report_text.splitlines():
        line = raw_line.strip()
        if not line:
            continue
        if line.startswith("*(연결 모델:") or line.startswith("*(로컬 모델:"):
            model_info = line.strip("*() ")
            continue
        if "확인 필요" in line or "이상 없음" in line:
            status_banner = line.lstrip("> ").replace("**", "")
            continue
        cleaned_lines.append(line)

    sections = []
    if model_info:
        sections.append({"구분 항목": "구동 AI 엔진 상세", "상세 분석 내용": model_info})
    if status_banner:
        sections.append({"구분 항목": "사내 기준(IF/ELSE) 자동 판정", "상세 분석 내용": status_banner})

    current_title = "핵심 분석 요약"
    current_body = []

    for line in cleaned_lines:
        header_match = re.match(r"^(?:#+\s*)?(\d+\.\s*.+)$", line)
        if header_match:
            if current_body:
                sections.append({
                    "구분 항목": current_title,
                    "상세 분석 내용": _format_section_text_for_excel("\n".join(current_body)),
                })
                current_body = []
            current_title = header_match.group(1).replace("**", "").strip()
        elif line.startswith("### "):
            if current_body:
                sections.append({
                    "구분 항목": current_title,
                    "상세 분석 내용": _format_section_text_for_excel("\n".join(current_body)),
                })
                current_body = []
            current_title = line.replace("### ", "").replace("**", "").strip()
        else:
            current_body.append(line)

    if current_body:
        sections.append({
            "구분 항목": current_title,
            "상세 분석 내용": _format_section_text_for_excel("\n".join(current_body)),
        })

    return sections


def build_korean_stat_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """수치 기초통계, 라인별 집계, 불량유형별 점유율, 주차별 추세표를 100% 한글로 생성합니다."""
    tables = {}

    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    if num_cols:
        stat_rows = []
        for col in num_cols:
            s = df[col].dropna()
            if len(s) == 0:
                continue
            stat_rows.append({
                "분석 항목": col,
                "데이터 건수": f"{len(s):,}건",
                "합계": round(float(s.sum()), 2),
                "평균": round(float(s.mean()), 2),
                "최솟값": round(float(s.min()), 2),
                "중앙값(50%)": round(float(s.median()), 2),
                "최댓값": round(float(s.max()), 2),
                "표준편차": round(float(s.std()), 2) if len(s) > 1 else 0.0,
            })
        if stat_rows:
            tables["1. 수치 항목별 기초 통계 요약"] = pd.DataFrame(stat_rows)

    line_col = "라인" if "라인" in df.columns else ("생산라인" if "생산라인" in df.columns else None)
    if line_col and "생산량" in df.columns and "불량수" in df.columns:
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
            row_dict = {
                "라인 구분": str(line_name),
                "집계 건수": f"{len(grp)}건",
                "총 생산량(개)": p_sum,
                "총 불량수(개)": d_sum,
                "평균 불량률(%)": f"{rate:.2f}%",
                "최고 불량률(%)": f"{max_rate:.2f}%",
                "주요 불량유형(최다)": top_defect,
                "사내 기준(1.5%) 판정": status,
            }
            if "설비가동률" in grp.columns:
                row_dict["평균 가동률(%)"] = f"{grp['설비가동률'].mean():.1f}%"
            line_rows.append(row_dict)

        if line_rows:
            tables["2. 라인별 생산·품질 및 주요 불량 종합 집계"] = pd.DataFrame(line_rows)

    if "불량유형" in df.columns and "불량수" in df.columns:
        defect_rows = []
        total_def = df["불량수"].sum()
        for d_type, grp in df.groupby("불량유형", sort=False):
            d_sum = int(grp["불량수"].sum())
            share = (d_sum / total_def * 100) if total_def > 0 else 0.0
            main_line = grp.groupby(line_col)["불량수"].sum().idxmax() if line_col else "-"
            main_shift = grp.groupby("시간대")["불량수"].sum().idxmax() if "시간대" in grp.columns else "-"
            defect_rows.append({
                "불량 유형(분류)": str(d_type),
                "발생 건수(빈도)": f"{len(grp)}회",
                "총 불량수(개)": d_sum,
                "전체 불량 중 비중(%)": f"{share:.1f}%",
                "최다 발생 라인": str(main_line),
                "최다 발생 시간대": str(main_shift),
            })
        if defect_rows:
            defect_df = pd.DataFrame(defect_rows).sort_values(by="총 불량수(개)", ascending=False)
            tables["3. 불량 유형(원인)별 분류 및 점유율 집계"] = defect_df

    time_col = "주차" if "주차" in df.columns else ("일자" if "일자" in df.columns else None)
    if time_col and "생산량" in df.columns and "불량수" in df.columns:
        week_rows = []
        prev_rate = None
        for t_val, grp in df.groupby(time_col, sort=False):
            p_sum = int(grp["생산량"].sum())
            d_sum = int(grp["불량수"].sum())
            rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
            if prev_rate is None:
                diff_str = "- (기준 시점)"
            else:
                diff = rate - prev_rate
                sign = "+" if diff > 0 else ""
                diff_str = f"{sign}{diff:.2f}%p"
            prev_rate = rate

            w_dict = {
                time_col: str(t_val),
                "총 생산량(개)": p_sum,
                "총 불량수(개)": d_sum,
                "불량률(%)": f"{rate:.2f}%",
                "이전 대비 증감(%p)": diff_str,
            }
            if "불량유형" in grp.columns:
                w_dict["해당 시점 최다 불량"] = str(grp.groupby("불량유형")["불량수"].sum().idxmax())
            week_rows.append(w_dict)

        if week_rows:
            tables[f"4. {time_col}별 품질 추세 및 증감(%p)"] = pd.DataFrame(week_rows)

    return tables


def _style_worksheet(ws):
    """
    엑셀 시트의 모든 셀에 자동 줄바꿈(wrap_text=True)을 강제 적용하고,
    열 너비 대비 자동 접힘(Wrap) 줄 수까지 계산하여 행 높이를 넉넉하게 자동 지정합니다.
    """
    header_fill = PatternFill(start_color="059669", end_color="059669", fill_type="solid")
    sub_title_fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
    key_col_fill = PatternFill(start_color="F9FAFB", end_color="F9FAFB", fill_type="solid")
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

    # 1단계: 열 너비(Column Width) 먼저 설정
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
        # 서술형 리포트 열은 넓게(최대 65), 일반 열은 적당히 배정
        width = max(min(max_len + 5, 65), 16)
        ws.column_dimensions[col_letter].width = width
        col_widths[col_idx] = width

    # 2단계: 셀 서식(wrap_text=True) 및 실제 줄바꿈 줄 수를 반영한 행 높이(Row Height) 설정
    for row in ws.iter_rows():
        max_visual_lines = 1
        for cell in row:
            if cell.value is not None and str(cell.value).strip() != "":
                val_str = str(cell.value)
                c_width = max(col_widths.get(cell.column, 30) - 4, 20)

                # 명시적 줄바꿈(\n) + 긴 문장이 셀 폭을 넘어 자동 줄바꿈되는 줄 수까지 합산
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
                    cell.border = thin_border

        if max_visual_lines > 1:
            ws.row_dimensions[row[0].row].height = min(max_visual_lines * 19 + 12, 380)
        else:
            ws.row_dimensions[row[0].row].height = 26


def _write_stat_tables_and_chart_to_sheet(writer, df: pd.DataFrame, sheet_name: str = "2_기초통계_및_집계"):
    """한글 집계표들을 시트에 기록하고 우측에 요약 막대 차트를 삽입합니다."""
    tables = build_korean_stat_tables(df)
    ws = writer.book.create_sheet(title=sheet_name)

    current_row = 1
    line_table_coords = None

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
        if "라인별" in title or "불량 유형" in title:
            if line_table_coords is None:
                line_table_coords = (header_row, end_row, title)

        current_row += 2

    _style_worksheet(ws)

    if line_table_coords is not None:
        h_row, e_row, c_title = line_table_coords
        if e_row > h_row:
            chart = BarChart()
            chart.type = "col"
            chart.style = 10
            chart.title = f"[요약 차트] {c_title}"
            chart.y_axis.title = "불량수(개)"
            chart.x_axis.title = "구분"
            chart.width = 16
            chart.height = 10

            val_col = 4 if ws.cell(row=h_row, column=4).value and "불량수" in str(ws.cell(row=h_row, column=4).value) else 3
            data_ref = Reference(ws, min_col=val_col, min_row=h_row, max_row=e_row)
            cats_ref = Reference(ws, min_col=1, min_row=h_row + 1, max_row=e_row)
            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            chart.legend = None
            ws.add_chart(chart, "J2")


def export_analysis_to_excel(
    df: pd.DataFrame,
    summary_df: pd.DataFrame,
    report_text: str,
    elapsed_time: float,
    mode: str,
) -> bytes:
    """단일 모델 분석 결과를 섹션별 줄바꿈이 적용된 엑셀로 추출합니다."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        report_rows = [
            {"구분 항목": "분석 AI 엔진", "상세 분석 내용": mode},
            {"구분 항목": "총 소요 시간", "상세 분석 내용": f"{elapsed_time}초"},
            {"구분 항목": "분석 대상 데이터", "상세 분석 내용": f"총 {len(df)}건"},
        ]
        report_rows.extend(parse_report_into_sections(report_text))

        pd.DataFrame(report_rows).to_excel(writer, index=False, sheet_name="1_AI분석리포트")
        _style_worksheet(writer.sheets["1_AI분석리포트"])

        _write_stat_tables_and_chart_to_sheet(writer, df, sheet_name="2_기초통계_및_집계")

        df.to_excel(writer, index=False, sheet_name="3_원본데이터")
        _style_worksheet(writer.sheets["3_원본데이터"])

    return output.getvalue()


def export_combined_analysis_to_excel(
    df: pd.DataFrame,
    saved_results: dict,
    model_labels: dict[str, str],
) -> bytes:
    """
    다중 모델 종합 비교 시트(1~4번 모든 항목에 강제 줄바꿈 및 자동 행 높이 적용)와
    한글 통계 시트, 개별 모델 시트를 통합 추출합니다.
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        comp_rows = []
        for m_code, res in saved_results.items():
            sections = parse_report_into_sections(res.get("report", ""))
            sec_map = {s["구분 항목"]: s["상세 분석 내용"] for s in sections}

            status_val = sec_map.get("사내 기준(IF/ELSE) 자동 판정", "기준 판정 완료")
            trend_val = next((v for k, v in sec_map.items() if "1." in k or "추세" in k), "")
            anom_val = next((v for k, v in sec_map.items() if "2." in k or "이상" in k or "불량" in k), "")
            cause_val = next((v for k, v in sec_map.items() if "3." in k or "원인" in k or "가설" in k), "")
            action_val = next((v for k, v in sec_map.items() if "4." in k or "조치" in k), "")

            comp_rows.append({
                "AI 모델명": model_labels.get(m_code, m_code),
                "소요 시간(초)": float(res.get("elapsed_time", 0.0)),
                "사내 기준 판정": _format_section_text_for_excel(status_val),
                "1. 핵심 추세 요약": _format_section_text_for_excel(trend_val) or "개별 시트 참조",
                "2. 주요 이상 항목": _format_section_text_for_excel(anom_val) or "개별 시트 참조",
                "3. 5M1E 원인 가설": _format_section_text_for_excel(cause_val) or "개별 시트 참조",
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

        _write_stat_tables_and_chart_to_sheet(writer, df, sheet_name="2_기초통계_및_집계")

        df.to_excel(writer, index=False, sheet_name="3_원본데이터")
        _style_worksheet(writer.sheets["3_원본데이터"])

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