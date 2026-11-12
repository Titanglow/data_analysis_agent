import io
import re
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference


def parse_report_into_sections(report_text: str) -> list[dict[str, str]]:
    """
    기존의 '리포트 본문 1, 본문 2...' 줄 단위 쪼개기를 없애고,
    경영진 보고서의 의미 있는 목차(판정 배너, 1.추세, 2.이상항목, 3.원인가설, 4.조치제안)별로 묶어 깔끔하게 반환합니다.
    """
    if not report_text or not report_text.strip():
        return [{"구분 항목": "분석 리포트", "상세 분석 내용": "내용 없음"}]

    # 마크다운 강조 기호(** 등)를 정리해 엑셀에서 읽기 편하게 변환
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
        # '### 1. 핵심 추세' 또는 '1. **추세**' 등의 대제목 감지
        header_match = re.match(r"^(?:#+\s*)?(\d+\.\s*.+)$", line)
        if header_match:
            if current_body:
                sections.append({
                    "구분 항목": current_title,
                    "상세 분석 내용": "\n".join(current_body).strip(),
                })
                current_body = []
            current_title = header_match.group(1).replace("**", "").strip()
        elif line.startswith("### "):
            if current_body:
                sections.append({
                    "구분 항목": current_title,
                    "상세 분석 내용": "\n".join(current_body).strip(),
                })
                current_body = []
            current_title = line.replace("### ", "").replace("**", "").strip()
        else:
            clean_text = line.replace("**", "")
            current_body.append(clean_text)

    if current_body:
        sections.append({
            "구분 항목": current_title,
            "상세 분석 내용": "\n".join(current_body).strip(),
        })

    return sections


def build_korean_stat_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    어떤 CSV/Excel 파일이 들어오든 컬럼을 자동 감지하여
    1) 수치 기초통계, 2) 라인/그룹별 집계, 3) 불량유형/분류별 집계, 4) 주차/일자별 추세표를 한글로 생성합니다.
    """
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

    if len(tables) <= 1:
        cat_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
        cat_rows = []
        for col in cat_cols:
            s = df[col].dropna().astype(str)
            if len(s) == 0:
                continue
            top_val = s.mode().iloc[0] if not s.mode().empty else "-"
            cat_rows.append({
                "분류 항목명": col,
                "데이터 건수": f"{len(s)}건",
                "고유 종류 수": f"{s.nunique()}종",
                "최다 빈도 값": top_val,
            })
        if cat_rows:
            tables["2. 범주형 항목별 분포 요약"] = pd.DataFrame(cat_rows)

    return tables


def _style_worksheet(ws):
    """엑셀 시트 서식(초록 헤더, 줄바꿈, 행 높이 자동 조절, 테두리, 열 너비)을 깔끔하게 정리합니다."""
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

    for row in ws.iter_rows():
        max_lines_in_row = 1
        for cell in row:
            if cell.value is not None and str(cell.value).strip() != "":
                val_str = str(cell.value)
                line_cnt = val_str.count("\n") + 1
                if line_cnt > max_lines_in_row:
                    max_lines_in_row = line_cnt

                if val_str.startswith("■"):
                    cell.font = title_font
                    cell.fill = sub_title_fill
                    cell.alignment = Alignment(vertical="center", horizontal="left")
                elif cell.row == 1 or (cell.offset(row=-1).value and str(cell.offset(row=-1).value).startswith("■")):
                    cell.font = header_font
                    cell.fill = header_fill
                    cell.alignment = Alignment(vertical="center", horizontal="center")
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

        if max_lines_in_row > 1:
            ws.row_dimensions[row[0].row].height = min(max_lines_in_row * 18 + 8, 220)
        else:
            ws.row_dimensions[row[0].row].height = 24

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or "")
            for line in val.split("\n"):
                char_len = sum(2 if ord(c) > 127 else 1 for c in line)
                if char_len > max_len:
                    max_len = char_len
        ws.column_dimensions[col_letter].width = max(min(max_len + 5, 80), 16)


def _write_stat_tables_and_chart_to_sheet(writer, df: pd.DataFrame, sheet_name: str = "2_기초통계_및_집계"):
    """한글 집계표들을 시트에 기록하고, 우측(J열)에 핵심 요약 막대 차트를 자동 삽입합니다."""
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
    """단일 모델 분석 결과를 섹션별로 깔끔하게 정리해 추출합니다 ('본문 1, 본문 2' 완전 제거)."""
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
    """다중 모델 종합 비교 시트(차트 포함) + 섹션별 정리된 개별 리포트 시트를 통합 추출합니다."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        comp_rows = []
        for m_code, res in saved_results.items():
            sections = parse_report_into_sections(res.get("report", ""))
            sec_map = {s["구분 항목"]: s["상세 분석 내용"] for s in sections}

            status_val = sec_map.get("사내 기준(IF/ELSE) 자동 판정", "기준 판정 완료")
            trend_val = next((v for k, v in sec_map.items() if "추세" in k), "")
            anom_val = next((v for k, v in sec_map.items() if "이상" in k or "불량" in k), "")
            cause_val = next((v for k, v in sec_map.items() if "원인" in k or "가설" in k), "")
            action_val = next((v for k, v in sec_map.items() if "조치" in k), "")

            comp_rows.append({
                "AI 모델명": model_labels.get(m_code, m_code),
                "소요 시간(초)": float(res.get("elapsed_time", 0.0)),
                "사내 기준 판정": status_val,
                "1. 핵심 추세 요약": trend_val or "개별 시트 참조",
                "2. 주요 이상 항목": anom_val or "개별 시트 참조",
                "3. 5M1E 원인 가설": cause_val or "개별 시트 참조",
                "4. 실무 조치 제안": action_val or "개별 시트 참조",
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