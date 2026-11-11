import io
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, Reference


def build_korean_stat_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    어떤 CSV/Excel 파일이 들어오든 컬럼을 자동 감지하여
    1) 수치 기초통계, 2) 라인/그룹별 집계, 3) 불량유형/분류별 집계, 4) 주차/일자/시간대별 추세표를 한글로 생성합니다.
    """
    tables = {}

    # 1. 수치형 항목별 한글 기초통계표
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

    # 2. 라인별 생산·품질 종합 집계표
    if "라인" in df.columns and "생산량" in df.columns and "불량수" in df.columns:
        line_rows = []
        total_prod = df["생산량"].sum()
        total_def = df["불량수"].sum()
        overall_rate = (total_def / total_prod * 100) if total_prod > 0 else 0.0

        for line_name, grp in df.groupby("라인", sort=False):
            p_sum = int(grp["생산량"].sum())
            d_sum = int(grp["불량수"].sum())
            rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
            row_rates = (grp["불량수"] / grp["생산량"] * 100).fillna(0)
            max_rate = float(row_rates.max()) if len(row_rates) > 0 else 0.0

            # 해당 라인의 최다 발생 불량유형 추출
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

    # 3. 불량유형(원인) 및 시간대(교대)별 세부 분류 집계표
    if "불량유형" in df.columns and "불량수" in df.columns:
        defect_rows = []
        total_def = df["불량수"].sum()
        for d_type, grp in df.groupby("불량유형", sort=False):
            d_sum = int(grp["불량수"].sum())
            share = (d_sum / total_def * 100) if total_def > 0 else 0.0
            main_line = grp.groupby("라인")["불량수"].sum().idxmax() if "라인" in grp.columns else "-"
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

    # 4. 주차별/시간대별 추세 및 전주 대비 증감(%p) 표
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

    # 5. 일반 범주형 파일일 때의 자동 분류 요약
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
    """엑셀 시트 서식(초록 헤더, 줄바꿈, 테두리, 열 너비)을 자동 정리합니다."""
    header_fill = PatternFill(start_color="059669", end_color="059669", fill_type="solid")
    sub_title_fill = PatternFill(start_color="E6F4EA", end_color="E6F4EA", fill_type="solid")
    header_font = Font(name="맑은 고딕", size=10, bold=True, color="FFFFFF")
    title_font = Font(name="맑은 고딕", size=11, bold=True, color="047857")
    body_font = Font(name="맑은 고딕", size=10)
    thin_border = Border(
        left=Side(style="thin", color="D1D5DB"),
        right=Side(style="thin", color="D1D5DB"),
        top=Side(style="thin", color="D1D5DB"),
        bottom=Side(style="thin", color="D1D5DB"),
    )

    for row in ws.iter_rows():
        for cell in row:
            if cell.value is not None and str(cell.value).strip() != "":
                val_str = str(cell.value)
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
                    cell.font = body_font
                    cell.alignment = Alignment(vertical="center", wrap_text=True)
                    cell.border = thin_border

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val = str(cell.value or "")
            first_line = val.split("\n")[0]
            char_len = sum(2 if ord(c) > 127 else 1 for c in first_line)
            if char_len > max_len:
                max_len = char_len
        ws.column_dimensions[col_letter].width = max(min(max_len + 5, 65), 14)


def _write_stat_tables_and_chart_to_sheet(writer, df: pd.DataFrame, sheet_name: str = "2_기초통계_및_집계"):
    """한글 집계표들을 시트에 기록하고, 우측(J열)에 핵심 요약 막대 차트를 자동 삽입합니다."""
    tables = build_korean_stat_tables(df)
    ws = writer.book.create_sheet(title=sheet_name)

    current_row = 1
    line_table_coords = None  # (header_row, end_row) 기록용

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

    # 엑셀 시트 우측(J2 위치)에 간략 핵심 차트(불량수 비교 막대그래프) 자동 삽입
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

            # 4번째 열(총 불량수) 또는 3번째 열(수치) 참조
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
    """단일 모델 분석 결과 + 한글 통계표(차트 포함) + 원본 데이터를 엑셀로 추출합니다."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        report_lines = [line.strip() for line in report_text.splitlines() if line.strip()]
        report_rows = [
            {"항목": "분석 엔진 모드", "내용": mode},
            {"항목": "총 소요 시간(초)", "내용": f"{elapsed_time}초"},
            {"항목": "분석 데이터 건수", "내용": f"{len(df)}건"},
        ]
        for idx, line in enumerate(report_lines, start=1):
            report_rows.append({"항목": f"리포트 본문 {idx}", "내용": line})

        pd.DataFrame(report_rows).to_excel(writer, index=False, sheet_name="1_AI분석리포트")
        _style_worksheet(writer.sheets["1_AI분석리포트"])

        # 시트 2: 한글 통계표 + 요약 차트 삽입
        _write_stat_tables_and_chart_to_sheet(writer, df, sheet_name="2_기초통계_및_집계")

        # 시트 3: 원본 데이터
        df.to_excel(writer, index=False, sheet_name="3_원본데이터")
        _style_worksheet(writer.sheets["3_원본데이터"])

    return output.getvalue()


def export_combined_analysis_to_excel(
    df: pd.DataFrame,
    saved_results: dict,
    model_labels: dict[str, str],
) -> bytes:
    """다중 모델 종합 비교 시트(소요시간 비교 차트 포함) + 한글 통계/불량분류 시트(차트 포함)를 통합 추출합니다."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1. 모델별 종합 비교 시트
        comp_rows = []
        for m_code, res in saved_results.items():
            anomalies = res.get("analysis_result", {}).get("anomalies", [])
            trends = res.get("analysis_result", {}).get("trends", [])
            comp_rows.append({
                "AI 모델명": model_labels.get(m_code, m_code),
                "모드 코드": m_code,
                "소요 시간(초)": float(res.get("elapsed_time", 0.0)),
                "핵심 이상 징후": " / ".join(anomalies) if anomalies else "본문 참조",
                "주요 트렌드": " / ".join(trends) if trends else "본문 참조",
                "AI 분석 리포트 전문": res.get("report", ""),
            })
        comp_df = pd.DataFrame(comp_rows)
        comp_df.to_excel(writer, index=False, sheet_name="1_모델별_종합비교")
        ws_comp = writer.sheets["1_모델별_종합비교"]
        _style_worksheet(ws_comp)

        # 1번 시트 우측(H2)에 [AI 모델별 응답 속도(초) 비교 막대 차트] 자동 삽입
        if len(comp_rows) >= 1:
            speed_chart = BarChart()
            speed_chart.type = "col"
            speed_chart.style = 11
            speed_chart.title = "AI 모델별 분석 소요 시간(초) 비교"
            speed_chart.y_axis.title = "소요 시간(초)"
            speed_chart.x_axis.title = "AI 모델"
            speed_chart.width = 15
            speed_chart.height = 9

            data_ref = Reference(ws_comp, min_col=3, min_row=1, max_row=len(comp_rows) + 1)
            cats_ref = Reference(ws_comp, min_col=1, min_row=2, max_row=len(comp_rows) + 1)
            speed_chart.add_data(data_ref, titles_from_data=True)
            speed_chart.set_categories(cats_ref)
            speed_chart.legend = None
            ws_comp.add_chart(speed_chart, "H2")

        # 2. 한글 기초통계 및 라인/불량유형별 집계 시트 (차트 포함)
        _write_stat_tables_and_chart_to_sheet(writer, df, sheet_name="2_기초통계_및_집계")

        # 3. 원본 데이터 시트
        df.to_excel(writer, index=False, sheet_name="3_원본데이터")
        _style_worksheet(writer.sheets["3_원본데이터"])

        # 4. 개별 모델 상세 리포트 시트
        for m_code, res in saved_results.items():
            safe_sheet = f"리포트_{m_code}"[:31]
            lines = [line.strip() for line in res.get("report", "").splitlines() if line.strip()]
            m_rows = [
                {"구분": "모델명", "내용": model_labels.get(m_code, m_code)},
                {"구분": "소요 시간", "내용": f"{res.get('elapsed_time', 0.0)}초"},
            ] + [{"구분": f"본문 {i}", "내용": l} for i, l in enumerate(lines, 1)]
            pd.DataFrame(m_rows).to_excel(writer, index=False, sheet_name=safe_sheet)
            _style_worksheet(writer.sheets[safe_sheet])

    return output.getvalue()