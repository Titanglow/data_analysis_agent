import io
import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter


def build_korean_stat_tables(df: pd.DataFrame) -> dict[str, pd.DataFrame]:
    """
    영어(count, mean, 25% 등)로 나오던 기존 describe()를 대체하여
    한눈에 보기 쉬운 100% 한글 기초통계표, 라인별 집계표, 주차별 추세표를 생성합니다.
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

    # 2. 라인별 생산·불량 종합 집계표 (라인 컬럼이 있을 때 자동 생성)
    if "라인" in df.columns and "생산량" in df.columns and "불량수" in df.columns:
        line_rows = []
        total_prod = df["생산량"].sum()
        total_def = df["불량수"].sum()
        overall_rate = (total_def / total_prod * 100) if total_prod > 0 else 0.0

        for line_name, grp in df.groupby("라인", sort=False):
            p_sum = grp["생산량"].sum()
            d_sum = grp["불량수"].sum()
            rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
            row_rates = (grp["불량수"] / grp["생산량"] * 100).fillna(0)
            max_rate = float(row_rates.max()) if len(row_rates) > 0 else 0.0

            status = "🚨 [확인 필요] 기준(1.5%) 초과" if rate > 1.5 else "✅ [이상 없음] 기준 이내"
            vs_avg = "평균 상회(주의)" if rate > overall_rate else "평균 이하(양호)"

            row_dict = {
                "라인 구분": str(line_name),
                "집계 건수": f"{len(grp)}건",
                "총 생산량(개)": int(p_sum),
                "평균 생산량(개)": round(float(grp["생산량"].mean()), 1),
                "총 불량수(개)": int(d_sum),
                "평균 불량률(%)": f"{rate:.2f}%",
                "최고 불량률(%)": f"{max_rate:.2f}%",
                "전체 평균 대비": vs_avg,
                "사내 기준(1.5%) 판정": status,
            }
            if "매출액" in grp.columns:
                row_dict["총 매출액"] = int(grp["매출액"].sum())
            line_rows.append(row_dict)

        if line_rows:
            tables["2. 라인별 생산·품질 종합 집계"] = pd.DataFrame(line_rows)

    # 3. 주차별 추세 및 전주 대비 불량률 증감(%p) 표
    time_col = "주차" if "주차" in df.columns else ("일자" if "일자" in df.columns else None)
    if time_col and "생산량" in df.columns and "불량수" in df.columns:
        week_rows = []
        prev_rate = None
        for t_val, grp in df.groupby(time_col, sort=False):
            p_sum = grp["생산량"].sum()
            d_sum = grp["불량수"].sum()
            rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
            if prev_rate is None:
                diff_str = "- (기준 주차)"
            else:
                diff = rate - prev_rate
                sign = "+" if diff > 0 else ""
                diff_str = f"{sign}{diff:.2f}%p"
            prev_rate = rate

            w_dict = {
                time_col: str(t_val),
                "총 생산량(개)": int(p_sum),
                "총 불량수(개)": int(d_sum),
                "주간 불량률(%)": f"{rate:.2f}%",
                "전주 대비 증감(%p)": diff_str,
            }
            if "매출액" in grp.columns:
                w_dict["주간 매출액"] = int(grp["매출액"].sum())
            week_rows.append(w_dict)

        if week_rows:
            tables[f"3. {time_col}별 품질 추세 및 전주 대비 증감"] = pd.DataFrame(week_rows)

    # 4. 문자/범주형 항목 요약 (라인·주차 등이 없는 일반 파일일 때 대비)
    if not tables:
        cat_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()
        cat_rows = []
        for col in cat_cols:
            s = df[col].dropna().astype(str)
            if len(s) == 0:
                continue
            top_val = s.mode().iloc[0] if not s.mode().empty else "-"
            cat_rows.append({
                "항목명": col,
                "유효 데이터 수": f"{len(s)}건",
                "고유값 종류 수": f"{s.nunique()}종",
                "최빈값(가장 많은 항목)": top_val,
            })
        if cat_rows:
            tables["1. 범주형 항목 요약"] = pd.DataFrame(cat_rows)

    return tables


def _style_worksheet(ws):
    """엑셀 시트의 헤더 색상, 테두리, 열 너비, 줄바꿈을 깔끔하게 자동 정렬합니다."""
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
            # 한글 가중치 반영하여 열 너비 계산
            char_len = sum(2 if ord(c) > 127 else 1 for c in first_line)
            if char_len > max_len:
                max_len = char_len
        ws.column_dimensions[col_letter].width = max(min(max_len + 5, 65), 14)


def _write_stat_tables_to_sheet(writer, df: pd.DataFrame, sheet_name: str = "2_기초통계_및_집계"):
    """하나의 시트에 한글화된 기초통계표, 라인별 집계표, 주차별 추세표를 순서대로 배치합니다."""
    tables = build_korean_stat_tables(df)
    ws = writer.book.create_sheet(title=sheet_name)

    current_row = 1
    for title, t_df in tables.items():
        ws.cell(row=current_row, column=1, value=f"■ {title}")
        current_row += 1

        # 컬럼 헤더 작성
        for col_idx, col_name in enumerate(t_df.columns, start=1):
            ws.cell(row=current_row, column=col_idx, value=col_name)
        current_row += 1

        # 데이터 행 작성
        for _, row_data in t_df.iterrows():
            for col_idx, val in enumerate(row_data, start=1):
                ws.cell(row=current_row, column=col_idx, value=val)
            current_row += 1

        current_row += 2  # 표 사이 여백

    _style_worksheet(ws)


def export_analysis_to_excel(
    df: pd.DataFrame,
    summary_df: pd.DataFrame,
    report_text: str,
    elapsed_time: float,
    mode: str,
) -> bytes:
    """단일 모델 분석 결과를 3개 시트(1_AI분석리포트, 2_기초통계_및_집계, 3_원본데이터)로 추출합니다."""
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 시트 1: AI 분석 리포트
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

        # 시트 2: 100% 한글화된 기초통계 및 라인/주차별 집계표
        _write_stat_tables_to_sheet(writer, df, sheet_name="2_기초통계_및_집계")

        # 시트 3: 필터링된 원본 데이터
        df.to_excel(writer, index=False, sheet_name="3_원본데이터")
        _style_worksheet(writer.sheets["3_원본데이터"])

    return output.getvalue()


def export_combined_analysis_to_excel(
    df: pd.DataFrame,
    saved_results: dict,
    model_labels: dict[str, str],
) -> bytes:
    """
    다중 모델 동시 실행 결과를 하나의 마스터 엑셀 파일로 통합 추출합니다.
    - 1_모델별_종합비교: 모든 모델의 소요 시간, 이상 징후, 리포트 전문을 한 시트에서 비교
    - 2_기초통계_및_집계: 한글화된 통계 요약 및 라인/주차별 집계표
    - 3_원본데이터: 조회된 원본 데이터
    - 개별 모델 시트: 각 AI 모델별 상세 리포트
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # 1. 종합 비교 요약 시트 생성
        comp_rows = []
        for m_code, res in saved_results.items():
            anomalies = res.get("analysis_result", {}).get("anomalies", [])
            trends = res.get("analysis_result", {}).get("trends", [])
            comp_rows.append({
                "AI 모델명": model_labels.get(m_code, m_code),
                "모드 코드": m_code,
                "소요 시간(초)": res.get("elapsed_time", 0.0),
                "핵심 이상 징후": " / ".join(anomalies) if anomalies else "본문 참조",
                "주요 트렌드": " / ".join(trends) if trends else "본문 참조",
                "AI 분석 리포트 전문": res.get("report", ""),
            })
        comp_df = pd.DataFrame(comp_rows)
        comp_df.to_excel(writer, index=False, sheet_name="1_모델별_종합비교")
        _style_worksheet(writer.sheets["1_모델별_종합비교"])

        # 2. 한글 기초통계 및 라인/주차별 집계 시트
        _write_stat_tables_to_sheet(writer, df, sheet_name="2_기초통계_및_집계")

        # 3. 원본 데이터 시트
        df.to_excel(writer, index=False, sheet_name="3_원본데이터")
        _style_worksheet(writer.sheets["3_원본데이터"])

        # 4. 각 모델별 개별 상세 시트 추가
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