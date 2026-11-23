# data_agent/excel_exporter.py
import io
import pandas as pd
from openpyxl.drawing.image import Image as ExcelImage
import plotly.io as pio
import plotly.express as px
from datetime import datetime
import re
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import nsdecls, qn

from data_agent.config import (
    DISTINCT_PALETTE,
    DEFAULT_DEFECT_THRESHOLD,
    MODEL_LABELS
)


def shorten_long_label(label: str, max_len: int = 35) -> str:
    s = str(label)
    if len(s) > max_len:
        return s[:max_len] + "..."
    return s


def is_manufacturing_df(df: pd.DataFrame) -> bool:
    cols = df.columns.tolist()
    return any(c in cols for c in ["생산량", "불량수"])


def parse_report_into_sections(report_text: str) -> list[dict[str, str]]:
    sections = [
        {"구분 항목": "1. 전체 현황 및 요약", "키워드": ["전체 검증", "현황 및 요약", "전체 현황", "요약"]},
        {"구분 항목": "2. 주요 분류·시나리오 분포", "키워드": ["실제 분류", "시나리오", "분류 컬럼별", "분포"]},
        {"구분 항목": "3. 품질 리스크·원인 진단", "키워드": ["리스크", "취약 구간", "이상 징후", "이상치", "결함"]},
        {"구분 항목": "4. 실무 조치 제안 (R&R)", "키워드": ["실무 조치", "조치 제안", "로드맵", "대책", "R&R"]}
    ]

    parsed_data = []
    lines = report_text.splitlines()
    current_sec = "요약 전 서론"
    buffer = []

    for line in lines:
        matched = False
        for sec in sections:
            if any(re.search(rf"(?:^|\s|\d\.?\s*\*?){kw}", line) for kw in sec["키워드"]) and len(line) < 60:
                if buffer:
                    parsed_data.append({"구분 항목": current_sec, "상세 분석 내용": "\n".join(buffer).strip()})
                    buffer = []
                current_sec = sec["구분 항목"]
                matched = True
                break
        if not matched:
            buffer.append(line)

    if buffer:
        parsed_data.append({"구분 항목": current_sec, "상세 분석 내용": "\n".join(buffer).strip()})

    filtered = [item for item in parsed_data if item["구분 항목"] != "요약 전 서론"]
    return filtered if filtered else [{"구분 항목": "전체 리포트 원문", "상세 분석 내용": report_text.strip()}]


def evaluate_threshold_verdict(df: pd.DataFrame, threshold: float = DEFAULT_DEFECT_THRESHOLD) -> str:
    cols = df.columns.tolist()

    if is_manufacturing_df(df):
        try:
            p_sum = df["생산량"].sum()
            d_sum = df["불량수"].sum()
            rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
            if rate > threshold:
                return f"🚨 [확인 필요] 공정 전체 불량률({rate:.2f}%)이 관리 기준({threshold}%)을 초과했습니다."
            return f"✅ [이상 없음] 공정 전체 불량률({rate:.2f}%)이 관리 기준({threshold}%) 이내로 양호합니다."
        except Exception:
            pass

    issue_msgs = []
    for col in cols:
        if str(col).lower() in ("_polarion", "id", "tc id", "번호", "시트명", "파일명"):
            continue

        s_col = df[col].astype(str).str.strip().str.upper()

        empty_count = s_col.isin(["-", "", "NAN", "NONE", "(빈칸)", "N/A"]).sum() + s_col.str.startswith("[빈칸").sum()
        notrun_count = s_col.str.contains("NOT RUN|미실행").sum()
        fail_count = s_col.str.contains("FAIL|불합격|빨강|주의").sum()

        total_defect = empty_count + notrun_count + fail_count
        defect_rate = (total_defect / len(df) * 100) if len(df) > 0 else 0.0

        if defect_rate > threshold:
            types = []
            if empty_count > 0: types.append(f"미기입 {empty_count}건")
            if notrun_count > 0: types.append(f"미실행/N/A {notrun_count}건")
            if fail_count > 0: types.append(f"주의/FAIL {fail_count}건")
            issue_msgs.append(f"[{col}] {'/'.join(types)}({defect_rate:.1f}%)")

    if issue_msgs:
        return f"🚨 [확인 필요 — 사내 기준({threshold}%) 초과] 기준 초과 항목 감지: " + ", ".join(issue_msgs[:4])
    return f"✅ [이상 없음 — 모든 항목 사내 기준({threshold}%) 충족] 미기입 및 주의 항목이 허용 범위 내에 있습니다."


def analyze_color_cell_contexts(df: pd.DataFrame) -> pd.DataFrame:
    rows = []
    for idx, row in df.iterrows():
        tc_id = row.get("ID", row.get("TC ID", idx))
        for col in df.columns:
            val = str(row[col])
            if "🔴" in val or "빨강" in val:
                rows.append({"ID": tc_id, "위치": col, "감지된 색상": "빨강 (FAIL/주의)", "셀 내용": val.replace("🔴", "").replace("[빨강]", "").strip()})
            elif "🟢" in val or "초록" in val:
                rows.append({"ID": tc_id, "위치": col, "감지된 색상": "초록 (PASS/정상)", "셀 내용": val.replace("🟢", "").replace("[초록]", "").strip()})
            elif "🟡" in val or "노랑" in val:
                rows.append({"ID": tc_id, "위치": col, "감지된 색상": "노랑 (경고/보류)", "셀 내용": val.replace("🟡", "").replace("[노랑]", "").strip()})
            elif "🔵" in val or "파랑" in val:
                rows.append({"ID": tc_id, "위치": col, "감지된 색상": "파랑 (참고)", "셀 내용": val.replace("🔵", "").replace("[파랑]", "").strip()})
    return pd.DataFrame(rows)


def build_korean_stat_tables(df: pd.DataFrame, aux_sheets: dict = None, threshold: float = DEFAULT_DEFECT_THRESHOLD) -> dict[str, pd.DataFrame]:
    res = {}

    # 1. 컬럼별 미기입/이상치 집계 및 해당 TC ID (Title 포함) 전수 명단
    missing_data = []
    total_len = len(df)

    has_id = "ID" in df.columns
    has_tc_id = "TC ID" in df.columns
    has_title = "Title" in df.columns or "제목" in df.columns

    for col in df.columns:
        if str(col).lower() in ("_polarion", "id", "tc id", "번호", "시트명", "파일명"):
            continue

        s_col = df[col].astype(str).str.strip()
        empty_mask = s_col.isin(["-", "", "nan", "None", "(빈칸)", "N/A"]) | s_col.str.startswith("[빈칸")
        notrun_mask = s_col.str.contains("not run|미실행", case=False, na=False)
        fail_mask = s_col.str.contains("fail|불합격|빨강|주의", case=False, na=False)

        issue_mask = empty_mask | notrun_mask | fail_mask
        issue_count = issue_mask.sum()

        if issue_count > 0:
            rate = (issue_count / total_len) * 100

            issue_ids = []
            for idx, is_issue in issue_mask.items():
                if is_issue:
                    row = df.loc[idx]
                    id_val = row["ID"] if has_id else (row["TC ID"] if has_tc_id else f"Row_{idx + 1}")
                    title_val = ""
                    if has_title:
                        title_col = "Title" if "Title" in df.columns else "제목"
                        title_val = f" ({shorten_long_label(row[title_col], 12)})"
                    issue_ids.append(f"{id_val}{title_val}")

            id_list_str = ", ".join(issue_ids)
            if len(id_list_str) > 300:
                id_list_str = id_list_str[:295] + " ... (외 다수)"

            status = "🚨 기준 초과" if rate > threshold else "✅ 양호"

            missing_data.append({
                "검증 컬럼명": col,
                "미기입 및 이상 건수": f"{issue_count}건",
                "비중(%)": f"{rate:.1f}%",
                "사내 기준(%) 판정": status,
                f"미기입(공란) 해당 TC ID 목록 (전체 {issue_count}건 명단)": id_list_str
            })

    if missing_data:
        res["[필독] 컬럼별 미기입/이상치 집계 및 해당 TC ID 명단"] = pd.DataFrame(missing_data)
    else:
        res["[필독] 컬럼별 미기입/이상치 집계 및 해당 TC ID 명단"] = pd.DataFrame([{"진단 결과": f"전체 {total_len}건 중 미기입 및 이상치 0건 (모두 기준 이내로 정상)"}])

    # 2. 유사 제목(Title) 시나리오 형태소 분석 및 그룹핑
    title_col = next((c for c in df.columns if c.lower() in ("title", "제목", "테스트케이스명")), None)
    if title_col:
        titles = df[title_col].dropna().astype(str)
        groups = {}
        for t in titles:
            t_clean = re.sub(r"\[.*?\]", "", t).strip()

            # (MFC DD) 같은 접두사 추출
            prefix_match = re.match(r"^\((.*?)\)", t_clean)
            if prefix_match:
                grp = f"({prefix_match.group(1)}) 모듈 시나리오"
            else:
                words = [w for w in re.split(r"[_\s\-]+", t_clean) if len(w) >= 2 and not w.isdigit()]
                if words:
                    grp = f"[{words[0]}] 관련 시나리오"
                else:
                    grp = "기타 일반 시나리오"

            groups[grp] = groups.get(grp, 0) + 1

        grp_data = [{"시나리오 그룹명 (유사 제목 묶음)": k, "해당 건수": f"{v}건", "비중(%)": f"{(v / total_len) * 100:.1f}%"}
                    for k, v in sorted(groups.items(), key=lambda x: x[1], reverse=True)]
        res["제목(Title) 기반 주요 테스트 시나리오 그룹핑 분포"] = pd.DataFrame(grp_data)

    # 3. 주요 분류 열 TOP 5 분포
    cat_cols = [c for c in ["Domain", "Sub System", "Category", "Priority", "Test Technique", "App", "Automated"] if c in df.columns]
    if not cat_cols:
        cat_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()[:3]

    dist_data = []
    for col in cat_cols:
        if str(col).lower() in ("_polarion", "id", "tc id", "번호", "시트명", "파일명", title_col.lower() if title_col else ""):
            continue
        val_counts = df[col].value_counts()
        for val, cnt in val_counts.head(5).items():
            if str(val).strip() and str(val).strip() != "-":
                rate = (cnt / total_len) * 100
                status = "🟢 주요 타겟" if rate > threshold else "⚪ 일반"
                dist_data.append({
                    "분류 항목(컬럼)": col,
                    "세부 항목명": str(val),
                    "분포 건수": f"{cnt}건",
                    "비중(%)": f"{rate:.1f}%",
                    "사내 기준 판정": status
                })
    if dist_data:
        res["주요 카테고리별 상위 분포 현황"] = pd.DataFrame(dist_data)

    return res


def create_excel_chart(df: pd.DataFrame, title: str) -> bytes:
    if len(df) == 0:
        return None
    num_cols = df.select_dtypes(include=["number"]).columns.tolist()
    cat_cols = df.select_dtypes(include=["object", "category"]).columns.tolist()

    fig = None
    if is_manufacturing_df(df) and "라인" in cat_cols and "불량수" in num_cols:
        fig = px.bar(
            df.groupby("라인", as_index=False)["불량수"].sum(),
            x="라인", y="불량수", text_auto=True,
            title=title, color_discrete_sequence=DISTINCT_PALETTE
        )
    elif cat_cols:
        x_col = cat_cols[0]
        agg_df = df.groupby(x_col, as_index=False).size().sort_values("size", ascending=False).head(10)
        agg_df.columns = [x_col, "건수"]
        fig = px.bar(
            agg_df, x=x_col, y="건수", text_auto=True,
            title=title, color_discrete_sequence=DISTINCT_PALETTE
        )

    if fig:
        fig.update_layout(width=700, height=400)
        return pio.to_image(fig, format="png", scale=2)
    return None


def export_analysis_to_excel(
        target_df: pd.DataFrame,
        summary_df: pd.DataFrame,
        report: str,
        elapsed: float,
        model_name: str,
        aux_sheets: dict = None,
        threshold: float = DEFAULT_DEFECT_THRESHOLD,
        analyzed_at: str = None
) -> bytes:
    output = io.BytesIO()
    analyzed_at_str = analyzed_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    verdict = evaluate_threshold_verdict(target_df, threshold)

    meta_df = pd.DataFrame([
        {"항목": "분석 일시", "내용": analyzed_at_str},
        {"항목": "AI 모델명", "내용": MODEL_LABELS.get(model_name, model_name)},
        {"항목": "적용 사내 기준값(임계치)", "내용": f"{threshold}%"},
        {"항목": "사내 기준 자동 판정", "내용": verdict},
        {"항목": "분석 소요 시간", "내용": f"{elapsed} 초"},
        {"항목": "분석 데이터 건수", "내용": f"{len(target_df):,} 건"},
    ])

    parsed_sections = parse_report_into_sections(report)
    report_df = pd.DataFrame(parsed_sections)

    # [수정] 2. 기초통계집계 시트 영어 번역
    korean_summary_df = summary_df.copy()
    korean_idx_map = {
        "count": "데이터 건수",
        "unique": "고유값 종류 수",
        "top": "최다 빈도값(최빈값)",
        "freq": "최빈값 등장 횟수",
        "mean": "평균",
        "std": "표준편차",
        "min": "최소값",
        "25%": "하위 25%",
        "50%": "중앙값(50%)",
        "75%": "상위 75%",
        "max": "최대값"
    }
    korean_summary_df.index = [korean_idx_map.get(idx, idx) for idx in korean_summary_df.index]

    kor_tables = build_korean_stat_tables(target_df, aux_sheets, threshold)
    img_bytes = create_excel_chart(target_df, f"[{MODEL_LABELS.get(model_name, model_name)}] 주요 항목 건수 분포")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        meta_df.to_excel(writer, sheet_name="1. 리포트 메타정보", index=False)
        report_df.to_excel(writer, sheet_name="1. 리포트 본문", index=False)
        korean_summary_df.to_excel(writer, sheet_name="2. 기초통계집계", index=True)

        # 3번 시트: 한국어 정밀 집계표 (여러 개를 아래로 이어서 붙임)
        start_row = 0
        for t_title, t_df in kor_tables.items():
            pd.DataFrame([[f"■ {t_title}"]]).to_excel(writer, sheet_name="3. 한글 정밀통계표", index=False, header=False, startrow=start_row)
            t_df.to_excel(writer, sheet_name="3. 한글 정밀통계표", index=False, startrow=start_row + 1)
            start_row += len(t_df) + 3

        target_df.to_excel(writer, sheet_name="4. 분석대상 원본데이터", index=False)

        if aux_sheets:
            for s_name, s_df in aux_sheets.items():
                clean_name = shorten_long_label(f"5. 원본_{s_name}", 30)
                s_df.to_excel(writer, sheet_name=clean_name, index=False)

        if img_bytes:
            workbook = writer.book
            ws_chart = workbook.create_sheet("6. 차트 시각화")
            img = ExcelImage(io.BytesIO(img_bytes))
            ws_chart.add_image(img, "B2")

    return output.getvalue()


def export_combined_analysis_to_excel(
        target_df: pd.DataFrame,
        results_dict: dict,
        model_labels: dict,
        aux_sheets: dict = None,
        threshold: float = DEFAULT_DEFECT_THRESHOLD
) -> bytes:
    output = io.BytesIO()

    verdict = evaluate_threshold_verdict(target_df, threshold)

    first_res = next(iter(results_dict.values()))
    analyzed_at_str = first_res.get("analyzed_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")

    meta_rows = [
        {"비교 항목": "분석 일시", "내용": analyzed_at_str},
        {"비교 항목": "분석 데이터 건수", "내용": f"{len(target_df):,} 건"},
        {"비교 항목": "적용 사내 기준값(임계치)", "내용": f"{threshold}%"},
        {"비교 항목": "사내 기준 자동 판정", "내용": verdict},
    ]
    for m_code, r_data in results_dict.items():
        meta_rows.append({"비교 항목": f"[{model_labels.get(m_code, m_code)}] 소요 시간", "내용": f"{r_data['elapsed_time']}초"})

    meta_df = pd.DataFrame(meta_rows)

    combined_reports = []
    for m_code, r_data in results_dict.items():
        m_label = model_labels.get(m_code, m_code)
        secs = parse_report_into_sections(r_data["report"])
        for s in secs:
            combined_reports.append({
                "AI 모델": m_label,
                "구분 항목": s["구분 항목"],
                "상세 분석 내용": s["상세 분석 내용"]
            })
    combined_rep_df = pd.DataFrame(combined_reports)

    kor_tables = build_korean_stat_tables(target_df, aux_sheets, threshold)
    img_bytes = create_excel_chart(target_df, "전체 모델 공통 주요 항목 건수 분포")

    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        meta_df.to_excel(writer, sheet_name="1. 다중모델 종합비교", index=False)
        combined_rep_df.to_excel(writer, sheet_name="1. 리포트 본문 통합", index=False)

        start_row = 0
        for t_title, t_df in kor_tables.items():
            pd.DataFrame([[f"■ {t_title}"]]).to_excel(writer, sheet_name="2. 시트통합요약", index=False, header=False, startrow=start_row)
            t_df.to_excel(writer, sheet_name="2. 시트통합요약", index=False, startrow=start_row + 1)
            start_row += len(t_df) + 3

        target_df.to_excel(writer, sheet_name="3. 분석대상 원본데이터", index=False)

        if img_bytes:
            workbook = writer.book
            ws_chart = workbook.create_sheet("4. 통합 차트 시각화")
            img = ExcelImage(io.BytesIO(img_bytes))
            ws_chart.add_image(img, "B2")

    return output.getvalue()


def _set_cell_bg(cell, hex_color):
    shading = parse_xml(f'<w:shd {nsdecls("w")} w:fill="{hex_color}"/>')
    cell._tc.get_or_add_tcPr().append(shading)


def _set_cell_margins(cell, top=90, bottom=90, left=130, right=130):
    tcPr = cell._tc.get_or_add_tcPr()
    tcMar = OxmlElement('w:tcMar')
    for m, val in [('top', top), ('bottom', bottom), ('left', left), ('right', right)]:
        node = OxmlElement(f'w:{m}')
        node.set(qn('w:w'), str(val))
        node.set(qn('w:type'), 'dxa')
        tcMar.append(node)
    tcPr.append(tcMar)


def _set_table_borders(table, border_color="CBD5E1", header_color="059669"):
    tblPr = table._tbl.tblPr
    borders = parse_xml(f'''
        <w:tblBorders {nsdecls("w")}>
            <w:top w:val="single" w:sz="10" w:space="0" w:color="{header_color}"/>
            <w:bottom w:val="single" w:sz="10" w:space="0" w:color="{header_color}"/>
            <w:insideH w:val="single" w:sz="4" w:space="0" w:color="{border_color}"/>
            <w:insideV w:val="single" w:sz="4" w:space="0" w:color="E2E8F0"/>
            <w:left w:val="none"/>
            <w:right w:val="none"/>
        </w:tblBorders>
    ''')
    tblPr.append(borders)


def _add_styled_table(doc: Document, headers: list[str], rows: list[list[str]], col_widths_in: list[float]):
    tbl = doc.add_table(rows=1 + len(rows), cols=len(headers))
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    tbl.autofit = False
    _set_table_borders(tbl)

    hdr_row = tbl.rows[0]
    trPr = hdr_row._tr.get_or_add_trPr()
    trPr.append(parse_xml(f'<w:tblHeader {nsdecls("w")}/>'))
    trPr.append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))

    for idx, h_text in enumerate(headers):
        cell = hdr_row.cells[idx]
        cell.width = Inches(col_widths_in[idx])
        cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
        _set_cell_bg(cell, "059669")
        _set_cell_margins(cell, top=85, bottom=85, left=100, right=100)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(h_text)
        r.font.name = '맑은 고딕'
        r._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
        r.font.size = Pt(9.5)
        r.font.bold = True
        r.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)

    for r_idx, row_vals in enumerate(rows):
        row = tbl.rows[r_idx + 1]
        row._tr.get_or_add_trPr().append(parse_xml(f'<w:cantSplit {nsdecls("w")}/>'))
        bg_row = "FFFFFF" if r_idx % 2 == 0 else "F8FAFC"
        for c_idx, val in enumerate(row_vals):
            cell = row.cells[c_idx]
            cell.width = Inches(col_widths_in[c_idx])
            cell.vertical_alignment = WD_ALIGN_VERTICAL.CENTER
            bg = "F0FDF4" if c_idx == 0 else bg_row
            _set_cell_bg(cell, bg)
            _set_cell_margins(cell, top=75, bottom=75, left=100, right=100)
            p = cell.paragraphs[0]
            p.paragraph_format.line_spacing = 1.2
            if c_idx == 0 or (c_idx == len(headers) - 1 and len(str(val)) < 15):
                p.alignment = WD_ALIGN_PARAGRAPH.CENTER
            r = p.add_run(str(val))
            r.font.name = '맑은 고딕'
            r._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')
            r.font.size = Pt(9)
            if c_idx == 0:
                r.font.bold = True
                r.font.color.rgb = RGBColor(0x06, 0x5F, 0x46)
            else:
                r.font.color.rgb = RGBColor(0x1F, 0x29, 0x37)


def export_analysis_to_docx(
        target_df: pd.DataFrame,
        report: str,
        elapsed: float,
        model_name: str,
        threshold: float = DEFAULT_DEFECT_THRESHOLD,
        analyzed_at: str = None
) -> bytes:
    doc = Document()

    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    style_normal = doc.styles['Normal']
    style_normal.font.name = '맑은 고딕'
    style_normal.font.size = Pt(10)
    style_normal._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

    analyzed_at_str = analyzed_at or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    verdict = evaluate_threshold_verdict(target_df, threshold)

    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.77)
    _set_cell_bg(cell, "064E3B")
    _set_cell_margins(cell, top=160, bottom=160, left=200, right=200)

    p1 = cell.paragraphs[0]
    r1 = p1.add_run(f"AI 데이터 분석 리포트  |  {model_name}")
    r1.font.color.rgb = RGBColor(0xA7, 0xF3, 0xD0)
    r1.font.bold = True
    r1.font.size = Pt(9.5)

    p2 = cell.add_paragraph()
    r2 = p2.add_run("사내 데이터 심층 분석 및 실무 조치 제안서")
    r2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r2.font.bold = True
    r2.font.size = Pt(16.5)

    p3 = cell.add_paragraph()
    r3 = p3.add_run(f"분석 일시: {analyzed_at_str}  |  소요 시간: {elapsed}초  |  데이터 건수: {len(target_df):,}건")
    r3.font.color.rgb = RGBColor(0xD1, 0xFA, 0xE5)
    r3.font.size = Pt(9)

    doc.add_paragraph()

    _add_styled_table(
        doc,
        ["평가 항목", "사내 기준 자동 판정 결과 (요약)"],
        [
            ["사내 관리 기준값", f"{threshold}%"],
            ["데이터 품질 진단", verdict],
        ],
        [1.5, 5.27]
    )

    parsed_sections = parse_report_into_sections(report)
    for sec in parsed_sections:
        p_head = doc.add_paragraph()
        p_head.paragraph_format.space_before = Pt(16)
        p_head.paragraph_format.space_after = Pt(6)
        p_head.paragraph_format.keep_with_next = True
        pPr = p_head._p.get_or_add_pPr()
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="12" w:space="4" w:color="059669"/></w:pBdr>')
        pPr.append(pBdr)

        r_head = p_head.add_run(sec["구분 항목"])
        r_head.font.bold = True
        r_head.font.size = Pt(12)
        r_head.font.color.rgb = RGBColor(0x04, 0x78, 0x57)

        lines = sec["상세 분석 내용"].splitlines()
        for line in lines:
            line = line.strip()
            if not line:
                continue

            p_body = doc.add_paragraph()
            p_body.paragraph_format.space_before = Pt(2)
            p_body.paragraph_format.space_after = Pt(4)
            p_body.paragraph_format.line_spacing = 1.3

            is_bullet = False
            if line.startswith("-") or line.startswith("*"):
                is_bullet = True
                line = line[1:].strip()
                p_body.paragraph_format.left_indent = Inches(0.2)

                r_bul = p_body.add_run("• ")
                r_bul.font.color.rgb = RGBColor(0x05, 0x96, 0x69)
                r_bul.font.bold = True

            bold_parts = re.split(r'(\*\*.*?\*\*)', line)
            for part in bold_parts:
                if part.startswith("**") and part.endswith("**"):
                    r_text = p_body.add_run(part[2:-2])
                    r_text.font.bold = True
                    r_text.font.color.rgb = RGBColor(0x11, 0x18, 0x27)
                else:
                    r_text = p_body.add_run(part)
                    r_text.font.color.rgb = RGBColor(0x37, 0x41, 0x51)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()


def export_combined_analysis_to_docx(
        target_df: pd.DataFrame,
        results_dict: dict,
        model_labels: dict,
        threshold: float = DEFAULT_DEFECT_THRESHOLD
) -> bytes:
    doc = Document()

    for section in doc.sections:
        section.page_width = Inches(8.27)
        section.page_height = Inches(11.69)
        section.top_margin = Inches(0.75)
        section.bottom_margin = Inches(0.75)
        section.left_margin = Inches(0.75)
        section.right_margin = Inches(0.75)

    style_normal = doc.styles['Normal']
    style_normal.font.name = '맑은 고딕'
    style_normal.font.size = Pt(10)
    style_normal._element.rPr.rFonts.set(qn('w:eastAsia'), '맑은 고딕')

    first_res = next(iter(results_dict.values()))
    analyzed_at_str = first_res.get("analyzed_at") or datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    verdict = evaluate_threshold_verdict(target_df, threshold)

    tbl = doc.add_table(rows=1, cols=1)
    tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = tbl.cell(0, 0)
    cell.width = Inches(6.77)
    _set_cell_bg(cell, "064E3B")
    _set_cell_margins(cell, top=160, bottom=160, left=200, right=200)

    p1 = cell.paragraphs[0]
    r1 = p1.add_run("전체 AI 모델 종합 비교 리포트")
    r1.font.color.rgb = RGBColor(0xA7, 0xF3, 0xD0)
    r1.font.bold = True
    r1.font.size = Pt(9.5)

    p2 = cell.add_paragraph()
    r2 = p2.add_run("사내 데이터 다중 AI 동시 분석 결과서")
    r2.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
    r2.font.bold = True
    r2.font.size = Pt(16.5)

    p3 = cell.add_paragraph()
    r3 = p3.add_run(f"분석 일시: {analyzed_at_str}  |  대상 데이터: 총 {len(target_df):,}건")
    r3.font.color.rgb = RGBColor(0xD1, 0xFA, 0xE5)
    r3.font.size = Pt(9)

    doc.add_paragraph()

    summary_rows = [
        ["적용 사내 기준", f"{threshold}%"],
        ["자동 품질 판정", verdict],
    ]
    for m_code, r_data in results_dict.items():
        summary_rows.append([f"[{model_labels.get(m_code, m_code)}] 분석 시간", f"{r_data['elapsed_time']}초"])

    _add_styled_table(
        doc,
        ["비교 지표", "결과 요약"],
        summary_rows,
        [2.0, 4.77]
    )

    p_info = doc.add_paragraph()
    p_info.paragraph_format.space_before = Pt(10)
    r_info = p_info.add_run("※ 2페이지부터 각 AI 모델별 상세 분석 리포트가 이어집니다.")
    r_info.font.color.rgb = RGBColor(0x6B, 0x72, 0x80)
    r_info.font.size = Pt(9)

    for idx, (m_code, r_data) in enumerate(results_dict.items()):
        p_break = doc.add_paragraph()
        p_break.add_run().add_break()
        p_break.paragraph_format.page_break_before = True

        m_label = model_labels.get(m_code, m_code)

        p_mhead = doc.add_paragraph()
        p_mhead.paragraph_format.space_before = Pt(0)
        p_mhead.paragraph_format.space_after = Pt(10)
        pPr = p_mhead._p.get_or_add_pPr()
        pBdr = parse_xml(f'<w:pBdr {nsdecls("w")}><w:bottom w:val="single" w:sz="18" w:space="6" w:color="059669"/></w:pBdr>')
        pPr.append(pBdr)

        r_mhead = p_mhead.add_run(f"■ {idx + 1}. {m_label} 상세 리포트")
        r_mhead.font.bold = True
        r_mhead.font.size = Pt(14)
        r_mhead.font.color.rgb = RGBColor(0x06, 0x4E, 0x3B)

        parsed_sections = parse_report_into_sections(r_data["report"])
        for sec in parsed_sections:
            p_head = doc.add_paragraph()
            p_head.paragraph_format.space_before = Pt(12)
            p_head.paragraph_format.space_after = Pt(4)
            p_head.paragraph_format.keep_with_next = True

            r_head = p_head.add_run(sec["구분 항목"])
            r_head.font.bold = True
            r_head.font.size = Pt(11)
            r_head.font.color.rgb = RGBColor(0x04, 0x78, 0x57)

            lines = sec["상세 분석 내용"].splitlines()
            for line in lines:
                line = line.strip()
                if not line:
                    continue

                p_body = doc.add_paragraph()
                p_body.paragraph_format.space_before = Pt(1)
                p_body.paragraph_format.space_after = Pt(3)
                p_body.paragraph_format.line_spacing = 1.2

                is_bullet = False
                if line.startswith("-") or line.startswith("*"):
                    is_bullet = True
                    line = line[1:].strip()
                    p_body.paragraph_format.left_indent = Inches(0.2)

                    r_bul = p_body.add_run("• ")
                    r_bul.font.color.rgb = RGBColor(0x05, 0x96, 0x69)
                    r_bul.font.bold = True

                bold_parts = re.split(r'(\*\*.*?\*\*)', line)
                for part in bold_parts:
                    if part.startswith("**") and part.endswith("**"):
                        r_text = p_body.add_run(part[2:-2])
                        r_text.font.bold = True
                        r_text.font.color.rgb = RGBColor(0x11, 0x18, 0x27)
                    else:
                        r_text = p_body.add_run(part)
                        r_text.font.color.rgb = RGBColor(0x37, 0x41, 0x51)

    buf = io.BytesIO()
    doc.save(buf)
    return buf.getvalue()