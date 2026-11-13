import io
import re
import sqlite3
import zipfile
import xml.etree.ElementTree as ET
import pandas as pd
import openpyxl
from pathlib import Path
from datetime import datetime

DB_PATH = Path("internal_analysis.db")
TABLE_NAME = "internal_data"

AUX_SHEET_KEYWORDS = [
    "summary", "요약", "environment", "환경", "cover", "표지", "개요",
    "설명", "info", "guide", "가이드", "목차", "history", "이력", "변경이력", "coverage"
]


def _classify_rgb_color(rgb_hex: str) -> str | None:
    """셀 배경색을 선입견 없이 순수 시각적 강조 태그로 반환합니다."""
    if not rgb_hex or not isinstance(rgb_hex, str):
        return None
    clean_hex = rgb_hex.strip().upper()
    if clean_hex in ("00000000", "FFFFFFFF", "00FFFFFF", "FF000000", "000000", "FFFFFF"):
        return None
    if len(clean_hex) == 8:
        clean_hex = clean_hex[2:]
    if len(clean_hex) != 6:
        return None
    try:
        r = int(clean_hex[0:2], 16)
        g = int(clean_hex[2:4], 16)
        b = int(clean_hex[4:6], 16)
    except ValueError:
        return None

    if r > 240 and g > 240 and b > 240:
        return None
    if abs(r - g) < 12 and abs(r - b) < 12 and r > 210:
        return None

    if r > 180 and g < 130 and b < 130:
        return "🎨빨강강조"
    if r > 200 and 100 <= g < 185 and b < 110:
        return "🎨주황강조"
    if r > 200 and g >= 185 and b < 150:
        return "🎨노랑강조"
    if g > 160 and r < 150 and b < 160:
        return "🎨초록강조"
    if b > 170 and r < 150 and g < 185:
        return "🎨파랑강조"
    if r > 160 and b > 160 and g < 140:
        return "🎨보라강조"
    return None


def _read_excel_sheet_with_colors(wb: openpyxl.Workbook, sheet_name: str) -> pd.DataFrame:
    """시트의 셀 값과 배경색 강조 태그를 읽어 DataFrame으로 변환합니다."""
    ws = wb[sheet_name]
    rows_data = []

    for row in ws.iter_rows(values_only=False):
        row_vals = []
        has_any_content = False
        for cell in row:
            raw_val = cell.value
            val_str = "" if raw_val is None else str(raw_val).strip()
            if val_str.lower() in ("none", "nan", "null"):
                val_str = ""

            color_tag = None
            fill = getattr(cell, "fill", None)
            if fill and getattr(fill, "fill_type", None) in ("solid", "mediumGray", "darkGray", "lightGray"):
                fg = getattr(fill, "start_color", None)
                if fg:
                    rgb_val = getattr(fg, "rgb", None)
                    if isinstance(rgb_val, str):
                        color_tag = _classify_rgb_color(rgb_val)

            if val_str and color_tag:
                final_val = val_str if cell.row == 1 else f"{val_str} [{color_tag}]"
                has_any_content = True
            elif not val_str and color_tag:
                final_val = f"[빈칸·{color_tag}]"
                has_any_content = True
            elif val_str:
                final_val = raw_val if isinstance(raw_val, (int, float)) else val_str
                has_any_content = True
            else:
                final_val = None

            row_vals.append(final_val)

        if has_any_content:
            rows_data.append(row_vals)

    if not rows_data:
        return pd.DataFrame()

    header_idx = 0
    for idx in range(min(5, len(rows_data))):
        non_empty = sum(1 for x in rows_data[idx] if x is not None and not str(x).startswith("[빈칸"))
        if non_empty >= 3:
            header_idx = idx
            break

    raw_headers = rows_data[header_idx]
    headers = []
    seen = {}
    for idx, h in enumerate(raw_headers):
        h_clean = re.sub(r"\s*\[🎨[^\]]+\]", "", str(h or "")).strip()
        h_name = h_clean if h_clean else f"항목_{idx + 1}"
        if h_name in seen:
            seen[h_name] += 1
            h_name = f"{h_name}_{seen[h_name]}"
        else:
            seen[h_name] = 1
        headers.append(h_name)

    df = pd.DataFrame(rows_data[header_idx + 1 :], columns=headers)
    df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)

    for col in df.columns:
        try:
            converted = pd.to_numeric(df[col])
            if converted.notna().sum() > 0 and converted.notna().sum() == df[col].notna().sum():
                df[col] = converted
        except Exception:
            pass

    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].fillna("-").astype(str).replace({"None": "-", "nan": "-", "": "-"})

    return df


def _parse_docx_bytes(
    file_bytes: bytes, selected_sheet: str = "AUTO_MAIN"
) -> tuple[pd.DataFrame, list[str], dict[str, int], dict[str, pd.DataFrame]]:
    """
    Word(.docx) 문서에서 본문 문단(목차/섹션별 자동 분류)과 문서 내 모든 표(Table)를 추출하여
    SQLite DB 조회, 차트 시각화, AI 리포트 생성이 가능하도록 정형 테이블로 변환합니다.
    """
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}
    with zipfile.ZipFile(io.BytesIO(file_bytes)) as z:
        xml_content = z.read("word/document.xml")
    tree = ET.fromstring(xml_content)
    body = tree.find("w:body", ns)

    para_rows = []
    current_section = "1. 문서 도입부"
    para_idx = 0

    if body is not None:
        for p in body.findall("w:p", ns):
            texts = [t.text for t in p.findall(".//w:t", ns) if t.text]
            full_text = " ".join("".join(texts).split()).strip()
            if not full_text:
                continue

            pStyle = p.find(".//w:pStyle", ns)
            style_val = pStyle.attrib.get(f"{{{ns['w']}}}val", "") if pStyle is not None else ""
            is_heading = (
                "Heading" in style_val
                or "제목" in style_val
                or (
                    len(full_text) <= 50
                    and full_text.startswith(("01", "02", "03", "04", "05", "1.", "2.", "3.", "4.", "5.", "■", "▶", "I.", "II."))
                )
            )
            if is_heading and len(full_text) <= 55:
                current_section = full_text
                continue

            para_idx += 1
            if any(k in full_text for k in ["불량", "오류", "에러", "문제", "실패", "FAIL", "주의", "미흡", "지연", "초과", "위험", "개선"]):
                status_tag = "🚨 중점/이슈 사항"
            elif any(k in full_text for k in ["완료", "성공", "정상", "PASS", "달성", "양호", "확보"]):
                status_tag = "✅ 완료/양호 사항"
            else:
                status_tag = "📌 일반 설명/현황"

            para_rows.append({
                "문단ID": f"DOC-{para_idx:03d}",
                "문서 섹션(목차)": current_section[:45],
                "핵심 분류": status_tag,
                "본문 상세 내용": full_text,
                "글자수": len(full_text),
            })

    table_dfs: dict[str, pd.DataFrame] = {}
    if body is not None:
        for t_idx, tbl in enumerate(body.findall("w:tbl", ns), start=1):
            t_rows = []
            for tr in tbl.findall("w:tr", ns):
                row_cells = []
                for tc in tr.findall("w:tc", ns):
                    cell_texts = [t.text for t in tc.findall(".//w:t", ns) if t.text]
                    row_cells.append(" ".join("".join(cell_texts).split()).strip() or "-")
                if any(c != "-" for c in row_cells):
                    t_rows.append(row_cells)
            if len(t_rows) >= 2:
                max_cols = max(len(r) for r in t_rows)
                norm_rows = [r + ["-"] * (max_cols - len(r)) for r in t_rows]
                seen = {}
                clean_headers = []
                for idx_h, h in enumerate(norm_rows[0]):
                    h_name = h[:30] if h and h != "-" else f"열_{idx_h + 1}"
                    if h_name in seen:
                        seen[h_name] += 1
                        h_name = f"{h_name}_{seen[h_name]}"
                    else:
                        seen[h_name] = 1
                    clean_headers.append(h_name)
                t_df = pd.DataFrame(norm_rows[1:], columns=clean_headers)
                t_name = f"워드내장표_{t_idx}"
                table_dfs[t_name] = t_df

    para_df = pd.DataFrame(para_rows)
    sheet_names = []
    sheet_counts = {}
    aux_sheets: dict[str, pd.DataFrame] = {}

    if len(para_df) > 0:
        sheet_names.append("워드_본문문단")
        sheet_counts["워드_본문문단"] = len(para_df)

    for t_name, t_df in table_dfs.items():
        sheet_names.append(t_name)
        sheet_counts[t_name] = len(t_df)
        aux_sheets[t_name] = t_df

    if selected_sheet in table_dfs:
        main_df = table_dfs[selected_sheet].copy()
        if len(para_df) > 0:
            aux_sheets["워드_본문문단"] = para_df
        aux_sheets.pop(selected_sheet, None)
    elif len(para_df) > 0:
        main_df = para_df
    elif table_dfs:
        first_tbl = next(iter(table_dfs.keys()))
        main_df = table_dfs[first_tbl].copy()
        aux_sheets.pop(first_tbl, None)
    else:
        main_df = pd.DataFrame([{"문단ID": "DOC-001", "문서 섹션(목차)": "빈 문서", "핵심 분류": "일반", "본문 상세 내용": "내용 없음", "글자수": 0}])

    return main_df, sheet_names, sheet_counts, aux_sheets


def _parse_text_bytes(
    file_bytes: bytes, file_name: str
) -> tuple[pd.DataFrame, list[str], dict[str, int], dict[str, pd.DataFrame]]:
    """텍스트(.txt) 및 마크다운(.md) 문서를 섹션별·문단별 정형 테이블로 변환합니다."""
    try:
        text = file_bytes.decode("utf-8")
    except UnicodeDecodeError:
        text = file_bytes.decode("cp949", errors="replace")

    rows = []
    current_sec = "문서 시작부"
    idx = 0
    for raw_line in text.splitlines():
        line = raw_line.strip()
        if not line or re.match(r"^[=\-*_#\s]+$", line):
            continue
        if line.startswith("#") or (len(line) <= 45 and re.match(r"^(?:\d+\.|\[|■|▶)", line)):
            current_sec = line.lstrip("# ").strip()
            continue
        idx += 1
        status_tag = "🚨 중점/이슈 사항" if any(k in line for k in ["불량", "오류", "에러", "문제", "실패", "주의", "미흡", "초과"]) else "✅ 일반 내용"
        rows.append({
            "행ID": f"TXT-{idx:03d}",
            "문서 섹션(목차)": current_sec[:45],
            "핵심 분류": status_tag,
            "본문 상세 내용": line,
            "글자수": len(line),
        })

    df = pd.DataFrame(rows) if rows else pd.DataFrame([{"행ID": "TXT-001", "문서 섹션(목차)": "기본", "핵심 분류": "일반", "본문 상세 내용": "내용 없음", "글자수": 0}])
    s_name = f"텍스트문서({Path(file_name).suffix})"
    return df, [s_name], {s_name: len(df)}, {}


def _is_aux_sheet(sheet_name: str, df: pd.DataFrame) -> bool:
    s_lower = sheet_name.lower()
    if any(k in s_lower for k in AUX_SHEET_KEYWORDS):
        return True
    if len(df.columns) <= 3 and len(df) <= 30:
        return True
    return False


def load_and_sync_to_db(
    uploaded_file, selected_sheet: str = "AUTO_MAIN"
) -> tuple[pd.DataFrame, list[str], dict[str, int], dict[str, pd.DataFrame]]:
    """
    CSV, 다중 시트 Excel(.xlsx), Word 문서(.docx), 텍스트/마크다운(.txt, .md) 파일을 모두 지원하며,
    즉시 로컬 SQLite DB(internal_analysis.db) 테이블로 동기화합니다.
    """
    uploaded_file.seek(0)
    file_bytes = uploaded_file.read()
    fname_lower = uploaded_file.name.lower()

    sheet_names = ["단일시트(CSV)"]
    sheet_counts = {}
    aux_sheets: dict[str, pd.DataFrame] = {}

    if fname_lower.endswith(".csv"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="cp949")
        for col in df.select_dtypes(include=["object"]).columns:
            df[col] = df[col].fillna("-").astype(str).replace({"None": "-", "nan": "-"})
        sheet_counts["단일시트(CSV)"] = len(df)

    elif fname_lower.endswith(".docx"):
        df, sheet_names, sheet_counts, aux_sheets = _parse_docx_bytes(file_bytes, selected_sheet=selected_sheet)

    elif fname_lower.endswith((".txt", ".md")):
        df, sheet_names, sheet_counts, aux_sheets = _parse_text_bytes(file_bytes, uploaded_file.name)

    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet_names = wb.sheetnames
        all_sheet_dfs: dict[str, pd.DataFrame] = {}
        main_dfs = []

        for s_name in sheet_names:
            s_df = _read_excel_sheet_with_colors(wb, s_name)
            sheet_counts[s_name] = len(s_df)
            if len(s_df) > 0:
                all_sheet_dfs[s_name] = s_df

        for s_name, s_df in all_sheet_dfs.items():
            if _is_aux_sheet(s_name, s_df) and len(all_sheet_dfs) > 1:
                aux_sheets[s_name] = s_df
            else:
                df_copy = s_df.copy()
                if len(all_sheet_dfs) > 1:
                    df_copy.insert(0, "시트명", s_name)
                main_dfs.append(df_copy)

        if not main_dfs and all_sheet_dfs:
            first_s = next(iter(all_sheet_dfs.keys()))
            df_copy = all_sheet_dfs[first_s].copy()
            df_copy.insert(0, "시트명", first_s)
            main_dfs.append(df_copy)
            aux_sheets.pop(first_s, None)

        if selected_sheet == "AUTO_MAIN":
            df = pd.concat(main_dfs, ignore_index=True) if main_dfs else pd.DataFrame()
        elif selected_sheet == "ALL_MERGE":
            merged = []
            for s_name, s_df in all_sheet_dfs.items():
                c_df = s_df.copy()
                c_df.insert(0, "시트명", s_name)
                merged.append(c_df)
            df = pd.concat(merged, ignore_index=True) if merged else pd.DataFrame()
        elif selected_sheet in all_sheet_dfs:
            df = all_sheet_dfs[selected_sheet].copy()
            df.insert(0, "시트명", selected_sheet)
        else:
            df = pd.concat(main_dfs, ignore_index=True) if main_dfs else pd.DataFrame()

        for col in df.select_dtypes(include=["object"]).columns:
            df[col] = df[col].fillna("-").astype(str).replace({"None": "-", "nan": "-"})

    if len(df) > 0:
        with sqlite3.connect(DB_PATH) as conn:
            df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)

    return df, sheet_names, sheet_counts, aux_sheets


def filter_by_keywords(df: pd.DataFrame, keyword_str: str, match_all: bool = False) -> pd.DataFrame:
    if not keyword_str or not keyword_str.strip():
        return df
    keywords = [k.strip() for k in keyword_str.split(",") if k.strip()]
    if not keywords:
        return df

    row_text = df.astype(str).agg(" ".join, axis=1)
    if match_all:
        mask = pd.Series(True, index=df.index)
        for kw in keywords:
            mask &= row_text.str.contains(kw, case=False, na=False)
    else:
        pattern = "|".join(keywords)
        mask = row_text.str.contains(pattern, case=False, na=False)
    return df[mask]


def run_sql_query(sql_query: str) -> pd.DataFrame:
    with sqlite3.connect(DB_PATH) as conn:
        res_df = pd.read_sql_query(sql_query, conn)
    for col in res_df.select_dtypes(include=["object"]).columns:
        res_df[col] = res_df[col].fillna("-").astype(str).replace({"None": "-", "nan": "-"})
    return res_df


def export_db_to_sql_script(current_sql: str = "") -> bytes:
    lines = [
        "-- ============================================================================",
        "-- [사내 데이터 분석 AI Agent] SQLite 데이터베이스 전체 덤프 (.sql)",
        f"-- 생성 일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"-- 대상 테이블명: {TABLE_NAME}",
        "-- * 이 파일은 Notepad++, DBeaver, VS Code 등에서 UTF-8로 바로 열람·실행 가능합니다.",
        "-- ============================================================================",
        "",
    ]
    if current_sql and current_sql.strip():
        lines.extend([
            "-- [1] 현재 대시보드에서 실행 중인 조회 SQL 쿼리",
            current_sql.strip(),
            "",
            "-- ============================================================================",
            "-- [2] 테이블 스키마(CREATE TABLE) 및 전체 데이터(INSERT INTO) 덤프",
            "-- ============================================================================",
        ])
    if DB_PATH.exists():
        with sqlite3.connect(DB_PATH) as conn:
            for dump_line in conn.iterdump():
                lines.append(dump_line)
    return "\n".join(lines).encode("utf-8-sig")