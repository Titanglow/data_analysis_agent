# data_agent/tools/data_loader.py
import io
import re
import sqlite3
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from datetime import datetime
import pandas as pd
import openpyxl

DB_PATH = Path("internal_analysis.db")
TABLE_NAME = "internal_data"


def _safe_str(val) -> str:
    """None, NaN, float 등 어떤 타입이 들어와도 안전하게 문자열로 변환합니다."""
    if val is None:
        return "-"
    try:
        if pd.isna(val):
            return "-"
    except Exception:
        pass
    s = str(val).strip()
    if s.lower() in ("", "none", "nan", "null", "<na>"):
        return "-"
    return s


def _strip_color_tag(val) -> str:
    """셀 값에서 [🎨...] 또는 [🔴...] 색상 태그만 안전하게 제거합니다 (float 에러 원천 방지)."""
    s = _safe_str(val)
    if s == "-":
        return "-"
    cleaned = re.sub(r"\s*\[(?:🎨|🔴|🟠|🟡|🟢|🔵|🟣|빈칸)[^\]]*\]", "", s).strip()
    return cleaned if cleaned else "-"


def _classify_rgb_color(rgb_hex: str) -> str | None:
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
    if r > 180 and g < 130 and b < 130:
        return "🎨빨강강조"
    if r > 200 and 100 <= g < 185 and b < 110:
        return "🎨주황강조"
    if r > 200 and g >= 185 and b < 150:
        return "🎨노랑강조"
    if g > 160 and r < 160 and b < 160:
        return "🎨초록강조"
    if b > 170 and r < 160 and g < 185:
        return "🎨파랑강조"
    if r > 150 and b > 150 and g < 140:
        return "🎨보라강조"
    return None


def _is_aux_sheet(sheet_name: str, df: pd.DataFrame) -> bool:
    s_low = sheet_name.strip().lower()
    aux_keywords = ["summary", "env", "environment", "cover", "history", "revision", "info", "요약", "환경", "표지", "이력", "개요"]
    if any(k in s_low for k in aux_keywords):
        return True
    if len(df) <= 3 and len(df.columns) <= 3:
        return True
    return False


def _read_excel_sheet_with_colors(wb: openpyxl.Workbook, sheet_name: str) -> pd.DataFrame:
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
                    elif getattr(fg, "theme", None) is not None:
                        if fg.theme in (4, 5):
                            color_tag = "🎨빨강강조"

            if val_str and color_tag:
                final_val = val_str if cell.row == 1 else f"{val_str} [{color_tag}]"
                has_any_content = True
            elif not val_str and color_tag:
                final_val = f"[빈칸·{color_tag}]"
                has_any_content = True
            elif val_str:
                final_val = raw_val if isinstance(raw_val, (int, float)) and not pd.isna(raw_val) else val_str
                has_any_content = True
            else:
                final_val = None

            row_vals.append(final_val)

        if has_any_content:
            rows_data.append(row_vals)

    if not rows_data:
        return pd.DataFrame()

    # 상단 10행 이내에서 가장 유효 컬럼이 많은 행을 헤더로 자동 선택 (배민 TC 등 상단 여백 대응)
    header_idx = 0
    best_non_empty = 0
    for idx in range(min(10, len(rows_data))):
        non_empty = sum(1 for v in rows_data[idx] if v is not None and not str(v).startswith("[빈칸"))
        if non_empty > best_non_empty:
            best_non_empty = non_empty
            header_idx = idx

    raw_headers = rows_data[header_idx]
    headers = []
    seen = {}
    for idx, h in enumerate(raw_headers):
        h_clean = _strip_color_tag(h)
        h_name = h_clean if h_clean != "-" else f"컬럼_{idx + 1}"
        if h_name in seen:
            seen[h_name] += 1
            h_name = f"{h_name}_{seen[h_name]}"
        else:
            seen[h_name] = 1
        headers.append(h_name)

    data_rows = rows_data[header_idx + 1 :]
    if not data_rows:
        return pd.DataFrame(columns=headers)

    df = pd.DataFrame(data_rows, columns=headers)
    df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)

    for col in df.columns:
        safe_list = [_safe_str(v) for v in df[col].tolist()]
        safe_series = pd.Series(safe_list, index=df.index, dtype=object)
        has_color = any(("🎨" in s or "🔴" in s or "[빈칸" in s) for s in safe_list)

        if not has_color:
            stripped_list = [_strip_color_tag(s) for s in safe_list]
            non_dash = [s for s in stripped_list if s != "-"]
            if len(non_dash) > 0 and len(non_dash) == len(stripped_list):
                converted = pd.to_numeric(pd.Series(stripped_list, index=df.index), errors="coerce")
                if converted.notna().all():
                    df[col] = converted
                    continue

        df[col] = safe_series

    return df


def _read_docx_file(file_bytes: bytes) -> tuple[pd.DataFrame, list[str], dict[str, int], dict[str, pd.DataFrame]]:
    paragraphs = []
    tables_dfs = []
    ns = {"w": "http://schemas.openxmlformats.org/wordprocessingml/2006/main"}

    try:
        with zipfile.ZipFile(io.BytesIO(file_bytes)) as zf:
            if "word/document.xml" in zf.namelist():
                xml_content = zf.read("word/document.xml")
                root = ET.fromstring(xml_content)

                for idx, p in enumerate(root.findall(".//w:body/w:p", ns), start=1):
                    texts = [t.text for t in p.findall(".//w:t", ns) if t.text]
                    full_txt = "".join(texts).strip()
                    if full_txt:
                        p_type = "제목/헤더" if len(full_txt) < 40 and (full_txt.startswith(("[", "■", "1.", "2.", "3.", "4.")) or ":" in full_txt) else "본문 문단"
                        paragraphs.append({"순번": idx, "구분": p_type, "내용": full_txt})

                for t_idx, tbl in enumerate(root.findall(".//w:body/w:tbl", ns), start=1):
                    t_rows = []
                    for tr in tbl.findall(".//w:tr", ns):
                        cells = []
                        for tc in tr.findall(".//w:tc", ns):
                            c_texts = [t.text for t in tc.findall(".//w:t", ns) if t.text]
                            cells.append(" ".join("".join(c_texts).split()) or "-")
                        if any(c != "-" for c in cells):
                            t_rows.append(cells)
                    if len(t_rows) >= 2:
                        headers = [f"{h}_{i}" if t_rows[0].count(h) > 1 else h for i, h in enumerate(t_rows[0])]
                        t_df = pd.DataFrame(t_rows[1:], columns=headers)
                        tables_dfs.append((f"워드표_{t_idx}", t_df))
    except Exception:
        pass

    main_df = pd.DataFrame(paragraphs) if paragraphs else pd.DataFrame([{"순번": 1, "구분": "알림", "내용": "추출된 텍스트가 없습니다."}])
    sheet_names = ["본문_문단"]
    sheet_counts = {"본문_문단": len(main_df)}
    aux_sheets = {}

    for t_name, t_df in tables_dfs:
        sheet_names.append(t_name)
        sheet_counts[t_name] = len(t_df)
        aux_sheets[t_name] = t_df

    return main_df, sheet_names, sheet_counts, aux_sheets


def load_and_sync_to_db(
    uploaded_file, selected_sheet: str = "ALL"
) -> tuple[pd.DataFrame, list[str], dict[str, int], dict[str, pd.DataFrame]]:
    uploaded_file.seek(0)
    file_bytes = uploaded_file.read()
    fname = uploaded_file.name.lower()

    aux_sheets: dict[str, pd.DataFrame] = {}

    if fname.endswith(".csv") or fname.endswith(".txt"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="cp949")
        except Exception:
            text_lines = file_bytes.decode("utf-8", errors="replace").splitlines()
            df = pd.DataFrame([{"행번호": i + 1, "내용": line.strip()} for i, line in enumerate(text_lines) if line.strip()])

        for col in df.columns:
            if not pd.api.types.is_numeric_dtype(df[col]):
                df[col] = pd.Series([_safe_str(v) for v in df[col].tolist()], index=df.index, dtype=object)
        sheet_names = ["단일시트"]
        sheet_counts = {"단일시트": len(df)}

    elif fname.endswith(".docx"):
        df, sheet_names, sheet_counts, aux_sheets = _read_docx_file(file_bytes)
        if selected_sheet in aux_sheets:
            df = aux_sheets[selected_sheet]

    else:
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
        sheet_names = wb.sheetnames
        sheet_counts = {}
        main_dfs = []
        all_sheet_dfs = {}

        for s_name in sheet_names:
            s_df = _read_excel_sheet_with_colors(wb, s_name)
            sheet_counts[s_name] = len(s_df)
            if len(s_df) > 0:
                all_sheet_dfs[s_name] = s_df
                if _is_aux_sheet(s_name, s_df):
                    aux_sheets[s_name] = s_df
                else:
                    s_copy = s_df.copy()
                    s_copy.insert(0, "시트명", s_name)
                    main_dfs.append(s_copy)

        if not main_dfs and all_sheet_dfs:
            for s_name, s_df in all_sheet_dfs.items():
                s_copy = s_df.copy()
                s_copy.insert(0, "시트명", s_name)
                main_dfs.append(s_copy)
            aux_sheets = {}

        if selected_sheet == "ALL" or selected_sheet not in all_sheet_dfs:
            df = pd.concat(main_dfs, ignore_index=True) if main_dfs else pd.DataFrame()
        else:
            df = all_sheet_dfs[selected_sheet].copy()
            if "시트명" not in df.columns:
                df.insert(0, "시트명", selected_sheet)

        for col in df.columns:
            if not pd.api.types.is_numeric_dtype(df[col]):
                df[col] = pd.Series([_safe_str(v) for v in df[col].tolist()], index=df.index, dtype=object)

    if len(df) > 0:
        with sqlite3.connect(DB_PATH) as conn:
            df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)

    return df, sheet_names, sheet_counts, aux_sheets


def filter_by_keywords(df: pd.DataFrame, keyword_str: str, match_all: bool = False) -> pd.DataFrame:
    if not keyword_str or not keyword_str.strip() or len(df) == 0:
        return df

    keywords = [k.strip() for k in keyword_str.split(",") if k.strip()]
    if not keywords:
        return df

    row_text = df.astype(str).agg(" ".join, axis=1)
    if match_all:
        mask = pd.Series(True, index=df.index)
        for kw in keywords:
            mask &= row_text.str.contains(re.escape(kw), case=False, na=False)
    else:
        pattern = "|".join(re.escape(k) for k in keywords)
        mask = row_text.str.contains(pattern, case=False, na=False)

    return df[mask]


def run_sql_query(sql_query: str) -> pd.DataFrame:
    with sqlite3.connect(DB_PATH) as conn:
        res_df = pd.read_sql_query(sql_query, conn)
    for col in res_df.columns:
        if not pd.api.types.is_numeric_dtype(res_df[col]):
            res_df[col] = pd.Series([_safe_str(v) for v in res_df[col].tolist()], index=res_df.index, dtype=object)
    return res_df


def export_db_to_sql_script(current_sql: str = "") -> bytes:
    lines = [
        "-- ============================================================================",
        "-- [사내 데이터 분석 AI Agent] SQLite 데이터베이스 전체 덤프 (.sql)",
        f"-- 생성 일시: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}",
        f"-- 대상 테이블명: {TABLE_NAME}",
        "-- ============================================================================",
        "",
    ]

    if current_sql and current_sql.strip():
        lines.extend([
            "-- [1] 현재 대시보드에서 실행 중인 조회 SQL 쿼리",
            current_sql.strip(),
            "",
        ])

    if DB_PATH.exists():
        with sqlite3.connect(DB_PATH) as conn:
            for dump_line in conn.iterdump():
                lines.append(dump_line)

    return "\n".join(lines).encode("utf-8-sig")