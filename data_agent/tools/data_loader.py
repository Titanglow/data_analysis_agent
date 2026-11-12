import io
import re
import sqlite3
import pandas as pd
import openpyxl
from pathlib import Path
from datetime import datetime

DB_PATH = Path("internal_analysis.db")
TABLE_NAME = "internal_data"

# 요약·환경·설명 등 보조 시트 판별용 키워드
AUX_SHEET_KEYWORDS = [
    "summary", "요약", "environment", "환경", "cover", "표지", "개요",
    "설명", "info", "guide", "가이드", "목차", "history", "이력", "변경이력", "coverage"
]


def _classify_rgb_color(rgb_hex: str) -> str | None:
    """셀 배경색을 선입견(불합격 단정 등) 없이 순수 시각적 강조 태그로 반환합니다."""
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

    # 흰색·연회색 배경 제외
    if r > 240 and g > 240 and b > 240:
        return None
    if abs(r - g) < 12 and abs(g - b) < 12 and r > 210:
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
                # 1행(헤더)은 컬럼명 매칭을 위해 색상 태그를 붙이지 않음
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

    # 가장 유효 열이 많은 초반 행을 헤더로 선택 (요약 시트 등 상단 제목 행 대응)
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


def _is_aux_sheet(sheet_name: str, df: pd.DataFrame) -> bool:
    """시트 이름과 컬럼 구조를 분석하여 Summary/Environment 같은 보조·요약 시트인지 판별합니다."""
    s_lower = sheet_name.lower()
    if any(k in s_lower for k in AUX_SHEET_KEYWORDS):
        return True
    # 컬럼 수가 3개 이하이면서 행 수가 적은 키-값 형태 시트도 보조 시트로 분류
    if len(df.columns) <= 3 and len(df) <= 30:
        return True
    return False


def load_and_sync_to_db(
    uploaded_file, selected_sheet: str = "AUTO_MAIN"
) -> tuple[pd.DataFrame, list[str], dict[str, int], dict[str, pd.DataFrame]]:
    """
    파일을 읽어 핵심 분석 테이블(DB 적재용)과 보조/요약 시트(Summary, Environment 등)를 분리해 반환합니다.
    반환값: (분석 대상 DataFrame, 전체 시트 목록, 시트별 건수, 보조 시트 DataFrame 딕셔너리)
    """
    uploaded_file.seek(0)
    file_bytes = uploaded_file.read()

    sheet_names = ["단일시트(CSV)"]
    sheet_counts = {}
    aux_sheets: dict[str, pd.DataFrame] = {}

    if uploaded_file.name.lower().endswith(".csv"):
        try:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="utf-8")
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(file_bytes), encoding="cp949")
        for col in df.select_dtypes(include=["object"]).columns:
            df[col] = df[col].fillna("-").astype(str).replace({"None": "-", "nan": "-"})
        sheet_counts["단일시트(CSV)"] = len(df)
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

        # 핵심 데이터 시트 vs 보조(Summary/Environment) 시트 분류
        for s_name, s_df in all_sheet_dfs.items():
            if _is_aux_sheet(s_name, s_df) and len(all_sheet_dfs) > 1:
                aux_sheets[s_name] = s_df
            else:
                df_copy = s_df.copy()
                if len(all_sheet_dfs) > 1:
                    df_copy.insert(0, "시트명", s_name)
                main_dfs.append(df_copy)

        # 만약 모든 시트가 보조 시트로 분류되었다면 첫 번째 시트를 메인으로 사용
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