import io
import sqlite3
import pandas as pd
import openpyxl
from pathlib import Path
from datetime import datetime

DB_PATH = Path("internal_analysis.db")
TABLE_NAME = "internal_data"


def _classify_rgb_color(rgb_hex: str) -> str | None:
    """엑셀 셀의 ARGB/RGB 16진수 색상 코드를 분석하여 빨강/초록/노랑/주황/파랑 등 의미 있는 색상 태그를 반환합니다."""
    if not rgb_hex or not isinstance(rgb_hex, str):
        return None
    clean_hex = rgb_hex.strip().upper()
    # 기본 투명/흰색/검정색(기본값) 제외
    if clean_hex in ("00000000", "FFFFFFFF", "00FFFFFF", "FF000000", "000000", "FFFFFF"):
        return None
    if len(clean_hex) == 8:
        clean_hex = clean_hex[2:]  # 앞 2자리 Alpha 채널 제거
    if len(clean_hex) != 6:
        return None
    try:
        r = int(clean_hex[0:2], 16)
        g = int(clean_hex[2:4], 16)
        b = int(clean_hex[4:6], 16)
    except ValueError:
        return None

    # 거의 흰색이거나 아주 연한 회색 테두리 배경은 제외
    if r > 242 and g > 242 and b > 242:
        return None

    # 색상 판별 로직 (빨강 / 주황 / 노랑 / 초록 / 파랑 / 보라 / 회색)
    if r > 180 and g < 120 and b < 120:
        return "🔴빨강(불합격/주의)"
    if r > 200 and g > 100 and g < 185 and b < 100:
        return "🟠주황(점검필요)"
    if r > 200 and g > 185 and b < 140:
        return "🟡노랑(보류/확인)"
    if g > 160 and r < 150 and b < 160:
        return "🟢초록(정상/PASS)"
    if b > 170 and r < 150 and g < 180:
        return "🔵파랑(참고/완료)"
    if r > 160 and b > 160 and g < 140:
        return "🟣보라(특이사항)"
    return f"🎨색상강조(#{clean_hex})"


def _read_excel_sheet_with_colors(wb: openpyxl.Workbook, sheet_name: str) -> pd.DataFrame:
    """
    openpyxl로 시트를 읽어 셀 값뿐만 아니라 배경색(예: Actual Result에 빨간색만 칠해진 칸)까지
    빠짐없이 텍스트 태그로 추출하고 None 빈칸을 깔끔하게 정리합니다.
    """
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

            # 셀 배경색 감지
            color_tag = None
            fill = getattr(cell, "fill", None)
            if fill and getattr(fill, "fill_type", None) in ("solid", "mediumGray", "darkGray", "lightGray"):
                fg = getattr(fill, "start_color", None)
                if fg:
                    rgb_val = getattr(fg, "rgb", None)
                    if isinstance(rgb_val, str):
                        color_tag = _classify_rgb_color(rgb_val)
                    elif getattr(fg, "theme", None) is not None:
                        # 엑셀 테마 색상 인덱스 매핑
                        theme_idx = fg.theme
                        if theme_idx in (4, 5):
                            color_tag = "🔴테마강조(주의)"
                        elif theme_idx in (6, 7, 8, 9):
                            color_tag = "🎨테마배경색"

            if val_str and color_tag:
                # 헤더 행(1행)은 컬럼명이 깨지지 않도록 원래 이름 유지
                if cell.row == 1:
                    final_val = val_str
                else:
                    final_val = f"{val_str} [{color_tag}]"
                has_any_content = True
            elif not val_str and color_tag:
                # 글자 없이 색깔만 칠해둔 셀 (예: Actual Result에 빨간색만 칠한 칸)
                final_val = f"[{color_tag}]"
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

    # 첫 번째 유효 행을 컬럼 헤더로 지정
    raw_headers = rows_data[0]
    headers = []
    seen = {}
    for idx, h in enumerate(raw_headers):
        h_name = str(h).strip() if h is not None and str(h).strip() != "" else f"컬럼_{idx + 1}"
        if h_name in seen:
            seen[h_name] += 1
            h_name = f"{h_name}_{seen[h_name]}"
        else:
            seen[h_name] = 1
        headers.append(h_name)

    df = pd.DataFrame(rows_data[1:], columns=headers)

    # 모든 값이 비어 있는 열 제거 및 숫자로 변환 가능한 열은 수치형 변환
    df = df.dropna(how="all", axis=1).dropna(how="all", axis=0)
    for col in df.columns:
        try:
            converted = pd.to_numeric(df[col])
            if converted.notna().sum() > 0 and converted.notna().sum() == df[col].notna().sum():
                df[col] = converted
        except Exception:
            pass

    # 텍스트 컬럼의 None / NaN 빈칸을 보기 좋은 '-' 로 치환
    for col in df.select_dtypes(include=["object"]).columns:
        df[col] = df[col].fillna("-").astype(str).replace({"None": "-", "nan": "-", "": "-"})

    return df


def load_and_sync_to_db(uploaded_file, selected_sheet: str = "ALL") -> tuple[pd.DataFrame, list[str], dict[str, int]]:
    """
    CSV 및 다중 시트 Excel(.xlsx) 파일을 읽고, 셀 배경색과 모든 시트 데이터를 추출하여
    로컬 SQLite DB(internal_analysis.db)에 자동 동기화합니다.
    반환값: (조회된 DataFrame, 전체 시트 이름 목록, 시트별 행 건수 딕셔너리)
    """
    uploaded_file.seek(0)
    file_bytes = uploaded_file.read()

    sheet_names = ["단일시트(CSV)"]
    sheet_counts = {}

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
        sheet_dfs = []

        for s_name in sheet_names:
            s_df = _read_excel_sheet_with_colors(wb, s_name)
            sheet_counts[s_name] = len(s_df)
            if len(s_df) > 0:
                s_df.insert(0, "시트명", s_name)
                sheet_dfs.append(s_df)

        if not sheet_dfs:
            df = pd.DataFrame()
        elif selected_sheet == "ALL" or selected_sheet not in sheet_names:
            df = pd.concat(sheet_dfs, ignore_index=True)
            for col in df.select_dtypes(include=["object"]).columns:
                df[col] = df[col].fillna("-").astype(str).replace({"None": "-", "nan": "-"})
        else:
            target_list = [d for d in sheet_dfs if d["시트명"].iloc[0] == selected_sheet]
            df = target_list[0] if target_list else pd.DataFrame()

    if len(df) > 0:
        with sqlite3.connect(DB_PATH) as conn:
            df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)

    return df, sheet_names, sheet_counts


def filter_by_keywords(df: pd.DataFrame, keyword_str: str, match_all: bool = False) -> pd.DataFrame:
    """쉼표(,)로 구분된 키워드 조합(OR / AND)으로 데이터프레임을 필터링합니다."""
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
    """자동 생성된 SQLite DB에 직접 SQL 쿼리를 실행하고 None 값을 '-'로 정제하여 반환합니다."""
    with sqlite3.connect(DB_PATH) as conn:
        res_df = pd.read_sql_query(sql_query, conn)
    for col in res_df.select_dtypes(include=["object"]).columns:
        res_df[col] = res_df[col].fillna("-").astype(str).replace({"None": "-", "nan": "-"})
    return res_df


def export_db_to_sql_script(current_sql: str = "") -> bytes:
    """Notepad++ 등에서 바로 열어볼 수 있는 UTF-8 SQL 덤프 파일을 생성합니다."""
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