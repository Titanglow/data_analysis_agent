# data_agent/tools/data_loader.py
import sqlite3
import pandas as pd
import io
import re
from pathlib import Path
import xml.etree.ElementTree as ET
import docx

DB_PATH = Path("internal_analysis.db")
TABLE_NAME = "internal_data"


def _safe_str(val) -> str:
    """Pandas 3.0 NaN(float) 정규식 에러 방지용 안전한 문자열 변환"""
    if pd.isna(val):
        return ""
    return str(val)


def _extract_word_content(file_obj) -> pd.DataFrame:
    doc = docx.Document(file_obj)
    rows = []

    for i, para in enumerate(doc.paragraphs):
        text = para.text.strip()
        if text:
            rows.append({
                "시트명": "문서본문",
                "항목_구분": "문단",
                "내용": text,
                "순서": i + 1
            })

    for t_idx, table in enumerate(doc.tables):
        for r_idx, row in enumerate(table.rows):
            row_data = [cell.text.strip() for cell in row.cells]
            if any(row_data):
                rows.append({
                    "시트명": f"워드내장표_{t_idx + 1}",
                    "항목_구분": f"표_{r_idx + 1}행",
                    "내용": " | ".join(row_data),
                    "순서": r_idx + 1
                })

    if rows:
        return pd.DataFrame(rows)
    return pd.DataFrame([{"시트명": "문서본문", "항목_구분": "알림", "내용": "문서에 텍스트가 없습니다.", "순서": 1}])


def load_and_sync_to_db(uploaded_file, selected_sheet="ALL") -> tuple[pd.DataFrame, list[str], dict[str, int], dict[str, pd.DataFrame]]:
    file_ext = Path(uploaded_file.name).suffix.lower()
    df_dict = {}

    if file_ext == ".csv":
        try:
            df = pd.read_csv(uploaded_file, encoding='utf-8-sig')
        except Exception:
            uploaded_file.seek(0)
            df = pd.read_csv(uploaded_file, encoding='cp949')
        df_dict["Sheet1"] = df
    elif file_ext == ".xlsx":
        xls = pd.ExcelFile(uploaded_file)
        for sheet_name in xls.sheet_names:
            df_dict[sheet_name] = pd.read_excel(xls, sheet_name=sheet_name)
    elif file_ext == ".docx":
        df_dict["Word_Content"] = _extract_word_content(uploaded_file)
    elif file_ext == ".txt":
        text = uploaded_file.read().decode("utf-8", errors="ignore")
        lines = [line.strip() for line in text.splitlines() if line.strip()]
        df_dict["Text_Content"] = pd.DataFrame({"시트명": "텍스트문서", "내용": lines})
    else:
        raise ValueError(f"지원하지 않는 파일 형식입니다: {file_ext}")

    main_sheets = []
    aux_sheets = {}
    sheet_counts = {}

    for s_name, s_df in df_dict.items():
        if s_df.empty:
            continue
        sheet_counts[s_name] = len(s_df)
        if any(kw in s_name.upper() for kw in ["SUMMARY", "ENV", "요약", "환경", "DASHBOARD"]):
            aux_sheets[s_name] = s_df
        else:
            main_sheets.append(s_name)

    if not main_sheets and aux_sheets:
        main_sheets = list(aux_sheets.keys())
        aux_sheets = {}

    if selected_sheet != "ALL" and selected_sheet in df_dict:
        combined_df = df_dict[selected_sheet]
    else:
        dfs_to_concat = []
        for s in main_sheets:
            temp_df = df_dict[s].copy()
            temp_df.insert(0, '시트명', s)
            dfs_to_concat.append(temp_df)
        combined_df = pd.concat(dfs_to_concat, ignore_index=True) if dfs_to_concat else pd.DataFrame()

    if "파일명" not in combined_df.columns and len(combined_df) > 0:
        combined_df.insert(0, '파일명', uploaded_file.name)

    if not combined_df.empty:
        with sqlite3.connect(DB_PATH) as conn:
            combined_df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)

    return combined_df, main_sheets, sheet_counts, aux_sheets


def filter_by_keywords(df: pd.DataFrame, keywords_str: str, match_all: bool = False) -> pd.DataFrame:
    keywords = [k.strip().lower() for k in keywords_str.split(',') if k.strip()]
    if not keywords:
        return df

    row_text = df.astype(str).apply(lambda row: ' '.join(row).lower(), axis=1)

    if match_all:
        mask = pd.Series(True, index=df.index)
        for kw in keywords:
            mask &= row_text.str.contains(re.escape(kw), regex=True, case=False, na=False)
    else:
        mask = pd.Series(False, index=df.index)
        for kw in keywords:
            mask |= row_text.str.contains(re.escape(kw), regex=True, case=False, na=False)

    return df[mask]


def run_sql_query(query: str) -> pd.DataFrame:
    if not DB_PATH.exists():
        raise FileNotFoundError("DB 파일이 없습니다. 먼저 데이터를 업로드해주세요.")
    with sqlite3.connect(DB_PATH) as conn:
        return pd.read_sql_query(query, conn)


def export_db_to_sql_script(current_sql: str = "") -> bytes:
    if not DB_PATH.exists():
        return b""
    with sqlite3.connect(DB_PATH) as conn:
        dump_lines = [line for line in conn.iterdump()]

    header = (
        "-- 사내 데이터 분석 Agent SQLite DB 덤프\n"
        "-- Notepad++ 등 텍스트 에디터에서 바로 확인 가능합니다.\n\n"
    )
    if current_sql:
        header += f"-- [현재 화면에 적용된 마지막 SQL 쿼리]\n-- {current_sql}\n\n"

    full_script = header + "\n".join(dump_lines)
    return full_script.encode("utf-8")


def extract_knowledge_text(uploaded_files: list) -> str:
    """
    RAG(검색 증강 생성)를 위해 업로드된 참고 문서(Word, TXT)에서 텍스트를 추출합니다.
    """
    knowledge_parts = []
    for f_obj in uploaded_files:
        ext = Path(f_obj.name).suffix.lower()
        try:
            if ext == ".txt":
                text = f_obj.read().decode("utf-8", errors="ignore")
                knowledge_parts.append(f"[{f_obj.name} 내용]\n{text[:3000]}")
            elif ext == ".docx":
                doc = docx.Document(f_obj)
                text = "\n".join([p.text for p in doc.paragraphs if p.text.strip()])
                knowledge_parts.append(f"[{f_obj.name} 내용]\n{text[:3000]}")
        except Exception as e:
            knowledge_parts.append(f"[{f_obj.name} 읽기 실패: {e}]")

    combined_knowledge = "\n\n".join(knowledge_parts)
    return combined_knowledge[:6000]