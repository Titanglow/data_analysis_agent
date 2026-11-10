import sqlite3
import pandas as pd
from pathlib import Path

DB_PATH = Path("internal_analysis.db")
TABLE_NAME = "internal_data"

def load_and_sync_to_db(uploaded_file) -> pd.DataFrame:
    """CSV/Excel 파일을 읽고 자동으로 로컬 SQLite DB(internal_analysis.db) 테이블로 변환·저장합니다."""
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    # 자동으로 SQLite DB 파일 생성 및 테이블 덮어쓰기 (수작업 불필요)
    with sqlite3.connect(DB_PATH) as conn:
        df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)

    return df

def filter_by_keywords(df: pd.DataFrame, keyword_str: str, match_all: bool = False) -> pd.DataFrame:
    """
    입력한 키워드(쉼표로 구분하여 조합 가능)가 포함된 행만 검색·수집합니다.
    예: 'B라인, W45' 입력 시 해당 키워드가 들어간 데이터만 추출
    """
    if not keyword_str or not keyword_str.strip():
        return df

    keywords = [k.strip() for k in keyword_str.split(",") if k.strip()]
    if not keywords:
        return df

    # 데이터프레임 전체 컬럼을 문자열로 변환하여 키워드 포함 여부 검색
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
    """자동 생성된 SQLite DB에 직접 SQL 쿼리를 실행하여 결과를 반환합니다."""
    with sqlite3.connect(DB_PATH) as conn:
        return pd.read_sql_query(sql_query, conn)