import sqlite3
import pandas as pd
from pathlib import Path
from datetime import datetime

DB_PATH = Path("internal_analysis.db")
TABLE_NAME = "internal_data"


def load_and_sync_to_db(uploaded_file) -> pd.DataFrame:
    """CSV/Excel 파일을 읽고 자동으로 로컬 SQLite DB(internal_analysis.db) 테이블로 변환·저장합니다."""
    if uploaded_file.name.endswith(".csv"):
        df = pd.read_csv(uploaded_file)
    else:
        df = pd.read_excel(uploaded_file)

    with sqlite3.connect(DB_PATH) as conn:
        df.to_sql(TABLE_NAME, conn, if_exists="replace", index=False)

    return df


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
    """자동 생성된 SQLite DB에 직접 SQL 쿼리를 실행하여 결과를 반환합니다."""
    with sqlite3.connect(DB_PATH) as conn:
        return pd.read_sql_query(sql_query, conn)


def export_db_to_sql_script(current_sql: str = "") -> bytes:
    """
    Notepad++ 등 텍스트 에디터에서 바로 열어 DB 스키마(CREATE TABLE),
    전체 데이터(INSERT INTO), 사용자 조회 쿼리(SELECT)를 한눈에 확인할 수 있는 .sql 파일을 생성합니다.
    """
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

    lines.extend([
        "",
        "-- [3] Notepad++ 참고용 자주 쓰는 분석 SQL 예시",
        f"-- SELECT * FROM {TABLE_NAME} WHERE 라인 IN ('A라인', 'B라인');",
        f"-- SELECT 라인, SUM(생산량) AS 총생산량, SUM(불량수) AS 총불량수, ROUND(CAST(SUM(불량수) AS FLOAT)/SUM(생산량)*100, 2) AS 불량률 FROM {TABLE_NAME} GROUP BY 라인;",
    ])

    # Notepad++에서 한글이 절대 깨지지 않도록 UTF-8 BOM 포함 인코딩 반환
    return "\n".join(lines).encode("utf-8-sig")