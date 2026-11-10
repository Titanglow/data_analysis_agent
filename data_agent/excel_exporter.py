import io
import pandas as pd

def export_analysis_to_excel(df: pd.DataFrame, summary_df: pd.DataFrame, report_text: str, elapsed_time: float, mode: str) -> bytes:
    """
    메모리 버퍼에 다중 시트 엑셀 파일을 생성하여 바이트 형태로 반환합니다.
    (Streamlit 다운로드 버튼과 즉시 연동)
    """
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        # Sheet 1: AI 인사이트 리포트 및 메타데이터
        report_lines = [line for line in report_text.split("\n") if line.strip()]
        report_df = pd.DataFrame({
            "항목": ["실행 모드", "응답 소요 시간(초)"] + [f"인사이트 {i+1}" for i in range(len(report_lines))],
            "내용": [mode, f"{elapsed_time}초"] + report_lines
        })
        report_df.to_excel(writer, sheet_name="1_AI_인사이트_리포트", index=False)

        # Sheet 2: 기초 통계 및 집계 데이터
        summary_df.to_excel(writer, sheet_name="2_기초통계_집계")

        # Sheet 3: 원본 데이터
        df.to_excel(writer, sheet_name="3_원본데이터", index=False)

    return output.getvalue()