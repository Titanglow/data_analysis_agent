import re
import streamlit as st
import pandas as pd
import plotly.express as px
from data_agent.config import DISTINCT_PALETTE
from data_agent.excel_exporter import (
    analyze_color_cell_contexts,
    cluster_text_series_by_topic,
    shorten_long_label,
)


def classify_verification_goal(text: str) -> str:
    """기대결과(Expected Result)나 본문 문장을 검증/업무 목적 유형으로 분류합니다."""
    s = str(text)
    if any(k in s for k in ["오류", "에러", "않는다", "없음", "실패", "차단", "불가"]):
        return "🚫 오류·예외 방지 및 미표시 검증"
    if any(k in s for k in ["이동", "전환", "열리", "열린", "돌아", "접속"]):
        return "🔗 화면/문서 이동 및 링크 전환 검증"
    if any(k in s for k in ["숨겨", "펼쳐", "접기", "펼치", "유지", "상태", "토글"]):
        return "⚙️ UI 상태 변화(접기/펼치기/유지) 검증"
    if any(k in s for k in ["표시된", "나타난", "보인다", "출력", "목록"]):
        return "🖥️ 화면/데이터 정상 표시 검증"
    return "📌 기타 기능 동작/설명"


def render_kpi_summary_cards(target_df: pd.DataFrame):
    """차트 상단에 핵심 수치 4가지를 한눈에 보여주는 KPI 메트릭 카드를 출력합니다."""
    all_cols = target_df.columns.tolist()
    k1, k2, k3, k4 = st.columns(4)

    if "생산량" in all_cols and "불량수" in all_cols:
        p_sum = int(target_df["생산량"].sum())
        d_sum = int(target_df["불량수"].sum())
        rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
        top_def = target_df.groupby("불량유형")["불량수"].sum().idxmax() if "불량유형" in all_cols and d_sum > 0 else "-"
        k1.metric("📌 분석 데이터 건수", f"{len(target_df):,} 건")
        k2.metric("🏭 총 생산량 / 불량수", f"{p_sum:,} 개 / {d_sum:,} 개")
        k3.metric("🚨 평균 불량률 (기준 1.5%)", f"{rate:.2f} %")
        k4.metric("🔍 최다 발생 불량유형", str(top_def))
    else:
        title_col = next(
            (c for c in all_cols if any(k in c.lower() for k in ["문서 섹션", "title", "제목", "분류", "category"])),
            all_cols[0],
        )
        topic_df = cluster_text_series_by_topic(target_df[title_col], max_groups=5)
        top_group = f"{topic_df.iloc[0]['주제/분류 그룹']} ({topic_df.iloc[0]['건수']}건)" if len(topic_df) > 0 else "-"
        color_df = analyze_color_cell_contexts(target_df)
        issue_cnt = len(color_df)
        if issue_cnt == 0 and "핵심 분류" in all_cols:
            issue_cnt = int(target_df["핵심 분류"].astype(str).str.contains("중점|주의|이슈").sum())

        k1.metric("📌 총 조회 항목 수", f"{len(target_df):,} 건")
        k2.metric("📂 분석 컬럼(속성) 수", f"{len(all_cols)} 개")
        k3.metric("🏆 1위 핵심 주제 묶음", top_group[:22])
        k4.metric("🎨 강조/주의 감지 건수", f"{issue_cnt:,} 건")