def render_pipeline_status(active_step: int = None):
    """
    1~5단계 실시간 파이프라인 가동현황 표시.
    분석 결과(st.session_state["analysis_results"])가 존재하면 4번·5번까지 모두 초록색 완료(5단계)로 자동 전환합니다.
    """
    if active_step is None:
        if st.session_state.get("analysis_results"):
            active_step = 5  # 분석 완료 시 1~5번 전체 초록색 활성화!
        elif st.session_state.get("filtered_df") is not None:
            active_step = 2
        else:
            active_step = 1

    steps = [
        (1, "1. 데이터 로드·색상 스캔"),
        (2, "2. 스마트 SQL·다차원 집계"),
        (3, "3. AI 비서 다중 모델 추론"),
        (4, "4. 리포트 검증·양식 보정"),
        (5, "5. 통합 대시보드·엑셀 완료"),
    ]

    cards_html = []
    for num, label in steps:
        if num <= active_step:
            bg = "#ecfdf5"       # 연한 에메랄드 배경
            border = "#10b981"   # 초록색 테두리
            text_col = "#065f46"
            icon = "✅"
            status_txt = "가동 완료"
        else:
            bg = "#f8fafc"       # 대기 중 회색 배경
            border = "#cbd5e1"
            text_col = "#64748b"
            icon = "⏳"
            status_txt = "대기 중"

        cards_html.append(
            f'<div style="flex:1; min-width:140px; background:{bg}; border:2px solid {border}; '
            f'border-radius:10px; padding:10px 12px; text-align:center; box-shadow:0 1px 3px rgba(0,0,0,0.05);">'
            f'<div style="font-size:0.85rem; font-weight:700; color:{text_col};">{icon} {label}</div>'
            f'<div style="font-size:0.75rem; color:{text_col}; margin-top:4px;">{status_txt}</div>'
            f'</div>'
        )

    html = (
        '<div style="margin: 8px 0 18px 0;">'
        '<div style="font-size:0.9rem; font-weight:700; color:#1e293b; margin-bottom:6px;">'
        '⚡ 실시간 파이프라인 가동현황</div>'
        f'<div style="display:flex; gap:10px; flex-wrap:wrap;">{"".join(cards_html)}</div>'
        '</div>'
    )
    return html