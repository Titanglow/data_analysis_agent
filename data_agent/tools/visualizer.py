# data_agent/tools/visualizer.py
import streamlit as st
import pandas as pd
import plotly.express as px
from data_agent.config import DISTINCT_PALETTE, DEFAULT_DEFECT_THRESHOLD
from data_agent.excel_exporter import (
    analyze_color_cell_contexts,
    cluster_text_series_by_topic,
    shorten_long_label,
    evaluate_threshold_verdict,
)


def is_manufacturing_df(df: pd.DataFrame) -> bool:
    if df is None or len(df) == 0:
        return False
    cols = df.columns.tolist()
    return "생산량" in cols and "불량수" in cols


def validate_uploaded_dataframe(df: pd.DataFrame) -> dict:
    if df is None or len(df) == 0:
        return {"status": "empty", "domain_type": "빈 데이터", "warnings": ["데이터가 비어 있습니다."], "metrics": {}}

    all_cols = df.columns.tolist()
    cols_lower = " ".join(str(c).lower() for c in all_cols)
    total_rows = len(df)

    if is_manufacturing_df(df):
        domain_type = "🏭 제조·생산 품질 데이터"
    elif any(k in cols_lower for k in ["procedure", "pass/fail", "sub system", "test type", "automated", "pre-condition", "actual", "priority", "category"]):
        domain_type = "🧪 SW/시스템 검증(TC) 데이터"
    else:
        domain_type = "📄 일반 사내 업무/문서 데이터"

    empty_cells = 0
    total_cells = total_rows * len(all_cols)
    single_val_cols = []
    informative_cols = []

    for col in all_cols:
        s = pd.Series([str(v).strip() for v in df[col].tolist()])
        empty_c = int(s.isin(["-", "", "None", "nan", "N/A"]).sum() + s.str.startswith("[빈칸").sum())
        empty_cells += empty_c
        valid_s = s[~s.isin(["-", "", "None", "nan", "N/A"])]
        nunq = valid_s.nunique()
        if nunq == 1 and total_rows > 1:
            single_val_cols.append(f"{col}('{valid_s.iloc[0]}')")
        elif 2 <= nunq <= 35:
            informative_cols.append(str(col))

    fill_rate = round((1 - (empty_cells / total_cells)) * 100, 1) if total_cells > 0 else 100.0
    color_df = analyze_color_cell_contexts(df)
    color_cnt = len(color_df)

    id_col = next((c for c in all_cols if str(c).lower() in ["id", "tc id", "tc_id", "_polarion"]), None)
    dup_id_cnt = 0
    if id_col:
        valid_ids = pd.Series([str(v).strip() for v in df[id_col].tolist()])
        valid_ids = valid_ids[~valid_ids.isin(["-", "", "None", "nan"])]
        dup_id_cnt = int(valid_ids.duplicated().sum())

    warnings = []
    if single_val_cols:
        warnings.append(f"단일값(100% 동일) 컬럼 {len(single_val_cols)}개 감지 ({', '.join(single_val_cols[:3])}) → 차트 왜곡 방지를 위해 주요 비교 축에서 자동 제외됨")
    if dup_id_cnt > 0:
        warnings.append(f"식별자([{id_col}]) 중복 {dup_id_cnt}건 감지 → 다중 시트 병합 또는 중복 기재 확인 권장")
    if fill_rate < 85.0:
        warnings.append(f"전체 셀 작성률 {fill_rate}% (미기입·N/A 비율 {round(100 - fill_rate, 1)}%) → 미기입 항목 점검 필요")

    return {
        "status": "ok",
        "domain_type": domain_type,
        "fill_rate": fill_rate,
        "color_cnt": color_cnt,
        "dup_id_cnt": dup_id_cnt,
        "single_val_cols": single_val_cols,
        "informative_cols": informative_cols,
        "warnings": warnings,
    }


def render_validation_banner(df: pd.DataFrame, threshold: float = 1.5):
    val_res = validate_uploaded_dataframe(df)
    if val_res["status"] == "empty":
        return

    verdict_str = evaluate_threshold_verdict(df, threshold=threshold)
    if "확인 필요" in verdict_str:
        st.warning(f"**{verdict_str}**")
    else:
        st.success(f"**{verdict_str}**")

    with st.expander(
        f"🛡️ [W47 양식 자동 검증 완료] 유형: {val_res['domain_type']} | "
        f"데이터 충실도(작성률): {val_res['fill_rate']}% | "
        f"비교 유효 차원: {len(val_res['informative_cols'])}개 컬럼 (상세 보기)",
        expanded=False,
    ):
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("감지된 문서 양식", val_res["domain_type"])
        c2.metric("전체 셀 작성률", f"{val_res['fill_rate']} %")
        c3.metric("시각화 적합 분류 열", f"{len(val_res['informative_cols'])} 개")
        c4.metric("색상강조 / 중복ID", f"{val_res['color_cnt']}건 / {val_res['dup_id_cnt']}건")

        if val_res["informative_cols"]:
            st.caption(f"✅ **자동 감지된 핵심 비교 컬럼:** `{', '.join(val_res['informative_cols'][:8])}`")
        if val_res["warnings"]:
            for w in val_res["warnings"]:
                st.info(f"💡 **양식 자동 보정 안내:** {w}")


def _pick_best_categorical_col(
    df: pd.DataFrame,
    preferred_names: list[str],
    exclude: list[str] | None = None,
    allow_single_value: bool = False,
) -> str | None:
    exclude_set = set(exclude or [])
    all_cols = [c for c in df.columns if c not in exclude_set]
    if not all_cols:
        return None

    for pref in preferred_names:
        for col in all_cols:
            if str(col).lower() == pref.lower() or pref.lower() in str(col).lower():
                valid_s = pd.Series([str(v).strip() for v in df[col].tolist()])
                valid_s = valid_s[~valid_s.isin(["-", "", "None", "nan"])]
                nunq = valid_s.nunique()
                if 2 <= nunq <= 40:
                    return col

    skip_keywords = ["_polarion", "id", "procedure", "description", "pre-condition", "post-condition", "inputdata", "title", "step", "expected"]
    for col in all_cols:
        if any(k in str(col).lower() for k in skip_keywords):
            continue
        valid_s = pd.Series([str(v).strip() for v in df[col].tolist()])
        valid_s = valid_s[~valid_s.isin(["-", "", "None", "nan"])]
        nunq = valid_s.nunique()
        if 2 <= nunq <= 25:
            return col

    if allow_single_value:
        for pref in preferred_names:
            for col in all_cols:
                if pref.lower() in str(col).lower():
                    return col

    return all_cols[0]


def render_kpi_summary_cards(target_df: pd.DataFrame, threshold: float = DEFAULT_DEFECT_THRESHOLD):
    all_cols = target_df.columns.tolist()
    k1, k2, k3, k4 = st.columns(4)

    if is_manufacturing_df(target_df):
        p_sum = int(target_df["생산량"].sum())
        d_sum = int(target_df["불량수"].sum())
        rate = (d_sum / p_sum * 100) if p_sum > 0 else 0.0
        top_def = target_df.groupby("불량유형")["불량수"].sum().idxmax() if "불량유형" in all_cols and d_sum > 0 else "-"
        delta_str = f"기준({threshold}%) 초과 🚨" if rate > threshold else f"기준({threshold}%) 이내 ✅"
        k1.metric("📌 조회 데이터 건수", f"{len(target_df):,} 건")
        k2.metric("🏭 총 생산량 / 불량수", f"{p_sum:,}개 / {d_sum:,}개")
        k3.metric(f"🚨 평균 불량률 (기준 {threshold}%)", f"{rate:.2f} %", delta=delta_str, delta_color="inverse" if rate > threshold else "normal")
        k4.metric("🔍 최다 발생 불량유형", str(top_def))
    else:
        cat_col = _pick_best_categorical_col(target_df, ["Sub System", "Category", "Domain", "App", "IP", "Priority", "시트명", "분류"])
        top_cat_str = "-"
        exceed_col_cnt = 0
        total_cnt = len(target_df)

        for c in all_cols:
            vals = [str(v).strip() for v in target_df[c].tolist()]
            empty_c = sum(1 for v in vals if v in ("-", "", "None", "nan"))
            if total_cnt > 0 and (empty_c / total_cnt * 100) > threshold:
                exceed_col_cnt += 1

        if cat_col:
            vc = target_df[cat_col].astype(str).value_counts()
            if not vc.empty:
                top_val = vc.index[0]
                top_cnt = int(vc.iloc[0])
                top_pct = round(top_cnt / len(target_df) * 100, 1)
                top_cat_str = f"{shorten_long_label(top_val, 14)} ({top_cnt}건·{top_pct}%)"

        auto_or_res_col = _pick_best_categorical_col(
            target_df,
            ["Priority", "Automated", "Test Type", "Execution Status", "PASS/FAIL", "Actual Result", "Result"],
            exclude=[cat_col] if cat_col else [],
        )
        sub_metric_label = "⚙️ 주요 속성 분포"
        sub_metric_val = f"총 {len(all_cols)}개 컬럼"
        if auto_or_res_col:
            vc2 = target_df[auto_or_res_col].astype(str).value_counts()
            if not vc2.empty:
                v_name = vc2.index[0]
                v_cnt = int(vc2.iloc[0])
                v_pct = round(v_cnt / len(target_df) * 100, 1)
                sub_metric_label = f"📊 [{auto_or_res_col}] 최다"
                sub_metric_val = f"{shorten_long_label(v_name, 14)} ({v_pct}%)"

        k1.metric("📌 총 조회 데이터 건수", f"{len(target_df):,} 건")
        k2.metric(f"🏆 [{cat_col or '분류'}] 1위 비중", top_cat_str)
        k3.metric(sub_metric_label, sub_metric_val)
        k4.metric(
            f"🎚️ 미기입률 기준({threshold}%) 초과 열",
            f"{exceed_col_cnt} 개 컬럼",
            delta="점검 필요 🚨" if exceed_col_cnt > 0 else "전 컬럼 기준 이내 ✅",
            delta_color="inverse" if exceed_col_cnt > 0 else "normal",
        )


def render_smart_charts(
    target_df: pd.DataFrame,
    key_prefix: str = "main",
    defect_threshold: float = DEFAULT_DEFECT_THRESHOLD,
):
    if len(target_df) == 0:
        st.info("조회된 데이터가 없습니다.")
        return

    render_kpi_summary_cards(target_df, threshold=defect_threshold)
    st.markdown("<div style='margin-bottom:8px;'></div>", unsafe_allow_html=True)

    num_cols = target_df.select_dtypes(include=["number"]).columns.tolist()
    all_cols = target_df.columns.tolist()

    if is_manufacturing_df(target_df) or (num_cols and len(num_cols) >= 2):
        cat_cols = target_df.select_dtypes(include=["object", "category"]).columns.tolist()
        line_col = "라인" if "라인" in all_cols else ("생산라인" if "생산라인" in all_cols else (cat_cols[0] if cat_cols else all_cols[0]))
        defect_type_col = next((c for c in ["불량유형", "불량원인", "유형", "분류", "카테고리"] if c in all_cols), None)
        time_col = next((c for c in ["일자", "날짜", "주차", "월"] if c in all_cols), None)
        shift_col = next((c for c in ["시간대", "교대", "근무조", "시간"] if c in all_cols), None)
        val_col = "불량수" if "불량수" in num_cols else num_cols[0]

        row1_c1, row1_c2 = st.columns(2)
        with row1_c1:
            if "생산량" in num_cols and "불량수" in num_cols and line_col in all_cols:
                agg_line = target_df.groupby(line_col, as_index=False)[["생산량", "불량수"]].sum()
                agg_line["불량률(%)"] = (agg_line["불량수"] / agg_line["생산량"] * 100).round(2)
                agg_line = agg_line.sort_values(by="불량수", ascending=True)
                agg_line["표시라벨"] = agg_line.apply(
                    lambda r: f"{int(r['불량수']):,}개 ({r['불량률(%)']}%) {'🚨초과' if r['불량률(%)'] > defect_threshold else '✅'}",
                    axis=1,
                )
                fig1 = px.bar(
                    agg_line, x="불량수", y=line_col, orientation="h", color=line_col, text="표시라벨",
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"① [{line_col}]별 총 불량수 및 기준({defect_threshold}%) 판정",
                )
                fig1.update_traces(textposition="outside", cliponaxis=False)
                fig1.update_layout(showlegend=False, margin=dict(l=20, r=70, t=45, b=20))
                st.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_num1")
            else:
                agg_basic = target_df.groupby(line_col, as_index=False)[val_col].sum().sort_values(by=val_col, ascending=True)
                fig1 = px.bar(
                    agg_basic, x=val_col, y=line_col, orientation="h", color=line_col, text_auto=True,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"① [{line_col}]별 {val_col} 합계 순위",
                )
                fig1.update_traces(textposition="outside", cliponaxis=False)
                st.plotly_chart(fig1, use_container_width=True, key=f"{key_prefix}_num1")

        with row1_c2:
            if defect_type_col:
                agg_pie = target_df.groupby(defect_type_col, as_index=False)[val_col].sum()
                fig2 = px.pie(
                    agg_pie, names=defect_type_col, values=val_col, hole=0.42,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"② [{defect_type_col}]별 {val_col} 점유율 비중",
                )
                fig2.update_traces(textinfo="label+percent+value")
                st.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_num2")
            else:
                agg_pie = target_df.groupby(line_col, as_index=False)[val_col].sum()
                fig2 = px.pie(
                    agg_pie, names=line_col, values=val_col, hole=0.42,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"② [{line_col}]별 {val_col} 비중",
                )
                st.plotly_chart(fig2, use_container_width=True, key=f"{key_prefix}_num2")

        row2_c1, row2_c2 = st.columns(2)
        with row2_c1:
            if time_col and "생산량" in num_cols and "불량수" in num_cols:
                agg_time = target_df.groupby([time_col, line_col], as_index=False)[["생산량", "불량수"]].sum()
                agg_time["불량률(%)"] = (agg_time["불량수"] / agg_time["생산량"] * 100).round(2)
                fig3 = px.line(
                    agg_time, x=time_col, y="불량률(%)", color=line_col, markers=True, text="불량률(%)",
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"③ [{time_col}]별 불량률(%) 추이 (사내 기준 {defect_threshold}% 빨간 점선)",
                )
                fig3.add_hline(
                    y=defect_threshold, line_dash="dash", line_color="#DC2626",
                    annotation_text=f"사내 기준 ({defect_threshold}%)",
                )
                st.plotly_chart(fig3, use_container_width=True, key=f"{key_prefix}_num3")
        with row2_c2:
            if shift_col:
                agg_shift = target_df.groupby([shift_col, defect_type_col or line_col], as_index=False)[val_col].sum()
                fig4 = px.bar(
                    agg_shift, x=shift_col, y=val_col, color=defect_type_col or line_col, barmode="group", text_auto=True,
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"④ [{shift_col}] 교대별 비교",
                )
                st.plotly_chart(fig4, use_container_width=True, key=f"{key_prefix}_num4")
        return

    primary_cat = _pick_best_categorical_col(
        target_df,
        ["Sub System", "Category", "Domain", "App", "IP", "Priority", "시트명", "분류"],
        allow_single_value=True,
    )
    secondary_cat = _pick_best_categorical_col(
        target_df,
        ["Domain", "Priority", "Automated", "Test Type", "Test Technique", "Execution Status", "Sub System", "IP", "Actual Result", "PASS/FAIL"],
        exclude=[primary_cat] if primary_cat else [],
        allow_single_value=False,
    ) or _pick_best_categorical_col(
        target_df, ["Automated", "Test Type", "Execution Status", "Actual Result", "Type"],
        exclude=[primary_cat] if primary_cat else [], allow_single_value=True,
    )

    title_col = next(
        (c for c in all_cols if any(k in str(c).lower() for k in ["title", "test case", "name", "제목", "시나리오", "purpose", "description", "내용"])),
        None,
    )

    cut_count_line = round(len(target_df) * (defect_threshold / 100.0), 1)

    r1_c1, r1_c2 = st.columns(2)

    with r1_c1:
        vc1 = target_df[primary_cat].astype(str).value_counts().head(12).reset_index()
        vc1.columns = [primary_cat, "건수"]
        vc1["비중(%)"] = (vc1["건수"] / len(target_df) * 100).round(1)
        vc1["표시라벨"] = vc1.apply(
            lambda r: f"{r['건수']:,}건 ({r['비중(%)']}%) {'🎯기준이상' if r['비중(%)'] >= defect_threshold else ''}".strip(),
            axis=1,
        )
        vc1["짧은라벨"] = vc1[primary_cat].apply(lambda x: shorten_long_label(str(x), 24))
        vc1 = vc1.sort_values(by="건수", ascending=True)

        fig_c1 = px.bar(
            vc1, x="건수", y="짧은라벨", orientation="h", color="짧은라벨", text="표시라벨",
            color_discrete_sequence=DISTINCT_PALETTE,
            title=f"① [{primary_cat}] TOP 분포 및 사내 기준({defect_threshold}% = {cut_count_line}건) 비교",
        )
        fig_c1.add_vline(
            x=cut_count_line, line_dash="dash", line_color="#DC2626",
            annotation_text=f"사내 기준 ({defect_threshold}%)",
        )
        fig_c1.update_traces(textposition="outside", cliponaxis=False)
        fig_c1.update_layout(
            xaxis_title="데이터 건수 (건)", yaxis_title=primary_cat,
            showlegend=False, margin=dict(l=20, r=85, t=48, b=25),
        )
        st.plotly_chart(fig_c1, use_container_width=True, key=f"{key_prefix}_cat1")

    with r1_c2:
        vc2 = target_df[secondary_cat].astype(str).value_counts().head(10).reset_index()
        vc2.columns = [secondary_cat, "건수"]
        vc2["짧은라벨"] = vc2[secondary_cat].apply(lambda x: shorten_long_label(str(x), 24))
        fig_c2 = px.pie(
            vc2, names="짧은라벨", values="건수", hole=0.42,
            color_discrete_sequence=DISTINCT_PALETTE,
            title=f"② [{secondary_cat}] 분포 및 점유율 비중",
        )
        fig_c2.update_traces(textinfo="label+percent+value")
        fig_c2.update_layout(margin=dict(l=20, r=20, t=48, b=25))
        st.plotly_chart(fig_c2, use_container_width=True, key=f"{key_prefix}_cat2")

    r2_c1, r2_c2 = st.columns(2)

    with r2_c1:
        if title_col and target_df[title_col].nunique() > 1:
            topic_df = cluster_text_series_by_topic(target_df[title_col], max_groups=10)
            if not topic_df.empty:
                topic_plot = topic_df.sort_values(by="건수", ascending=True).copy()
                topic_plot["표시라벨"] = topic_plot.apply(lambda r: f"{r['건수']:,}건 ({r['비중(%)']})", axis=1)
                fig_c3 = px.bar(
                    topic_plot, x="건수", y="주제/분류 그룹", orientation="h",
                    color="주제/분류 그룹", text="표시라벨",
                    color_discrete_sequence=DISTINCT_PALETTE,
                    title=f"③ [{title_col}] 유사 시나리오·키워드 자동 그룹핑 TOP 분포",
                )
                fig_c3.update_traces(textposition="outside", cliponaxis=False)
                fig_c3.update_layout(
                    xaxis_title="건수 (건)", yaxis_title="핵심 시나리오 그룹",
                    showlegend=False, margin=dict(l=20, r=70, t=48, b=25),
                )
                st.plotly_chart(fig_c3, use_container_width=True, key=f"{key_prefix}_cat3")

    with r2_c2:
        third_cat = _pick_best_categorical_col(
            target_df,
            ["Priority", "Test Technique", "Automated", "Test Type", "IP", "Sub System", "Domain", "PASS/FAIL"],
            exclude=[primary_cat] if primary_cat else [],
            allow_single_value=False,
        ) or secondary_cat

        if primary_cat and third_cat and primary_cat != third_cat:
            top_prim_vals = target_df[primary_cat].astype(str).value_counts().head(10).index.tolist()
            sub_cross = target_df[target_df[primary_cat].astype(str).isin(top_prim_vals)].copy()
            cross_df = sub_cross.groupby([primary_cat, third_cat], as_index=False).size()
            cross_df.columns = [primary_cat, third_cat, "건수"]
            cross_df[primary_cat] = cross_df[primary_cat].apply(lambda x: shorten_long_label(str(x), 18))
            cross_df[third_cat] = cross_df[third_cat].apply(lambda x: shorten_long_label(str(x), 18))

            fig_c4 = px.bar(
                cross_df, x=primary_cat, y="건수", color=third_cat, barmode="group", text_auto=True,
                color_discrete_sequence=DISTINCT_PALETTE,
                title=f"④ [{primary_cat}] × [{third_cat}] 다차원 교차 분포",
            )
            fig_c4.update_traces(textposition="outside", cliponaxis=False)
            fig_c4.update_layout(margin=dict(l=20, r=20, t=48, b=25))
            st.plotly_chart(fig_c4, use_container_width=True, key=f"{key_prefix}_cat4")