import streamlit as st
import pandas as pd
import datetime as dt
import uuid

from common_ui import (
    inject_logistics_theme,
    set_page,
    KPI,
    render_kpis,
    bar_topN,
    table_block,
    download_excel,
    card_open,
    card_close,
)

from qc_core import run_qc_efficiency
from audit_store import sha256_bytes, upload_export_bytes, insert_audit_run


def main():
    inject_logistics_theme()
    set_page("驗收作業效能（KPI）", icon="✅")
    st.caption("驗收作業｜人時效率｜AM / PM 班別｜KPI 達標分析")

    # ======================
    # Sidebar：計算條件設定
    # ======================
    with st.sidebar:
        st.header("⚙️ 計算條件設定")

        operator = st.text_input("分析執行人（Operator）")
        top_n = st.number_input("效率排行顯示人數（Top N）", 10, 100, 30, step=5)

        st.markdown("#### 排除區間（非作業時段）")
        st.caption("用於排除支援、離站、停機、非驗收作業時間")

        if "skip_rules" not in st.session_state:
            st.session_state.skip_rules = []

        user = st.text_input("資料登錄人（Data Entry，可留空）")
        t_start = st.time_input("開始時間")
        t_end = st.time_input("結束時間")

        if st.button("➕ 新增排除區間"):
            st.session_state.skip_rules.append(
                {
                    "user": user.strip(),
                    "t_start": t_start,
                    "t_end": t_end,
                }
            )

        if st.session_state.skip_rules:
            st.dataframe(
                pd.DataFrame(st.session_state.skip_rules),
                use_container_width=True,
                hide_index=True,
            )

    # ======================
    # 上傳資料
    # ======================
    card_open("📤 上傳作業原始資料（驗收）")
    uploaded = st.file_uploader(
        "上傳驗收作業原始資料",
        type=["xlsx", "xls", "csv"],
        label_visibility="collapsed",
    )
    run = st.button("🚀 產出 KPI", type="primary", disabled=uploaded is None)
    card_close()

    if not run:
        st.info("請先上傳驗收作業原始資料")
        return

    # ======================
    # 計算
    # ======================
    with st.spinner("KPI 計算中，請稍候..."):
        result = run_qc_efficiency(
            uploaded.getvalue(),
            uploaded.name,
            st.session_state.skip_rules,
        )

    df = result.get("ampm_df", pd.DataFrame())
    idle_df = result.get("idle_df", pd.DataFrame())
    target = float(result.get("target_eff", 20.0))

    if df.empty or "時段" not in df.columns:
        st.error("資料缺少『時段』欄位，無法區分 AM / PM 班別")
        return

    # 顯示層轉換為 AM / PM
    df["班別"] = df["時段"].replace({"上午": "AM 班", "下午": "PM 班"})
    am_df = df[df["班別"] == "AM 班"].copy()
    pm_df = df[df["班別"] == "PM 班"].copy()

    # ======================
    # KPI 區塊
    # ======================
    col_l, col_r = st.columns(2)

    def render_shift(title, sdf):
        card_open(f"{title} KPI")
        render_kpis(
            [
                KPI("人數", f"{len(sdf):,}"),
                KPI("總驗收筆數", f"{sdf['筆數'].sum():,}"),
                KPI("總工時", f"{sdf['總工時'].sum():.2f}"),
                KPI("平均效率", f"{sdf['效率'].mean():.2f}"),
                KPI(
                    "達標率",
                    f"{(sdf['效率'] >= target).mean():.0%}",
                ),
            ]
        )
        card_close()

        card_open(f"{title} 效率排行（Top {top_n}）")
        bar_topN(
            sdf,
            x_col="姓名",
            y_col="效率",
            hover_cols=["筆數", "總工時"],
            top_n=top_n,
            target=target,
        )
        card_close()

    with col_l:
        render_shift("🌓 AM 班（驗收）", am_df)

    with col_r:
        render_shift("🌙 PM 班（驗收）", pm_df)

    # ======================
    # 匯出
    # ======================
    if result.get("xlsx_bytes"):
        card_open("⬇️ 匯出 KPI 報表")
        download_excel(result["xlsx_bytes"], result.get("xlsx_name", "驗收作業KPI.xlsx"))
        card_close()

    # ======================
    # 稽核留存
    # ======================
    st.divider()
    st.subheader("🧾 稽核留存狀態")

    try:
        export_path = None
        if result.get("xlsx_bytes"):
            export_path = upload_export_bytes(
                content=result["xlsx_bytes"],
                object_path=f"qc_runs/{dt.datetime.now():%Y%m%d}/{uuid.uuid4().hex}.xlsx",
            )

        payload = {
            "app_name": "驗收作業效能（KPI）",
            "operator": operator or None,
            "source_filename": uploaded.name,
            "source_sha256": sha256_bytes(uploaded.getvalue()),
            "params": {
                "top_n": top_n,
                "target_eff": target,
                "skip_rules": st.session_state.skip_rules,
            },
            "kpi_am": {"avg_eff": am_df["效率"].mean(), "people": len(am_df)},
            "kpi_pm": {"avg_eff": pm_df["效率"].mean(), "people": len(pm_df)},
            "export_object_path": export_path,
        }

        row = insert_audit_run(payload)
        st.success(f"✅ 已成功留存本次分析（ID：{row.get('id')}）")

    except Exception as e:
        st.error("❌ 稽核留存失敗")
        st.code(str(e))


if __name__ == "__main__":
    main()
