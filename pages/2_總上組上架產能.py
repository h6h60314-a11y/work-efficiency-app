import streamlit as st
import pandas as pd

from common_ui import (
    set_page,
    KPI,
    render_kpis,
    bar_topN,
    pivot_am_pm,
    table_block,
    download_excel,
)

from shelf_core import run_shelf_efficiency


def render_params():
    """Sidebar 參數：達標門檻 / 空窗門檻 / 排行顯示人數"""
    target_eff = st.number_input("達標門檻（件/小時）", min_value=1, max_value=200, value=20, step=1)
    idle_threshold = st.number_input("空窗門檻（分鐘）", min_value=1, max_value=120, value=10, step=1)
    top_n = st.number_input("排行顯示人數", min_value=10, max_value=100, value=30, step=10)

    return {
        "target_eff": float(target_eff),
        "idle_threshold": int(idle_threshold),
        "top_n": int(top_n),
    }


def _fmt_num(x, digits=2):
    try:
        if x is None:
            return "—"
        return f"{float(x):,.{digits}f}"
    except Exception:
        return "—"


def _fmt_int(x):
    try:
        if x is None:
            return "—"
        return f"{int(x):,}"
    except Exception:
        return "—"


def main():
    set_page("總上組上架產能", icon="📦")

    # ===== Sidebar：只放參數 =====
    with st.sidebar:
        st.header("⚙️ 參數設定")
        params = render_params()

    # ===== 主畫面中央：上傳檔案 =====
    st.markdown("## 📤 上傳資料檔案")
    st.caption("請上傳上架作業資料（Excel / CSV）。上傳後按『開始計算』即可產出 KPI、圖表與下載報表。")

    uploaded = st.file_uploader(
        "請上傳上架資料",
        type=["xlsx", "xlsm", "xls", "xlsb", "csv"],
        label_visibility="collapsed",
    )

    run_clicked = st.button(
        "🚀 開始計算",
        type="primary",
        disabled=(uploaded is None),
    )

    if not run_clicked:
        st.info("請先上傳檔案，再點『開始計算』。")
        return

    # ===== 計算 =====
    with st.spinner("計算中..."):
        result = run_shelf_efficiency(uploaded.getvalue(), uploaded.name, params)

    summary_df = result.get("summary_df", pd.DataFrame())
    ampm_df = result.get("ampm_df", pd.DataFrame())
    detail_df = result.get("detail_df", pd.DataFrame())

    target = float(result.get("target_eff", params.get("target_eff", 20.0)))

    # ===== KPI =====
    kpis = [
        KPI("人數", _fmt_int(result.get("people"))),
        KPI("總筆數", _fmt_int(result.get("total_count"))),
        KPI("總工時", _fmt_num(result.get("total_hours"))),
        KPI("平均效率", _fmt_num(result.get("avg_eff"))),
        KPI("達標率", result.get("pass_rate", "—")),
    ]
    st.divider()
    render_kpis(kpis)
    st.divider()

    # ===== 圖表：左效率排行 / 右 AM-PM 對照 =====
    left, right = st.columns([1.2, 1])

    with left:
        if isinstance(summary_df, pd.DataFrame) and not summary_df.empty:
            x_col = "姓名" if "姓名" in summary_df.columns else summary_df.columns[0]
            y_col = "效率" if "效率" in summary_df.columns else summary_df.columns[-1]

            bar_topN(
                summary_df,
                x_col=x_col,
                y_col=y_col,
                hover_cols=[c for c in ["記錄輸入人", "筆數", "總分鐘"] if c in summary_df.columns],
                top_n=params["top_n"],
                target=target,
                title="全日效率排行（Top N）",
            )
        else:
            st.info("彙總資料為空，請確認檔案內容是否正確。")

    with right:
        pivot_am_pm(
            ampm_df,
            index_col="姓名",
            segment_col="時段",
            value_col="效率_件每小時",
            title="上午 vs 下午效率（平均）",
        )

    st.divider()

    # ===== 表格 =====
    table_block(
        summary_title="彙總表",
        summary_df=summary_df if isinstance(summary_df, pd.DataFrame) else pd.DataFrame(),
        detail_title="明細表（收合）",
        detail_df=detail_df if isinstance(detail_df, pd.DataFrame) else pd.DataFrame(),
        detail_expanded=False,
    )

    # ===== 下載 =====
    if result.get("xlsx_bytes"):
        download_excel(result["xlsx_bytes"], filename=result.get("xlsx_name", "上架績效.xlsx"))


if __name__ == "__main__":
    main()
