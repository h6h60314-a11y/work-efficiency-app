"""
共用 UI 模板（A：主管快速看版）
- 固定：Sidebar（上傳/參數/開始/下載）
- 固定：KPI 卡片（5 張）
- 固定：左右雙欄圖表（效率 Top N / 問題導向圖）
- 固定：彙總表（展開）+ 明細表（收合）
"""
from __future__ import annotations

from dataclasses import dataclass
from typing import Callable, Optional, Dict, Any, List, Tuple

import pandas as pd
import streamlit as st

try:
    import plotly.express as px
except Exception:
    px = None


@dataclass
class KPI:
    label: str
    value: str
    help: str = ""


def set_page(title: str, icon: str = "📊"):
    st.set_page_config(page_title=title, layout="wide")
    st.title(f"{icon} {title}")


def sidebar_uploader_and_actions(
    *,
    file_types: List[str],
    params_renderer: Callable[[], Dict[str, Any]],
    run_label: str = "🚀 開始計算",
    clear_label: str = "🧹 清空參數",
) -> Tuple[Optional[Any], Dict[str, Any], bool]:
    """
    回傳：uploaded_file, params(dict), run_clicked(bool)
    """
    with st.sidebar:
        st.header("操作")
        uploaded = st.file_uploader("📤 上傳檔案", type=file_types)

        st.divider()
        st.subheader("⚙️ 參數")
        if st.button(clear_label):
            # 交給 params_renderer 自己用 session_state 管理；這裡只觸發 rerun
            st.rerun()
        params = params_renderer() or {}

        st.divider()
        run_clicked = st.button(run_label, disabled=(uploaded is None))

    return uploaded, params, run_clicked


def render_kpis(kpis: List[KPI]):
    cols = st.columns(len(kpis))
    for c, k in zip(cols, kpis):
        if k.help:
            c.metric(k.label, k.value, help=k.help)
        else:
            c.metric(k.label, k.value)


def _color_pass_fail(series: pd.Series, target: float) -> pd.Series:
    # 回傳 '達標' / '未達標' / '—'
    def f(x):
        try:
            if pd.isna(x):
                return "—"
            return "達標" if float(x) >= target else "未達標"
        except Exception:
            return "—"
    return series.apply(f)


def bar_topN(
    df: pd.DataFrame,
    *,
    x_col: str,
    y_col: str,
    hover_cols: List[str],
    top_n: int = 30,
    target: float = 20.0,
    title: str = "效率排行（Top N）",
):
    st.subheader(title)
    if df is None or df.empty or y_col not in df.columns:
        st.info("沒有可顯示的資料。")
        return

    view = df.copy()
    view = view.sort_values(y_col, ascending=False).head(top_n)

    if px is None:
        st.warning("plotly 未安裝：改用表格呈現（如需圖表請在 requirements.txt 加上 plotly）。")
        st.dataframe(view[[x_col, y_col] + [c for c in hover_cols if c in view.columns]], use_container_width=True)
        return

    view["_達標狀態"] = _color_pass_fail(view[y_col], target)
    fig = px.bar(
        view,
        x=x_col,
        y=y_col,
        color="_達標狀態",
        hover_data=[c for c in hover_cols if c in view.columns],
    )
    st.plotly_chart(fig, use_container_width=True)


def pivot_am_pm(
    ampm_df: pd.DataFrame,
    *,
    index_col: str = "姓名",
    segment_col: str = "時段",
    value_col: str = "效率",
    title: str = "上午 vs 下午效率（平均）",
):
    st.subheader(title)
    if ampm_df is None or ampm_df.empty:
        st.info("沒有 AM/PM 資料。")
        return
    need = {index_col, segment_col, value_col}
    if not need.issubset(set(ampm_df.columns)):
        st.info("AM/PM 欄位不足，無法製作對照。")
        return
    pivot = ampm_df.pivot_table(index=index_col, columns=segment_col, values=value_col, aggfunc="mean").reset_index()
    st.dataframe(pivot, use_container_width=True)


def table_block(
    *,
    summary_title: str,
    summary_df: pd.DataFrame,
    detail_title: str,
    detail_df: pd.DataFrame,
    detail_expanded: bool = False,
):
    st.subheader(summary_title)
    st.dataframe(summary_df, use_container_width=True)

    with st.expander(detail_title, expanded=detail_expanded):
        st.dataframe(detail_df, use_container_width=True)


def download_excel(xlsx_bytes: bytes, filename: str):
    st.download_button(
        "⬇️ 下載 Excel 結果",
        data=xlsx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
