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
    variant: str = "purple"  # purple/blue/green/orange/gray


def inject_purple_theme():
    st.markdown(
        """
<style>
/* ===== Purple SaaS Theme ===== */
.stApp {
    background: #F5F7FB;
}

/* 讓內容左右留白更像 Dashboard */
.block-container {
    padding-top: 1.25rem;
    padding-bottom: 2.5rem;
}

/* Sidebar 更乾淨 */
section[data-testid="stSidebar"] {
    background: #FFFFFF;
    border-right: 1px solid rgba(15, 23, 42, 0.06);
}

/* 全局卡片風格 */
.gt-card {
    background: #FFFFFF;
    border-radius: 18px;
    padding: 18px 18px;
    box-shadow: 0 12px 30px rgba(15, 23, 42, 0.06);
    border: 1px solid rgba(15, 23, 42, 0.04);
}

/* KPI 漸層卡 */
.gt-kpi {
    border-radius: 18px;
    padding: 18px 18px;
    color: #FFFFFF;
    box-shadow: 0 14px 34px rgba(15, 23, 42, 0.12);
    border: 1px solid rgba(255,255,255,0.18);
}

.gt-kpi .kpi-label {
    font-size: 0.92rem;
    opacity: 0.92;
    margin: 0 0 8px 0;
}

.gt-kpi .kpi-value {
    font-size: 1.7rem;
    font-weight: 800;
    letter-spacing: 0.2px;
    margin: 0;
    line-height: 1.1;
}

.gt-kpi .kpi-help {
    font-size: 0.82rem;
    opacity: 0.85;
    margin-top: 10px;
}

/* variants */
.gt-purple { background: linear-gradient(135deg, #8E2DE2, #4A00E0); }
.gt-blue   { background: linear-gradient(135deg, #36D1DC, #5B86E5); }
.gt-green  { background: linear-gradient(135deg, #43E97B, #38F9D7); }
.gt-orange { background: linear-gradient(135deg, #FF512F, #DD2476); }
.gt-gray   { background: linear-gradient(135deg, #64748B, #334155); }

/* 讓按鈕更像 SaaS */
.stButton>button {
    border-radius: 12px !important;
    padding: 0.6rem 1rem !important;
    font-weight: 700 !important;
}

.stDownloadButton>button {
    border-radius: 12px !important;
    padding: 0.6rem 1rem !important;
    font-weight: 700 !important;
}

/* dataframes 外觀更乾淨 */
[data-testid="stDataFrame"] {
    background: white;
    border-radius: 16px;
    border: 1px solid rgba(15, 23, 42, 0.05);
    overflow: hidden;
}

/* Plotly 圖表外框像卡片 */
[data-testid="stPlotlyChart"] > div {
    background: white;
    border-radius: 16px;
    border: 1px solid rgba(15, 23, 42, 0.05);
    box-shadow: 0 12px 30px rgba(15, 23, 42, 0.06);
    padding: 10px;
}
</style>
""",
        unsafe_allow_html=True,
    )


def set_page(title: str, icon: str = "📊"):
    st.set_page_config(page_title=title, layout="wide", page_icon=icon)
    inject_purple_theme()
    st.title(f"{icon} {title}")


def sidebar_params_only(params_renderer: Callable[[], Dict[str, Any]]) -> Dict[str, Any]:
    with st.sidebar:
        st.header("⚙️ 參數設定")
        return params_renderer() or {}


def render_kpis(kpis: List[KPI]):
    if not kpis:
        return

    cols = st.columns(len(kpis))
    for c, k in zip(cols, kpis):
        variant = (k.variant or "purple").strip().lower()
        if variant not in {"purple", "blue", "green", "orange", "gray"}:
            variant = "purple"

        help_text = k.help.strip() if k.help else ""
        help_html = f'<div class="kpi-help">{help_text}</div>' if help_text else ""

        c.markdown(
            f"""
<div class="gt-kpi gt-{variant}">
  <div class="kpi-label">{k.label}</div>
  <div class="kpi-value">{k.value}</div>
  {help_html}
</div>
""",
            unsafe_allow_html=True,
        )


def _pass_fail(series: pd.Series, target: float) -> pd.Series:
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
    view = view.sort_values(y_col, ascending=False).head(int(top_n))

    if px is None:
        st.warning("plotly 未安裝：改用表格呈現（如需圖表請在 requirements.txt 加上 plotly）。")
        st.dataframe(view[[x_col, y_col] + [c for c in hover_cols if c in view.columns]], use_container_width=True)
        return

    view["_達標狀態"] = _pass_fail(view[y_col], target)
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
    st.markdown('<div class="gt-card">', unsafe_allow_html=True)
    st.dataframe(summary_df, use_container_width=True)
    st.markdown("</div>", unsafe_allow_html=True)

    with st.expander(detail_title, expanded=detail_expanded):
        st.markdown('<div class="gt-card">', unsafe_allow_html=True)
        st.dataframe(detail_df, use_container_width=True)
        st.markdown("</div>", unsafe_allow_html=True)


def download_excel(xlsx_bytes: bytes, filename: str):
    st.download_button(
        "⬇️ 下載 Excel 結果",
        data=xlsx_bytes,
        file_name=filename,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
