import io
import uuid
import datetime as dt

import streamlit as st
import pandas as pd

from common_ui import (
    inject_logistics_theme,
    set_page,
    KPI,
    render_kpis,
    bar_topN,
    card_open,
    card_close,
    download_excel,
)

from audit_store import sha256_bytes, upload_export_bytes, insert_audit_run


def _read_any(uploaded):
    name = (uploaded.name or "").lower()
    b = uploaded.getvalue()
    if name.endswith(".csv"):
        return pd.read_csv(io.BytesIO(b))
    return pd.read_excel(io.BytesIO(b))


def _to_excel_bytes(df: pd.DataFrame, sheet_name="Putaway KPI"):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
    return out.getvalue()


def _kpi_pack(df: pd.DataFrame, target: float):
    if df is None or df.empty:
        return {"people": 0, "total_cnt": None, "total_hours": None, "avg_eff": None, "pass_rate": None}
    return {
        "people": int(df["姓名"].nunique()) if "姓名" in df.columns else int(len(df)),
        "total_cnt": float(df["箱數"].sum()) if "箱數" in df.columns else None,
        "total_hours": float(df["工時"].sum()) if "工時" in df.columns else None,
        "avg_eff": float(df["效率"].mean()) if "效率" in df.columns else None,
        "pass_rate": float((df["效率"] >= target).mean()) if "效率" in df.columns else None,
    }


def main():
    inject_logistics_theme()
    set_page("上架產能分析（Putaway KPI）", icon="📦")
    st.caption("上架作業｜Putaway｜人時效率｜AM / PM 班別｜稽核留存")

    with st.sidebar:
        st.header("⚙️ 計算條件設定")
        operator = st.text_input("分析執行人（Operator）")
        top_n = st.number_input("效率排行顯示人數（Top N）", 10, 100, 30, step=5)
        target = st.number_input("目標效率（KPI Target）", value=20.0, step=1.0)

        st.divider()
        st.caption("資料欄位需求（至少）：班別/姓名/箱數/工時/效率")
        st.caption("班別可為：AM/PM 或 上午/下午（系統會自動轉換）")

    card_open("📤 上傳作業原始資料（上架）")
    uploaded = st.file_uploader(
        "上傳上架作業原始資料",
        type=["xlsx", "xls", "csv"],
        label_visibility="collapsed",
    )
    run = st.button("🚀 產出 KPI", type="primary", disabled=uploaded is None)
    card_close()

    if not run:
        st.info("請先上傳上架作業原始資料")
        return

    with st.spinner("KPI 計算中，請稍候..."):
        df = _read_any(uploaded)

    # 標準化欄位
    if "班別" not in df.columns:
        # 容錯：若使用「時段」
        if "時段" in df.columns:
            df["班別"] = df["時段"]
        else:
            st.error("資料缺少『班別』欄位（或『時段』欄位），無法切分 AM/PM")
            return

    df["班別"] = df["班別"].astype(str).replace({"上午": "AM", "下午": "PM", "AM 班": "AM", "PM 班": "PM"}).str.upper()

    need_cols = ["姓名", "箱數", "工時", "效率"]
    missing = [c for c in need_cols if c not in df.columns]
    if missing:
        st.error(f"資料缺少必要欄位：{missing}")
        return

    df["箱數"] = pd.to_numeric(df["箱數"], errors="coerce")
    df["工時"] = pd.to_numeric(df["工時"], errors="coerce")
    df["效率"] = pd.to_numeric(df["效率"], errors="coerce")
    df = df.dropna(subset=["效率"])

    am = df[df["班別"] == "AM"].copy()
    pm = df[df["班別"] == "PM"].copy()

    col_l, col_r = st.columns(2)

    def render_shift(title, sdf):
        if sdf.empty:
            st.warning(f"{title} 無資料")
            return

        card_open(f"{title} KPI")
        render_kpis(
            [
                KPI("人數", f"{sdf['姓名'].nunique():,}"),
                KPI("上架箱數", f"{sdf['箱數'].sum():,.0f}"),
                KPI("總工時", f"{sdf['工時'].sum():,.2f}"),
                KPI("平均效率", f"{sdf['效率'].mean():,.2f}"),
                KPI("達標率", f"{(sdf['效率'] >= float(target)).mean():.0%}"),
            ]
        )
        card_close()

        card_open(f"{title} 效率排行（Top {int(top_n)}）")
        bar_topN(
            sdf.groupby("姓名", as_index=False).agg(效率=("效率", "mean"), 箱數=("箱數", "sum"), 工時=("工時", "sum")),
            x_col="姓名",
            y_col="效率",
            hover_cols=["箱數", "工時"],
            top_n=int(top_n),
            target=float(target),
        )
        card_close()

    with col_l:
        render_shift("🌓 AM 班（上架）", am)
    with col_r:
        render_shift("🌙 PM 班（上架）", pm)

    # 匯出（本次計算結果）
    export_df = df.copy()
    export_df["班別"] = export_df["班別"].replace({"AM": "AM 班", "PM": "PM 班"})
    xlsx_bytes = _to_excel_bytes(export_df, sheet_name="Putaway_KPI")

    card_open("⬇️ 匯出 KPI 報表")
    download_excel(xlsx_bytes, "上架產能_Putaway_KPI.xlsx")
    card_close()

    # 稽核留存（DB + Storage）
    st.divider()
    st.subheader("🧾 稽核留存狀態")
    try:
        export_path = upload_export_bytes(
            content=xlsx_bytes,
            object_path=f"putaway_runs/{dt.datetime.now():%Y%m%d}/{uuid.uuid4().hex}_putaway.xlsx",
        )

        payload = {
            "app_name": "上架產能分析（Putaway KPI）",
            "operator": operator or None,
            "source_filename": uploaded.name,
            "source_sha256": sha256_bytes(uploaded.getvalue()),
            "params": {"top_n": int(top_n), "target_eff": float(target)},
            "kpi_am": _kpi_pack(am, float(target)),
            "kpi_pm": _kpi_pack(pm, float(target)),
            "export_object_path": export_path,
        }

        row = insert_audit_run(payload)
        st.success(f"✅ 已成功留存本次分析（ID：{row.get('id','')}）")

    except Exception as e:
        st.error("❌ 稽核留存發生錯誤")
        st.code(repr(e))


if __name__ == "__main__":
    main()
