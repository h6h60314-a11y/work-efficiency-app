import streamlit as st
import pandas as pd
from supabase import create_client
from postgrest.exceptions import APIError

from common_ui import inject_logistics_theme, set_page, card_open, card_close


def sb():
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_SERVICE_ROLE_KEY"])


def _human_api_error(e: Exception) -> str:
    try:
        if hasattr(e, "args") and e.args:
            return str(e.args[0])
    except Exception:
        pass
    return str(e)


def self_check():
    card_open("🧪 資料庫連線狀態（Supabase）")
    st.write("SUPABASE_URL：", (st.secrets.get("SUPABASE_URL", "")[:40] + "...") if st.secrets.get("SUPABASE_URL") else "（未設定）")
    st.write("SUPABASE_BUCKET：", st.secrets.get("SUPABASE_BUCKET", "work-efficiency-exports"))
    st.write("KEY 前綴：", (st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")[:12] + "...") if st.secrets.get("SUPABASE_SERVICE_ROLE_KEY") else "（未設定）")
    try:
        _ = sb().schema("public").table("audit_runs").select("id,created_at").limit(1).execute()
        st.success("✅ audit_runs 可讀取（連線/權限/表名 OK）")
    except APIError as e:
        st.error("❌ 讀取 audit_runs 失敗")
        st.code(_human_api_error(e))
        st.stop()
    card_close()


def _rate_light(x: float | None):
    # 你可調整門檻：>=85% 綠、70-85 黃、<70 紅
    if x is None:
        return ("—", "⚪")
    try:
        x = float(x)
    except Exception:
        return ("—", "⚪")

    if x >= 0.85:
        return (f"{x:.0%}", "🟢")
    if x >= 0.70:
        return (f"{x:.0%}", "🟡")
    return (f"{x:.0%}", "🔴")


def download_from_storage(object_path: str) -> bytes:
    client = sb()
    bucket = st.secrets.get("SUPABASE_BUCKET", "work-efficiency-exports")
    return client.storage.from_(bucket).download(object_path)


def main():
    inject_logistics_theme()
    set_page("營運稽核與復盤中心", icon="📊")
    st.caption("歷次分析留存｜AM/PM 班 KPI｜達標燈號｜下載留存報表")

    self_check()

    rows = (
        sb()
        .schema("public")
        .table("audit_runs")
        .select("*")
        .order("created_at", desc=True)
        .limit(2000)
        .execute()
        .data
        or []
    )

    if not rows:
        st.info("目前 audit_runs 沒有任何留存紀錄。請先跑一次模組並確認「稽核留存狀態」成功。")
        return

    df = pd.DataFrame(rows)
    df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce")

    # Sidebar filters
    with st.sidebar:
        st.header("🔎 查詢條件（管理用）")
        min_d = df["created_at"].dt.date.min()
        max_d = df["created_at"].dt.date.max()
        date_range = st.date_input("分析日期區間", value=(min_d, max_d))

        ops = sorted([x for x in df.get("operator", pd.Series([])).dropna().unique()])
        operator = st.selectbox("分析執行人（Operator）", ["全部"] + ops)

        apps = sorted([x for x in df.get("app_name", pd.Series([])).dropna().unique()])
        app_name = st.selectbox("模組別", ["全部"] + apps)

    mask = (df["created_at"].dt.date >= date_range[0]) & (df["created_at"].dt.date <= date_range[1])
    if operator != "全部":
        mask &= df["operator"] == operator
    if app_name != "全部":
        mask &= df["app_name"] == app_name

    df_f = df[mask].copy()
    if df_f.empty:
        st.warning("篩選後沒有資料。")
        return

    # KPI trend (avg_eff)
    card_open("📈 KPI 趨勢（AM 班 vs PM 班）")
    trend = []
    for _, r in df_f.iterrows():
        for k, label in [("kpi_am", "AM 班"), ("kpi_pm", "PM 班")]:
            obj = r.get(k) or {}
            trend.append(
                {
                    "分析時間": r["created_at"],
                    "班別": label,
                    "平均效率": obj.get("avg_eff"),
                    "達標率": obj.get("pass_rate"),
                }
            )
    tdf = pd.DataFrame(trend).dropna(subset=["分析時間"])
    st.line_chart(tdf, x="分析時間", y="平均效率", color="班別")
    card_close()

    # Runs table with lights
    card_open("📄 歷次分析留存紀錄（含達標燈號）")

    def _light_for(row, key):
        obj = row.get(key) or {}
        rate = obj.get("pass_rate")
        pct, lamp = _rate_light(rate)
        return f"{lamp} {pct}"

    df_f["AM達標"] = df_f.apply(lambda r: _light_for(r, "kpi_am"), axis=1)
    df_f["PM達標"] = df_f.apply(lambda r: _light_for(r, "kpi_pm"), axis=1)

    show_cols = ["created_at", "app_name", "operator", "source_filename", "AM達標", "PM達標", "id", "export_object_path"]
    for c in show_cols:
        if c not in df_f.columns:
            df_f[c] = None

    st.dataframe(
        df_f[show_cols].rename(
            columns={
                "created_at": "分析時間",
                "app_name": "模組別",
                "operator": "分析執行人",
                "source_filename": "來源檔案",
                "id": "紀錄ID",
                "export_object_path": "報表留存路徑",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )
    card_close()

    # Download selected
    card_open("⬇️ 下載當次 KPI 報表（留存）")
    idxs = df_f.index.tolist()
    selected = st.selectbox(
        "選擇紀錄",
        options=idxs,
        format_func=lambda i: f"{df_f.loc[i,'created_at']}｜{df_f.loc[i,'app_name']}｜{df_f.loc[i,'source_filename']}",
    )

    obj_path = df_f.loc[selected].get("export_object_path")
    if obj_path:
        if st.button("準備下載"):
            try:
                content = download_from_storage(obj_path)
                st.download_button(
                    "點此下載 Excel",
                    data=content,
                    file_name=obj_path.split("/")[-1],
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            except Exception as e:
                st.error("下載失敗")
                st.code(repr(e))
    else:
        st.warning("此筆紀錄未留存 Excel（export_object_path 為空）。")
    card_close()


if __name__ == "__main__":
    main()
