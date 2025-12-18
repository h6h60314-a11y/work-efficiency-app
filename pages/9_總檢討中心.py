import streamlit as st
import pandas as pd
import datetime as dt
from supabase import create_client

from common_ui import set_page, card_open, card_close


# ======================
# Supabase client
# ======================
def sb():
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_SERVICE_ROLE_KEY"],
    )


# ======================
# 讀取 audit_runs
# ======================
@st.cache_data(ttl=60)
def load_audit_runs():
    client = sb()
    res = (
        client.table("audit_runs")
        .select("*")
        .order("created_at", desc=True)
        .execute()
    )
    return res.data or []


# ======================
# 下載 Excel
# ======================
def download_from_storage(object_path: str):
    client = sb()
    bucket = st.secrets["SUPABASE_BUCKET"]
    return client.storage.from_(bucket).download(object_path)


# ======================
# Main
# ======================
def main():
    set_page("總檢討中心", icon="📊")

    rows = load_audit_runs()
    if not rows:
        st.info("目前尚無任何留存紀錄。")
        return

    df = pd.DataFrame(rows)
    df["created_at"] = pd.to_datetime(df["created_at"]).dt.tz_convert("Asia/Taipei")

    # ======================
    # Filters
    # ======================
    with st.sidebar:
        st.header("🔍 篩選條件")

        date_range = st.date_input(
            "執行日期區間",
            value=(
                df["created_at"].dt.date.min(),
                df["created_at"].dt.date.max(),
            ),
        )

        operator = st.selectbox(
            "執行人",
            options=["全部"] + sorted([x for x in df["operator"].dropna().unique()]),
        )

    mask = (df["created_at"].dt.date >= date_range[0]) & (
        df["created_at"].dt.date <= date_range[1]
    )
    if operator != "全部":
        mask &= df["operator"] == operator

    df = df[mask]

    # ======================
    # KPI 趨勢
    # ======================
    card_open("📈 KPI 歷史趨勢")

    kpi_rows = []
    for _, r in df.iterrows():
        for seg in ["am", "pm"]:
            k = r.get(f"kpi_{seg}") or {}
            kpi_rows.append(
                {
                    "時間": r["created_at"],
                    "時段": "上午" if seg == "am" else "下午",
                    "人數": k.get("people"),
                    "總筆數": k.get("total_cnt"),
                    "總工時": k.get("total_hours"),
                    "平均效率": k.get("avg_eff"),
                    "達標率": k.get("pass_rate"),
                }
            )

    kpi_df = pd.DataFrame(kpi_rows)

    st.line_chart(
        kpi_df,
        x="時間",
        y=["平均效率"],
        color="時段",
    )

    card_close()

    # ======================
    # 紀錄清單
    # ======================
    card_open("📄 執行紀錄")

    show_cols = [
        "created_at",
        "operator",
        "source_filename",
        "app_name",
    ]

    st.dataframe(
        df[show_cols].rename(
            columns={
                "created_at": "執行時間",
                "operator": "執行人",
                "source_filename": "來源檔案",
                "app_name": "功能",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )

    card_close()

    # ======================
    # 下載 Excel
    # ======================
    card_open("⬇️ 下載歷史報表")

    selected = st.selectbox(
        "選擇一筆紀錄",
        options=df.index,
        format_func=lambda i: f"{df.loc[i,'created_at']}｜{df.loc[i,'source_filename']}",
    )

    obj_path = df.loc[selected, "export_object_path"]
    if obj_path:
        if st.button("下載該次 Excel"):
            content = download_from_storage(obj_path)
            st.download_button(
                "點此下載",
                data=content,
                file_name=obj_path.split("/")[-1],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
    else:
        st.warning("此筆紀錄未留存 Excel。")

    card_close()


if __name__ == "__main__":
    main()
