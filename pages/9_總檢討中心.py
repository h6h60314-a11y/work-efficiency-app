import streamlit as st
import pandas as pd
from supabase import create_client
from postgrest.exceptions import APIError

from common_ui import set_page, card_open, card_close


def sb():
    url = st.secrets.get("SUPABASE_URL", "")
    key = st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")
    if not url or not key:
        st.error("缺少 Secrets：SUPABASE_URL 或 SUPABASE_SERVICE_ROLE_KEY")
        st.stop()
    return create_client(url, key)


def _human_api_error(e: Exception) -> str:
    try:
        if hasattr(e, "args") and e.args:
            return str(e.args[0])
    except Exception:
        pass
    return str(e)


def load_audit_runs_no_cache(limit: int = 1000):
    client = sb()
    return (
        client.schema("public")
        .table("audit_runs")
        .select("*")
        .order("created_at", desc=True)
        .limit(limit)
        .execute()
        .data
        or []
    )


def download_from_storage(object_path: str) -> bytes:
    client = sb()
    bucket = st.secrets.get("SUPABASE_BUCKET", "work-efficiency-exports")
    return client.storage.from_(bucket).download(object_path)


def self_check():
    card_open("🧪 Supabase 連線自檢")
    st.write(
        "SUPABASE_URL：",
        st.secrets.get("SUPABASE_URL", "")[:40] + "..."
        if st.secrets.get("SUPABASE_URL")
        else "（未設定）",
    )
    st.write("SUPABASE_BUCKET：", st.secrets.get("SUPABASE_BUCKET", "work-efficiency-exports"))
    st.write(
        "KEY 前綴：",
        (st.secrets.get("SUPABASE_SERVICE_ROLE_KEY", "")[:12] + "...")
        if st.secrets.get("SUPABASE_SERVICE_ROLE_KEY")
        else "（未設定）",
    )

    try:
        client = sb()
        _ = (
            client.schema("public")
            .table("audit_runs")
            .select("id,created_at")
            .limit(1)
            .execute()
        )
        st.success("✅ audit_runs 可讀取（連線/權限/表名 OK）")
    except APIError as e:
        st.error("❌ 讀取 audit_runs 失敗（通常是：表不存在 / 權限 / RLS / key 錯）")
        st.code(_human_api_error(e))
        st.stop()
    except Exception as e:
        st.error("❌ 連線失敗（通常是 URL/key 不對）")
        st.code(str(e))
        st.stop()

    card_close()


def main():
    set_page("總檢討中心", icon="📊")
    self_check()

    try:
        rows = load_audit_runs_no_cache(limit=1000)
    except APIError as e:
        st.error("讀取 audit_runs 時發生 APIError：")
        st.code(_human_api_error(e))
        st.stop()

    if not rows:
        st.info("目前 audit_runs 沒有任何紀錄。請先去『驗收達標效率』跑一次，確認有寫入。")
        return

    df = pd.DataFrame(rows)
    df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce")

    # ========== Filters ==========
    with st.sidebar:
        st.header("🔍 篩選條件")
        min_d = df["created_at"].dt.date.min()
        max_d = df["created_at"].dt.date.max()
        date_range = st.date_input("執行日期區間", value=(min_d, max_d))

        ops = sorted([x for x in df.get("operator", pd.Series([])).dropna().unique()])
        operator = st.selectbox("執行人", options=["全部"] + ops)

        apps = sorted([x for x in df.get("app_name", pd.Series([])).dropna().unique()])
        app_name = st.selectbox("功能", options=["全部"] + apps)

    mask = (df["created_at"].dt.date >= date_range[0]) & (df["created_at"].dt.date <= date_range[1])
    if operator != "全部":
        mask &= df["operator"] == operator
    if app_name != "全部":
        mask &= df["app_name"] == app_name

    df_f = df[mask].copy()
    if df_f.empty:
        st.warning("篩選後沒有資料。")
        return

    # ========== KPI Trend ==========
    card_open("📈 KPI 歷史趨勢（上午 vs 下午）")
    kpi_rows = []
    for _, r in df_f.iterrows():
        for seg in ["am", "pm"]:
            k = r.get(f"kpi_{seg}") or {}
            kpi_rows.append(
                {
                    "時間": r["created_at"],
                    "時段": "上午" if seg == "am" else "下午",
                    "平均效率": k.get("avg_eff"),
                    "達標率": k.get("pass_rate"),
                    "總工時": k.get("total_hours"),
                    "總筆數": k.get("total_cnt"),
                    "人數": k.get("people"),
                }
            )
    kpi_df = pd.DataFrame(kpi_rows).sort_values("時間")
    st.line_chart(kpi_df, x="時間", y="平均效率", color="時段")
    card_close()

    # ========== Runs Table ==========
    card_open("📄 執行紀錄")
    show_cols = ["created_at", "operator", "source_filename", "app_name", "id", "export_object_path"]
    for c in show_cols:
        if c not in df_f.columns:
            df_f[c] = None

    st.dataframe(
        df_f[show_cols].rename(
            columns={
                "created_at": "執行時間",
                "operator": "執行人",
                "source_filename": "來源檔案",
                "app_name": "功能",
                "id": "紀錄ID",
                "export_object_path": "報表路徑",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )
    card_close()

    # ========== Download ==========
    card_open("⬇️ 下載歷史報表（當次匯出 Excel）")
    idxs = df_f.index.tolist()
    selected = st.selectbox(
        "選擇一筆紀錄",
        options=idxs,
        format_func=lambda i: f"{df_f.loc[i,'created_at']}｜{df_f.loc[i,'source_filename']}",
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
            except APIError as e:
                st.error("下載 Storage 檔案失敗：")
                st.code(_human_api_error(e))
            except Exception as e:
                st.error("下載失敗：")
                st.code(str(e))
    else:
        st.warning("此筆紀錄沒有留存 Excel（export_object_path 為空）。")
    card_close()


if __name__ == "__main__":
    main()
