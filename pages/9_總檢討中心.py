import streamlit as st
import pandas as pd
import datetime as dt
from supabase import create_client
from postgrest.exceptions import APIError

from common_ui import inject_logistics_theme, set_page, card_open, card_close


# ========= Supabase =========
def sb():
    return create_client(
        st.secrets["SUPABASE_URL"],
        st.secrets["SUPABASE_SERVICE_ROLE_KEY"],
    )


# ========= Utilities =========
def _human_api_error(e: Exception) -> str:
    try:
        if hasattr(e, "args") and e.args:
            return str(e.args[0])
    except Exception:
        pass
    return str(e)


def current_delete_password():
    """
    依月份取得刪除密碼
    Key 格式：DELETE_PASSWORD_YYYYMM
    """
    ym = dt.datetime.now().strftime("%Y%m")
    key = f"DELETE_PASSWORD_{ym}"
    return key, st.secrets.get(key)


def download_from_storage(object_path: str) -> bytes:
    bucket = st.secrets.get("SUPABASE_BUCKET", "work-efficiency-exports")
    return sb().storage.from_(bucket).download(object_path)


def remove_from_storage(object_path: str):
    bucket = st.secrets.get("SUPABASE_BUCKET", "work-efficiency-exports")
    sb().storage.from_(bucket).remove([object_path])


def delete_audit_run(run_id: str):
    sb().schema("public").table("audit_runs").delete().eq("id", run_id).execute()


def _rate_light(x):
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


# ========= Page =========
def main():
    inject_logistics_theme()
    set_page("營運稽核與復盤中心", icon="📊")
    st.caption("歷次分析留存｜AM/PM KPI｜下載 / 刪除（每月輪替密碼）")

    # 取得當月密碼
    pwd_key, expected_pwd = current_delete_password()

    if not expected_pwd:
        st.error(
            f"❌ 尚未設定本月刪除密碼：{pwd_key}\n"
            "請至 Streamlit Secrets 設定後再使用刪除功能。"
        )
        st.stop()

    # 讀取資料
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
        st.info("目前尚無任何留存紀錄")
        return

    df = pd.DataFrame(rows)
    df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce")

    # ===== 表格 =====
    card_open("📄 歷次分析留存紀錄")

    def _light_for(row, key):
        obj = row.get(key) or {}
        rate = obj.get("pass_rate")
        pct, lamp = _rate_light(rate)
        return f"{lamp} {pct}"

    df["AM達標"] = df.apply(lambda r: _light_for(r, "kpi_am"), axis=1)
    df["PM達標"] = df.apply(lambda r: _light_for(r, "kpi_pm"), axis=1)

    st.dataframe(
        df[
            [
                "created_at",
                "app_name",
                "operator",
                "source_filename",
                "AM達標",
                "PM達標",
                "id",
            ]
        ].rename(
            columns={
                "created_at": "分析時間",
                "app_name": "模組別",
                "operator": "分析執行人",
                "source_filename": "來源檔案",
                "id": "紀錄ID",
            }
        ),
        use_container_width=True,
        hide_index=True,
    )
    card_close()

    # ===== 操作 =====
    card_open("🧰 紀錄操作（下載 / 刪除）")

    idx = st.selectbox(
        "選擇一筆紀錄",
        options=df.index.tolist(),
        format_func=lambda i: f"{df.loc[i,'created_at']}｜{df.loc[i,'app_name']}｜{df.loc[i,'source_filename']}",
    )

    run_id = df.loc[idx, "id"]
    obj_path = df.loc[idx, "export_object_path"]

    st.markdown(f"- **紀錄 ID**：`{run_id}`")
    st.markdown(f"- **本月刪除密碼 Key**：`{pwd_key}`")

    col1, col2 = st.columns(2)

    # 下載
    with col1:
        if obj_path and st.button("⬇️ 下載留存 Excel", use_container_width=True):
            content = download_from_storage(obj_path)
            st.download_button(
                "點此下載",
                data=content,
                file_name=obj_path.split("/")[-1],
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )

    # 刪除（每月輪替密碼）
    with col2:
        st.warning("⚠️ 刪除為不可逆操作（DB + Storage）")
        confirm = st.checkbox("我已確認要刪除此筆紀錄")
        keyword = st.text_input("輸入 DELETE 以解鎖", value="")
        pwd = st.text_input("輸入本月刪除密碼", type="password")

        unlocked = confirm and keyword.strip().upper() == "DELETE" and pwd == expected_pwd

        if st.button("🗑️ 刪除紀錄", disabled=not unlocked, type="primary", use_container_width=True):
            try:
                if obj_path:
                    remove_from_storage(obj_path)
                delete_audit_run(run_id)
                st.success("✅ 刪除完成（已套用當月密碼）")
                st.info("請重新整理頁面以更新清單")
            except APIError as e:
                st.error("❌ 刪除失敗（APIError）")
                st.code(_human_api_error(e))
            except Exception as e:
                st.error("❌ 刪除失敗")
                st.code(repr(e))

    card_close()


if __name__ == "__main__":
    main()
