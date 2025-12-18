import streamlit as st
import pandas as pd
from supabase import create_client

from common_ui import inject_logistics_theme, set_page, card_open, card_close


def sb():
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_SERVICE_ROLE_KEY"])


def main():
    inject_logistics_theme()
    set_page("人員 AM/PM 對比檢討", icon="🧑‍💼")
    st.caption("主管檢討｜以留存紀錄為基礎｜比較 AM / PM 班 KPI 趨勢")

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
        st.info("目前沒有留存紀錄。")
        return

    df = pd.DataFrame(rows)
    df["created_at"] = pd.to_datetime(df["created_at"], errors="coerce")

    with st.sidebar:
        st.header("🔎 檢討條件")
        apps = sorted([x for x in df.get("app_name", pd.Series([])).dropna().unique()])
        app_name = st.selectbox("模組別", apps)

        # 這頁以「分析執行人」角度作對比（若你要改成「作業人員」，需要把作業人員明細存入 DB）
        ops = sorted([x for x in df.get("operator", pd.Series([])).dropna().unique()])
        operator = st.selectbox("分析執行人（Operator）", ["全部"] + ops)

    dff = df[df["app_name"] == app_name].copy()
    if operator != "全部":
        dff = dff[dff["operator"] == operator].copy()

    if dff.empty:
        st.warning("篩選後沒有資料")
        return

    # Build trend
    trend = []
    for _, r in dff.iterrows():
        for k, label in [("kpi_am", "AM 班"), ("kpi_pm", "PM 班")]:
            obj = r.get(k) or {}
            trend.append(
                {
                    "分析時間": r["created_at"],
                    "班別": label,
                    "平均效率": obj.get("avg_eff"),
                    "達標率": obj.get("pass_rate"),
                    "來源檔案": r.get("source_filename"),
                }
            )
    tdf = pd.DataFrame(trend).dropna(subset=["分析時間"])

    card_open("📈 AM / PM 平均效率趨勢")
    st.line_chart(tdf, x="分析時間", y="平均效率", color="班別")
    card_close()

    card_open("📄 歷次留存（摘要）")
    st.dataframe(
        tdf.sort_values("分析時間", ascending=False),
        use_container_width=True,
        hide_index=True,
    )
    card_close()

    st.info("如果你要做到『作業人員』AM/PM 對比（而非 Operator），我可以下一步把「每次計算的彙總人員表」也寫入 Supabase，這樣主管就能指定某位作業員做長期對比。")


if __name__ == "__main__":
    main()
