import streamlit as st
import pandas as pd
import plotly.express as px
from qc_core import run_qc_efficiency

st.set_page_config(page_title="驗收達標可視化", layout="wide")
st.title("📦 驗收達標效率看板")

uploaded = st.file_uploader("上傳 Excel/CSV", type=["xlsx","xlsm","xls","csv","txt"])

if "skip_rules" not in st.session_state:
    st.session_state.skip_rules = []

with st.sidebar:
    st.header("排除規則（不納入統計/不算空窗/扣總分鐘）")
    user = st.text_input("記錄輸入人（可空白=全員）", "")
    t1 = st.time_input("開始時間")
    t2 = st.time_input("結束時間")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("➕ 加入"):
            if t2 < t1:
                st.error("結束時間需 >= 開始時間")
            else:
                st.session_state.skip_rules.append({"user": user.strip(), "t_start": t1, "t_end": t2})
    with c2:
        if st.button("🧹 清空"):
            st.session_state.skip_rules = []

    if st.session_state.skip_rules:
        st.dataframe(pd.DataFrame(st.session_state.skip_rules), use_container_width=True)

if st.button("🚀 開始計算", disabled=(uploaded is None)) and uploaded:
    with st.spinner("計算中..."):
        result = run_qc_efficiency(uploaded.getvalue(), uploaded.name, st.session_state.skip_rules)

    full_df = result["full_df"]
    ampm_df = result["ampm_df"]

    st.subheader("全日效率排行")
    if not full_df.empty:
        fig = px.bar(full_df.sort_values("效率", ascending=False).head(30), x="姓名", y="效率")
        st.plotly_chart(fig, use_container_width=True)
        st.dataframe(full_df, use_container_width=True)

    st.subheader("上午/下午效率")
    if not ampm_df.empty:
        pivot = ampm_df.pivot_table(index="姓名", columns="時段", values="效率", aggfunc="mean").reset_index()
        st.dataframe(pivot, use_container_width=True)

    st.download_button(
        "⬇️ 下載 Excel 結果",
        data=result["xlsx_bytes"],
        file_name="驗收達標_含空窗_AMPM.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
