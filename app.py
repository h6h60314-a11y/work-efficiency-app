import streamlit as st
import pandas as pd
import numpy as np
import io
from datetime import datetime, time
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

# --- 直接移植您原有的邏輯函數 (calc_rest_minutes, annotate_idle 等) ---
# (為了篇幅，這裡僅展示結構改動，邏輯部分保持不變)

# ===== 原有參數與對照表 =====
ID_TO_NAME = {
    "09440": "張予軒","10137": "徐嘉蔆","10818": "葉青芳","11797": "賴泉和",
    "20201109001": "吳振凱","10003": "李茂銓","10471": "余興炫","10275": "羅仲宇",
    "9440": "張予軒",
}

def map_name_from_id(x):
    s = str(x).strip() if x else ""
    return ID_TO_NAME.get(s, ID_TO_NAME.get(s.lstrip("0"), ""))

# ... 此處包含您原本所有的 calc_exclude_minutes_for_range, annotate_idle, 
# build_efficiency_table_full, build_efficiency_table_ampm 等函數 ...
# (請將您原本 .py 檔中的函數內容貼到這裡)

# --- Streamlit 介面設計 ---
st.set_page_config(page_title="驗收達標效率統計 v18", layout="wide")

st.title("📊 驗收達標效率統計系統 (Streamlit 版)")
st.markdown("此版本支援 **多名人員排除規則** 與 **AM/PM 自動分段**")

# 側邊欄：設定參數
with st.sidebar:
    st.header("⚙️ 參數設定")
    threshold = st.number_input("空窗門檻 (分鐘)", value=10)
    
    st.subheader("🚫 排除規則設定")
    rule_input = st.text_area("格式：人員ID,HH:MM,HH:MM (每行一筆)", 
                              placeholder="20201109001,15:00,16:00\n,09:00,10:00")
    
    skip_rules = []
    if rule_input:
        for line in rule_input.split('\n'):
            parts = line.replace("，", ",").split(",")
            if len(parts) == 3:
                try:
                    skip_rules.append({
                        "user": parts[0].strip(),
                        "t_start": datetime.strptime(parts[1].strip(), "%H:%M").time(),
                        "t_end": datetime.strptime(parts[2].strip(), "%H:%M").time()
                    })
                except: pass

# 檔案上傳
uploaded_file = st.file_uploader("請上傳 Excel 或 CSV 檔案", type=["xlsx", "xls", "csv"])

if uploaded_file:
    # 讀取檔案
    if uploaded_file.name.endswith('.csv'):
        df_dict = {"Sheet1": pd.read_csv(uploaded_file)}
    else:
        df_dict = pd.read_excel(uploaded_file, sheet_name=None)

    if st.button("🚀 開始計算"):
        # 執行您原本的處理邏輯
        # ... (調用處理函數) ...
        
        # 建立下載用的 BytesIO
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # (執行您原本的 df.to_excel 邏輯)
            # ...
            st.success("✅ 計算完成！")
            
        st.download_button(
            label="💾 下載分析報表",
            data=output.getvalue(),
            file_name=f"分析結果_{datetime.now().strftime('%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
