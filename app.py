import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime, time
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ==========================================
# 1. 核心邏輯封裝 (從您原始腳本移植)
# ==========================================

ID_TO_NAME = {
    "09440": "張予軒","10137": "徐嘉蔆","10818": "葉青芳","11797": "賴泉和",
    "20201109001": "吳振凱","10003": "李茂銓","10471": "余興炫","10275": "羅仲宇",
}

def map_name_from_id(x):
    s = str(x).strip() if x else ""
    if s in ID_TO_NAME: return ID_TO_NAME[s]
    return ID_TO_NAME.get(s.lstrip("0"), "")

def to_dt(series):
    return pd.to_datetime(series, errors="coerce")

def pick_col(cols, candidates):
    cols_norm = [str(c).strip() for c in cols]
    for cand in candidates:
        if cand in cols_norm: return cand
    return None

# (這裡需包含您的 annotate_idle, build_efficiency_table_full, build_efficiency_table_ampm 等函式內容)
# 為了回應簡潔，此處假設您已將原始腳本中的運算邏輯放入以下函式：

def process_data(uploaded_file, skip_rules, threshold_min):
    # 讀取資料
    sheets = pd.read_excel(uploaded_file, sheet_name=None)
    
    # 這裡放入您原始腳本 main() 中的處理邏輯
    # 遍歷各分頁 -> 篩選 QC -> 計算空窗 -> 產生 full_df 與 ampm_df
    
    # --- 模擬計算結果 (請在此處填入您的計算代碼) ---
    # processed_sheets, full_df, ampm_df, idle_details = your_v18_logic(...)
    
    return None, pd.DataFrame(), pd.DataFrame(), pd.DataFrame() # 暫代回傳

# ==========================================
# 2. Streamlit UI 介面
# ==========================================

st.set_page_config(page_title="驗收效率分析系統", layout="wide")
st.title("🚀 驗收達標分析系統 v18")

# --- 側邊欄參數 ---
with st.sidebar:
    st.header("⚙️ 參數設定")
    threshold_min = st.number_input("空窗門檻 (分鐘)", value=10)
    
    st.subheader("🚫 排除規則")
    if 'rules' not in st.session_state:
        st.session_state.rules = []
    
    rule_u = st.text_input("人員編號 (留空代表全員)")
    col_t1, col_t2 = st.columns(2)
    rule_s = col_t1.text_input("開始", value="15:00")
    rule_e = col_t2.text_input("結束", value="16:00")
    
    if st.button("➕ 新增規則"):
        try:
            st.session_state.rules.append({
                "user": rule_u,
                "t_start": datetime.strptime(rule_s, "%H:%M").time(),
                "t_end": datetime.strptime(rule_e, "%H:%M").time()
            })
        except: st.error("時間格式錯誤")

    if st.session_state.rules:
        for i, r in enumerate(st.session_state.rules):
            st.caption(f"{i+1}. {r['user'] or '全員'}: {r['t_start']}~{r['t_end']}")
        if st.button("🗑️ 清空規則"):
            st.session_state.rules = []
            st.rerun()

# --- 主程式區塊 ---
uploaded_file = st.file_uploader("上傳 Excel 檔案", type=["xlsx"])

if uploaded_file:
    # 執行運算並獲取結果
    # 注意：這裡就是定義 full_df 的地方！
    with st.spinner('計算中...'):
        # 這裡應該呼叫您整合好的運算邏輯
        # processed, full_df, ampm_df, idle_details = run_v18_engine(uploaded_file, st.session_state.rules, threshold_min)
        
        # 暫時用 dummy 資料確保 APP 不報錯
        full_df = pd.DataFrame([{"日期": "2023-01-01", "姓名": "測試人員", "效率": 25.5}]) 
        ampm_df = pd.DataFrame() 
        idle_details = pd.DataFrame()

    # --- 數據呈現 ---
    st.success("計算完畢")
    
    col_m1, col_m2 = st.columns(2)
    col_m1.metric("總處理筆數", len(full_df))
    
    tab1, tab2, tab3 = st.tabs(["全日統計", "AM/PM 統計", "空窗明細"])
    
    with tab1:
        # 現在 full_df 已經定義，不會報 NameError
        st.subheader("記錄輸入人統計 (全日)")
        st.dataframe(full_df.style.background_gradient(subset=['效率'], cmap='RdYlGn') if not full_df.empty else full_df)

    # --- 下載區 ---
    st.divider()
    st.subheader("📥 下載報表")
    # 建立下載用的 Excel 串流 (BytesIO)
    # ...