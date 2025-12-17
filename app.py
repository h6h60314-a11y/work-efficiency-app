import streamlit as st
import pandas as pd
import numpy as np
import os
from datetime import datetime, time
from io import BytesIO
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.formatting.rule import FormulaRule
from openpyxl.utils import get_column_letter

# ==========================================
# 1. 核心邏輯 (移植自 v18 原始碼)
# ==========================================

ID_TO_NAME = {
    "09440": "張予軒","10137": "徐嘉蔆","10818": "葉青芳","11797": "賴泉和",
    "20201109001": "吳振凱","10003": "李茂銓","10471": "余興炫","10275": "羅仲宇",
    "9440": "張予軒",
}

THRESHOLD_MIN = 10
USER_COLS = ["記錄輸入人","建立人員","建立者","輸入人","建立者姓名","操作人員","建立人"]
TIME_COLS = ["修訂日期","更新日期","異動日期","修改日期","最後更新時間","時間戳記","Timestamp"]
DEST_COL = "到"; DEST_VALUE_QC = "QC"
AM_START, AM_END, PM_START = time(9, 0), time(12, 30), time(13, 30)
LUNCH_START, LUNCH_END = time(12, 30), time(13, 30)

def map_name_from_id(x):
    s = str(x).strip() if x else ""
    return ID_TO_NAME.get(s, ID_TO_NAME.get(s.lstrip("0"), ""))

def to_dt(series):
    return pd.to_datetime(series, errors="coerce")

def pick_col(cols, candidates):
    cols_norm = [str(c).strip() for c in cols]
    for cand in candidates:
        if cand in cols_norm: return cand
    return None

# 休息與空窗核心計算 (與原 v18 完全一致)
def calc_exclude_minutes_for_range(date_obj, user_id, first_ts, last_ts, skip_rules):
    if pd.isna(first_ts) or pd.isna(last_ts) or not skip_rules: return 0
    segs = []
    user_id_str = str(user_id).strip()
    for rule in skip_rules:
        rule_user = str(rule["user"]).strip()
        if rule_user and rule_user != user_id_str: continue
        s_dt = datetime.combine(date_obj, rule["t_start"])
        e_dt = datetime.combine(date_obj, rule["t_end"])
        left, right = max(first_ts, s_dt), min(last_ts, e_dt)
        if right > left: segs.append((left, right))
    if not segs: return 0
    segs.sort(key=lambda x: x[0])
    merged = [list(segs[0])]
    for s, e in segs[1:]:
        if s <= merged[-1][1]: merged[-1][1] = max(merged[-1][1], e)
        else: merged.append([s, e])
    return sum([(e - s).total_seconds() / 60.0 for s, e in merged])

def annotate_idle(qc_df, user_col, time_col, skip_rules=None):
    merged = qc_df.copy()
    for col in ["空窗分鐘","空窗旗標","空窗區間","午後空窗分鐘","午後空窗旗標","午後空窗區間"]:
        merged[col] = pd.NA
    tmp = merged[[user_col, time_col]].copy()
    tmp["_user"] = tmp[user_col].astype(str).str.strip()
    tmp["_dt"] = to_dt(tmp[time_col])
    tmp = tmp.loc[tmp["_dt"].notna()].sort_values(by=["_user","_dt"])
    tmp["_prev_dt"] = tmp.groupby("_user")["_dt"].shift(1)
    
    results = []
    for _, r in tmp.iterrows():
        prev_dt, cur_dt, user_id = r["_prev_dt"], r["_dt"], r["_user"]
        if pd.isna(prev_dt):
            results.append([np.nan, 0, "", np.nan, 0, ""])
            continue
        gap = (cur_dt - prev_dt).total_seconds() / 60.0
        segs = []
        if prev_dt.date() == cur_dt.date():
            l_s, l_e = datetime.combine(cur_dt.date(), LUNCH_START), datetime.combine(cur_dt.date(), LUNCH_END)
            if min(cur_dt, l_e) > max(prev_dt, l_s): segs.append((max(prev_dt, l_s), min(cur_dt, l_e)))
            for rule in (skip_rules or []):
                if rule["user"] and str(rule["user"]).strip() != user_id: continue
                r_s, r_e = datetime.combine(cur_dt.date(), rule["t_start"]), datetime.combine(cur_dt.date(), rule["t_end"])
                if min(cur_dt, r_e) > max(prev_dt, r_s): segs.append((max(prev_dt, r_s), min(cur_dt, r_e)))
        
        overlap = 0.0
        if segs:
            segs.sort(); m_seg = [list(segs[0])]
            for s, e in segs[1:]:
                if s <= m_seg[-1][1]: m_seg[-1][1] = max(m_seg[-1][1], e)
                else: m_seg.append([s, e])
            overlap = sum([(e - s).total_seconds() / 60.0 for s, e in m_seg])
        
        eff_gap = gap - overlap
        idle = [int(eff_gap), 1, f"{prev_dt.strftime('%H:%M')}~{cur_dt.strftime('%H:%M')}"] if eff_gap > THRESHOLD_MIN else [np.nan, 0, ""]
        pm_gap = eff_gap if (prev_dt.date() == cur_dt.date() and prev_dt.time() >= LUNCH_END) else 0
        pm = [int(pm_gap), 1, idle[2]] if pm_gap > THRESHOLD_MIN else [np.nan, 0, ""]
        results.append(idle + pm)
    
    merged.loc[tmp.index, ["空窗分鐘","空窗旗標","空窗區間","午後空窗分鐘","午後空窗旗標","午後空窗區間"]] = results
    return merged

# ... (其餘 build_efficiency_table_full, build_efficiency_table_ampm, write_grouped_ampm_sheet 邏輯同 v18) ...
# [此處為了執行需包含 v18 中所有 build_ 開頭的函式內容]

# ==========================================
# 2. Streamlit 介面與流程
# ==========================================

st.set_page_config(page_title="驗收分析系統", layout="wide")
st.title("🚀 驗收達標效率分析 v18")

if 'rules' not in st.session_state: st.session_state.rules = []

with st.sidebar:
    st.header("⚙️ 排除規則")
    with st.form("rule_form", clear_on_submit=True):
        u = st.text_input("人員編號")
        s, e = st.text_input("開始(HH:MM)", "15:00"), st.text_input("結束(HH:MM)", "16:00")
        if st.form_submit_button("新增規則"):
            try: st.session_state.rules.append({"user": u, "t_start": datetime.strptime(s, "%H:%M").time(), "t_end": datetime.strptime(e, "%H:%M").time()})
            except: st.error("格式錯誤")
    if st.button("清空規則"): st.session_state.rules = []; st.rerun()

uploaded_file = st.file_uploader("上傳 Excel", type=["xlsx"])

if uploaded_file:
    sheets = pd.read_excel(uploaded_file, sheet_name=None)
    processed = {}
    
    # 比照 v18 main() 邏輯處理每個 Sheet
    for name, df in sheets.items():
        if df.empty: continue
        # ... (執行 QC 篩選、時間轉換、空窗標記)
        # 這裡需完整實現您 v18 腳本中 main() 內的 df 處理邏輯
        processed[name] = df # 示意

    if processed:
        # 產出下載檔案
        output = BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            # 確保至少有一個 Sheet 寫入 [此處修正 IndexError]
            for name, pdf in processed.items():
                pdf.to_excel(writer, sheet_name=name[:31], index=False)
            
            # 寫入統計頁面 (build_efficiency_table_full 等結果)
            # ampm_df.to_excel(writer, sheet_name="分段統計", index=False)
            
        st.success("分析完畢！")
        st.download_button("📥 下載結果報表", output.getvalue(), "分析結果.xlsx")
    else:
        st.warning("檔案中沒有可處理的資料。")
