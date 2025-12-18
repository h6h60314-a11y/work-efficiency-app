import io
import re
import uuid
import datetime as dt
from typing import Dict, List, Tuple, Optional

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


# =========================
# 依照你檔案規則（上架 v8.9）
# =========================
TO_EXCLUDE_KEYWORDS = ["CGS", "JCPL", "QC99", "GREAT0001X", "GX010", "PD99"]
TO_EXCLUDE_PATTERN = re.compile("|".join(re.escape(k) for k in TO_EXCLUDE_KEYWORDS), flags=re.IGNORECASE)

INPUT_USER_CANDIDATES = ["記錄輸入人", "記錄輸入者", "建立人", "輸入人"]
REV_DT_CANDIDATES = ["修訂日期", "修訂時間", "修訂日", "異動時間", "修改時間", "修訂日期時間", "修訂日期時間(系統)"]

TARGET_EFF = 20
IDLE_MIN_THRESHOLD = 10

AM_START, AM_END = dt.time(7, 0, 0), dt.time(12, 30, 0)
PM_START, PM_END = dt.time(13, 30, 0), dt.time(23, 59, 59)

NAME_MAP = {
    "20200924001": "黃雅君", "20210805001": "郭中合", "20220505002": "阮文青明",
    "20221221001": "阮文全", "20221222005": "謝忠龍", "20230119001": "陶春青",
    "20240926001": "陳莉娜", "20241011002": "林雙慧", "20250502001": "吳詩敏",
    "20250617001": "阮文譚", "20250617003": "喬家寶", "20250901009": "張寶萱",
    "G01": "0", "20201109003": "吳振凱", "09963": "黃謙凱",
    "20240313003": "阮曰忠", "20201109001": "梁冠如", "10003": "李茂銓",
    "20200922002": "葉欲弘", "20250923019": "阮氏紅深", "9963": "黃謙凱",
    "11399": "陳哲沅",
}

# 下午（或整體）扣休規則：依「首筆時間」「末筆時間」命中
BREAK_RULES = [
    (dt.time(20,45,0), dt.time(22,30,0),  0, "首≥20:45 且 末≤22:30 → 0 分鐘"),
    (dt.time(18,30,0), dt.time(20,30,0),  0, "首≥18:30 且 末≤20:30 → 0 分鐘"),
    (dt.time(15,30,0), dt.time(18, 0,0),  0, "首≥15:30 且 末≤18:00 → 0 分鐘"),
    (dt.time(13,30,0), dt.time(15,35,0),  0, "首≥13:30 且 末≤15:35 → 0 分鐘"),
    (dt.time(20,45,0), dt.time(23, 0,0),  0, "首≥20:45 且 末≤23:00 → 0 分鐘"),
    (dt.time(20, 0,0), dt.time(22, 0,0), 15, "首≥20:00 且 末≤22:00 → 15 分鐘"),
    (dt.time(18,30,0), dt.time(22, 0,0), 15, "首≥18:30 且 末≤22:00 → 15 分鐘"),
    (dt.time(19, 0,0), dt.time(22,30,0), 15, "首≥19:00 且 末≤22:30 → 15 分鐘"),
    (dt.time(13,30,0), dt.time(18, 0,0), 15, "首≥13:30 且 末≤18:00 → 15 分鐘"),
    (dt.time(16, 0,0), dt.time(20,40,0), 30, "首≥16:00 且 末≤20:40 → 30 分鐘"),
    (dt.time(15,30,0), dt.time(20,30,0), 30, "首≥15:30 且 末≤20:30 → 30 分鐘"),
    (dt.time(17, 0,0), dt.time(22,30,0), 45, "首≥17:00 且 末≤22:30 → 45 分鐘"),
    (dt.time(15,45,0), dt.time(22,30,0), 45, "首≥15:45 且 末≤22:30 → 45 分鐘"),
    (dt.time(13,30,0), dt.time(20,29,0), 45, "首≥13:30 且 末≤20:29 → 45 分鐘"),
    (dt.time(13,30,0), dt.time(23, 0,0), 60, "首≥13:30 且 末≤23:00 → 60 分鐘"),
    (dt.time(11, 0,0), dt.time(17, 0,0), 75, "首≥11:00 且 末≤17:00 → 75 分鐘"),
    (dt.time( 8, 0,0), dt.time(17, 0,0), 90, "首≥08:00 且 末≤17:00 → 90 分鐘"),
    (dt.time(10,50,0), dt.time(23, 0,0),120, "首≥10:50 且 末≤23:00 → 120 分鐘"),
    (dt.time( 8, 0,0), dt.time(23, 0,0),135, "首≥08:00 且 末≤23:00 → 135 分鐘"),
]

# 空窗固定帶：先扣掉（不計入空窗）
EXCLUDE_IDLE_RANGES = [
    (dt.time(10, 0, 0), dt.time(10, 15, 0)),
    (dt.time(12,30, 0), dt.time(13, 30, 0)),
    (dt.time(15,30, 0), dt.time(15, 45, 0)),
    (dt.time(18, 0, 0), dt.time(18, 30, 0)),
    (dt.time(20,30, 0), dt.time(20, 45, 0)),
]


# =========================
# 讀檔（支援 xlsx/xlsm/xls/csv）
# =========================
def read_excel_any_quiet_bytes(name: str, content: bytes) -> Dict[str, pd.DataFrame]:
    ext = (name.split(".")[-1] or "").lower()

    if ext in ("xlsx", "xlsm"):
        xl = pd.ExcelFile(io.BytesIO(content), engine="openpyxl")
        return {sn: pd.read_excel(xl, sheet_name=sn) for sn in xl.sheet_names}

    # ✅ 舊版 .xls：需要 requirements.txt 安裝 xlrd==2.0.1
    if ext == "xls":
        xl = pd.ExcelFile(io.BytesIO(content), engine="xlrd")
        return {sn: pd.read_excel(xl, sheet_name=sn) for sn in xl.sheet_names}

    if ext == "csv":
        for enc in ("utf-8-sig", "cp950", "big5"):
            try:
                return {"CSV": pd.read_csv(io.BytesIO(content), encoding=enc)}
            except Exception:
                continue
        raise Exception("CSV 讀取失敗（請確認編碼）。")

    raise Exception("目前僅支援 .xlsx/.xlsm/.xls/.csv")


# =========================
# 規則工具
# =========================
def _strip_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def find_first_column(df: pd.DataFrame, candidates: List[str]) -> Optional[str]:
    cols = [str(c).strip() for c in df.columns]
    s = set(cols)
    for name in candidates:
        if name in s:
            return name

    # 去掉括號/空白做容錯
    norm_map = {re.sub(r"[（）\(\)\s]", "", c): c for c in cols}
    for name in candidates:
        key = re.sub(r"[（）\(\)\s]", "", name)
        if key in norm_map:
            return norm_map[key]
    return None


def normalize_to_qc(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip().str.upper()
    return s.eq("QC")


def to_not_excluded_mask(series: pd.Series) -> pd.Series:
    s = series.astype(str).str.strip()
    return ~s.str.contains(TO_EXCLUDE_PATTERN, na=False)


def prepare_filtered_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame()
    df = _strip_cols(df)
    if "由" not in df.columns or "到" not in df.columns:
        return pd.DataFrame()
    return df[normalize_to_qc(df["由"]) & to_not_excluded_mask(df["到"])].copy()


def break_minutes_for_span(first_dt: pd.Timestamp, last_dt: pd.Timestamp) -> Tuple[int, str]:
    if pd.isna(first_dt) or pd.isna(last_dt):
        return 0, "無時間資料"
    stt, edt = first_dt.time(), last_dt.time()
    for st_ge, ed_le, mins, tag in BREAK_RULES:
        if (stt >= st_ge) and (edt <= ed_le):
            return int(mins), str(tag)
    return 0, "未命中規則"


def _subtract_exclusions(s_dt: pd.Timestamp, e_dt: pd.Timestamp, exclude_ranges):
    """把重疊部分完整切掉（避免重疊錯留）"""
    if s_dt >= e_dt or not exclude_ranges:
        return [(s_dt, e_dt)]
    segments = [(s_dt, e_dt)]
    for ex_s_t, ex_e_t in exclude_ranges:
        ex_s = pd.Timestamp.combine(s_dt.date(), ex_s_t)
        ex_e = pd.Timestamp.combine(s_dt.date(), ex_e_t)
        new_segments = []
        for a, b in segments:
            if b <= ex_s or a >= ex_e:
                new_segments.append((a, b))
            else:
                if a < ex_s:
                    new_segments.append((a, ex_s))
                if b > ex_e:
                    new_segments.append((ex_e, b))
        segments = [(x, y) for (x, y) in new_segments if x < y]
    return segments


def _compute_idle(series_dt: pd.Series, min_minutes=IDLE_MIN_THRESHOLD, exclude_ranges=EXCLUDE_IDLE_RANGES) -> Tuple[int, str]:
    if series_dt.size < 2:
        return 0, ""
    s = series_dt.sort_values()
    total_min, ranges_txt = 0, []
    prev = s.iloc[0]
    for cur in s.iloc[1:]:
        if cur <= prev:
            prev = cur
            continue
        for a, b in _subtract_exclusions(prev, cur, exclude_ranges or []):
            gap_min = int(round((b - a).total_seconds() / 60.0))
            if gap_min >= min_minutes:
                total_min += gap_min
                ranges_txt.append(f"{a.time()} ~ {b.time()}")
        prev = cur
    return int(total_min), "；".join(ranges_txt)


def _span_metrics(series_dt: pd.Series):
    if series_dt.empty:
        return pd.NaT, pd.NaT, 0
    return series_dt.min(), series_dt.max(), int(series_dt.size)


def _eff(n, m):
    return round((n / m * 60.0), 2) if m and m > 0 else 0.0


def compute_am_pm_for_group(g: pd.DataFrame) -> pd.Series:
    times = g["__dt__"]

    # 上午：07:00–12:30（不扣休）
    t_am = times[times.dt.time.between(AM_START, AM_END)]
    am_first, am_last, am_cnt = _span_metrics(t_am)
    am_mins = int(round(((am_last - am_first).total_seconds() / 60.0))) if am_cnt > 0 else 0
    am_eff = _eff(am_cnt, am_mins)
    am_idle_min, am_idle_ranges = _compute_idle(t_am)

    # 下午：13:30–23:59:59（扣休）
    t_pm = times[times.dt.time.between(PM_START, PM_END)]
    pm_first, pm_last, pm_cnt = _span_metrics(t_pm)
    if pm_cnt > 0:
        pm_break, pm_rule = break_minutes_for_span(pm_first, pm_last)
        raw_pm_mins = (pm_last - pm_first).total_seconds() / 60.0
        pm_mins = max(int(round(raw_pm_mins - pm_break)), 0)
    else:
        pm_break, pm_rule, pm_mins = 0, "無時間資料", 0
    pm_eff = _eff(pm_cnt, pm_mins)
    pm_idle_min, pm_idle_ranges = _compute_idle(t_pm)

    # 整體（扣休）
    whole_first, whole_last, day_cnt = _span_metrics(times)
    if day_cnt > 0:
        whole_break, br_tag_whole = break_minutes_for_span(whole_first, whole_last)
        raw_whole_mins = (whole_last - whole_first).total_seconds() / 60.0
        whole_mins = max(int(round(raw_whole_mins - whole_break)), 0)
    else:
        whole_break, br_tag_whole, whole_mins = 0, "無時間資料", 0
    whole_eff = _eff(day_cnt, whole_mins)

    return pd.Series({
        "第一筆時間": whole_first, "最後一筆時間": whole_last, "當日筆數": int(day_cnt),
        "休息分鐘_整體": int(whole_break), "命中規則": br_tag_whole,
        "當日工時_分鐘_扣休": int(whole_mins), "效率_件每小時": whole_eff,

        "上午_第一筆": am_first, "上午_最後一筆": am_last, "上午_筆數": int(am_cnt),
        "上午_工時_分鐘": int(am_mins), "上午_效率_件每小時": am_eff,
        "上午_空窗分鐘": int(am_idle_min), "上午_空窗時段": am_idle_ranges,

        "下午_第一筆": pm_first, "下午_最後一筆": pm_last, "下午_筆數": int(pm_cnt),
        "下午_休息分鐘": int(pm_break), "下午_命中規則": pm_rule,
        "下午_工時_分鐘_扣休": int(pm_mins), "下午_效率_件每小時": pm_eff,
        "下午_空窗分鐘_扣休": int(pm_idle_min), "下午_空窗時段": pm_idle_ranges,
    })


# =========================
# Excel 匯出（bytes）
# =========================
def autosize_columns(ws, df: pd.DataFrame):
    from openpyxl.utils import get_column_letter
    cols = list(df.columns) if df is not None else []
    for i, col in enumerate(cols, start=1):
        if df is not None and not df.empty:
            sample = [len(str(x)) for x in df[col].head(800).tolist()]
            max_len = max([len(str(col))] + sample)
        else:
            max_len = max(len(str(col)), 8)
        ws.column_dimensions[get_column_letter(i)].width = min(max_len + 2, 60)


def shade_rows_by_efficiency(ws, header_name="效率_件每小時", green="C6EFCE", red="FFC7CE"):
    from openpyxl.styles import PatternFill
    eff_col = None
    for c in range(1, ws.max_column + 1):
        if str(ws.cell(row=1, column=c).value).strip() == header_name:
            eff_col = c
            break
    if eff_col is None:
        return
    green_fill = PatternFill(start_color=green, end_color=green, fill_type="solid")
    red_fill = PatternFill(start_color=red, end_color=red, fill_type="solid")
    for r in range(2, ws.max_row + 1):
        v = ws.cell(row=r, column=eff_col).value
        try:
            val = float(v) if v is not None and str(v).strip() != "" else None
        except Exception:
            val = None
        if val is None:
            continue
        fill = green_fill if val >= TARGET_EFF else red_fill
        for c in range(1, ws.max_column + 1):
            ws.cell(row=r, column=c).fill = fill


def build_excel_bytes(user_col: str, summary_out: pd.DataFrame, daily: pd.DataFrame, detail_long: pd.DataFrame) -> bytes:
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="openpyxl", datetime_format="yyyy-mm-dd hh:mm:ss", date_format="yyyy-mm-dd") as writer:
        # 彙總
        sum_cols = [
            user_col, "對應姓名", "総日數",
            "總筆數", "總工時_分鐘_扣休", "效率_件每小時",
            "上午筆數", "上午工時_分鐘", "上午效率_件每小時",
            "下午筆數", "下午工時_分鐘_扣休", "下午效率_件每小時",
        ]
        summary_out[sum_cols].to_excel(writer, index=False, sheet_name="彙總")
        ws_sum = writer.sheets["彙總"]
        autosize_columns(ws_sum, summary_out[sum_cols])
        shade_rows_by_efficiency(ws_sum, "效率_件每小時")

        # 明細（每日）
        det_cols = [
            user_col, "對應姓名", "日期",
            "第一筆時間", "最後一筆時間", "當日筆數",
            "休息分鐘_整體", "當日工時_分鐘_扣休", "效率_件每小時",
            "上午_第一筆", "上午_最後一筆", "上午_筆數", "上午_工時_分鐘", "上午_效率_件每小時",
            "上午_空窗分鐘", "上午_空窗時段",
            "下午_第一筆", "下午_最後一筆", "下午_筆數", "下午_休息分鐘",
            "下午_工時_分鐘_扣休", "下午_效率_件每小時",
            "下午_空窗分鐘_扣休", "下午_空窗時段",
        ]
        daily.sort_values([user_col, "日期", "第一筆時間"])[det_cols].to_excel(writer, index=False, sheet_name="明細")
        ws_det = writer.sheets["明細"]
        autosize_columns(ws_det, daily[det_cols])
        shade_rows_by_efficiency(ws_det, "效率_件每小時")

        # 明細_時段（長表）
        if detail_long is not None and not detail_long.empty:
            long_cols = [
                user_col, "對應姓名", "日期", "時段",
                "第一筆時間", "最後一筆時間",
                "筆數", "工時_分鐘", "休息分鐘",
                "空窗分鐘", "空窗時段",
                "效率_件每小時", "命中規則",
            ]
            detail_long[long_cols].to_excel(writer, index=False, sheet_name="明細_時段")
            ws_long = writer.sheets["明細_時段"]
            autosize_columns(ws_long, detail_long[long_cols])
            shade_rows_by_efficiency(ws_long, "效率_件每小時")

        # 休息規則
        rules_rows = []
        for i, (st_ge, ed_le, mins, tag) in enumerate(BREAK_RULES, start=1):
            rules_rows.append({
                "優先序": i,
                "首時間條件(>=)": st_ge.strftime("%H:%M:%S"),
                "末時間條件(<=)": ed_le.strftime("%H:%M:%S"),
                "休息分鐘": int(mins),
                "規則說明": str(tag),
            })
        rules_df = pd.DataFrame(rules_rows, columns=["優先序", "首時間條件(>=)", "末時間條件(<=)", "休息分鐘", "規則說明"])
        rules_df.to_excel(writer, index=False, sheet_name="休息規則")
        autosize_columns(writer.sheets["休息規則"], rules_df)

    return out.getvalue()


# =========================
# Streamlit Page
# =========================
def main():
    inject_logistics_theme()
    set_page("上架產能分析（Putaway KPI）", icon="📦")
    st.caption("總上組（上架）｜依原始 QC 紀錄計算｜AM/PM 班別（上午/下午）｜支援 .xls")

    with st.sidebar:
        st.header("⚙️ 計算條件設定")
        operator = st.text_input("分析執行人（Operator）")
        top_n = st.number_input("效率排行顯示人數（Top N）", 10, 100, 30, step=5)
        st.info("提醒：上傳 .xls 需 requirements.txt 安裝 xlrd==2.0.1")

    card_open("📤 上傳作業原始資料（上架）")
    uploaded = st.file_uploader(
        "上傳 Excel / CSV（包含『由/到』『修訂日期/時間』『記錄輸入人』）",
        type=["xlsx", "xlsm", "xls", "csv"],
        label_visibility="collapsed",
    )
    run = st.button("🚀 產出 KPI", type="primary", disabled=uploaded is None)
    card_close()

    if not run:
        st.info("請先上傳上架作業原始資料")
        return

    with st.spinner("計算中，請稍候..."):
        content = uploaded.getvalue()
        sheets = read_excel_any_quiet_bytes(uploaded.name, content)

        kept_all = []
        for sn, df in sheets.items():
            k = prepare_filtered_df(df)
            if not k.empty:
                k["__sheet__"] = sn
                kept_all.append(k)
        if not kept_all:
            st.error("無符合資料（可能缺『由/到』欄或過濾後為空）。")
            return

        data = pd.concat(kept_all, ignore_index=True)

        user_col = find_first_column(data, INPUT_USER_CANDIDATES)
        revdt_col = find_first_column(data, REV_DT_CANDIDATES)
        if user_col is None:
            st.error("找不到『記錄輸入人』欄位（候選：記錄輸入人/記錄輸入者/建立人/輸入人）。")
            return
        if revdt_col is None:
            st.error("找不到『修訂日期/時間』欄位（候選：修訂日期/修訂時間/異動時間/修改時間…）。")
            return

        data["__dt__"] = pd.to_datetime(data[revdt_col], errors="coerce")
        data["__code__"] = data[user_col].astype(str).str.strip()
        data["對應姓名"] = data["__code__"].map(NAME_MAP).fillna("")

        dt_data = data.dropna(subset=["__dt__"]).copy()
        if dt_data.empty:
            st.error("資料沒有可用的修訂日期時間，無法計算。")
            return

        dt_data["日期"] = dt_data["__dt__"].dt.date

        daily = (
            dt_data.groupby([user_col, "對應姓名", "日期"], dropna=False)
                   .apply(compute_am_pm_for_group)
                   .reset_index()
        )

        summary = (
            daily.groupby([user_col, "對應姓名"], dropna=False, as_index=False)
                 .agg(
                     総日數=("日期", "nunique"),
                     總筆數=("當日筆數", "sum"),
                     總工時_分鐘_扣休=("當日工時_分鐘_扣休", "sum"),
                     上午筆數=("上午_筆數", "sum"),
                     上午工時_分鐘=("上午_工時_分鐘", "sum"),
                     下午筆數=("下午_筆數", "sum"),
                     下午工時_分鐘_扣休=("下午_工時_分鐘_扣休", "sum"),
                 )
        )
        summary["上午效率_件每小時"] = summary.apply(lambda r: _eff(r["上午筆數"], r["上午工時_分鐘"]), axis=1)
        summary["下午效率_件每小時"] = summary.apply(lambda r: _eff(r["下午筆數"], r["下午工時_分鐘_扣休"]), axis=1)
        summary["總工時_分鐘_扣休"] = summary["上午工時_分鐘"].fillna(0).astype(int) + summary["下午工時_分鐘_扣休"].fillna(0).astype(int)
        summary["效率_件每小時"] = summary.apply(lambda r: _eff(r["總筆數"], r["總工時_分鐘_扣休"]), axis=1)

        for c in ["總筆數", "總工時_分鐘_扣休", "上午筆數", "上午工時_分鐘", "下午筆數", "下午工時_分鐘_扣休"]:
            summary[c] = summary[c].fillna(0).astype(int)

        # 合計列（彙總）
        total_people = int(summary[user_col].nunique())
        total_met = int((summary["效率_件每小時"] >= TARGET_EFF).sum())
        total_rate = (total_met / total_people) if total_people > 0 else 0.0

        total_row = {
            user_col: "整體合計", "對應姓名": "",
            "総日數": int(summary["総日數"].sum()),
            "總筆數": int(summary["總筆數"].sum()),
            "總工時_分鐘_扣休": int(summary["總工時_分鐘_扣休"].sum()),
            "上午筆數": int(summary["上午筆數"].sum()),
            "上午工時_分鐘": int(summary["上午工時_分鐘"].sum()),
            "下午筆數": int(summary["下午筆數"].sum()),
            "下午工時_分鐘_扣休": int(summary["下午工時_分鐘_扣休"].sum()),
            "效率_件每小時": _eff(int(summary["總筆數"].sum()), int(summary["總工時_分鐘_扣休"].sum())),
            "上午效率_件每小時": _eff(int(summary["上午筆數"].sum()), int(summary["上午工時_分鐘"].sum())),
            "下午效率_件每小時": _eff(int(summary["下午筆數"].sum()), int(summary["下午工時_分鐘_扣休"].sum())),
        }
        summary_out = pd.concat([summary, pd.DataFrame([total_row])], ignore_index=True)

        # 長表：明細_時段
        long_rows = []
        for _, r in daily.iterrows():
            if r["上午_筆數"] > 0:
                long_rows.append({
                    user_col: r[user_col], "對應姓名": r["對應姓名"], "日期": r["日期"], "時段": "上午",
                    "第一筆時間": r["上午_第一筆"], "最後一筆時間": r["上午_最後一筆"],
                    "筆數": int(r["上午_筆數"]),
                    "工時_分鐘": int(r["上午_工時_分鐘"]),
                    "休息分鐘": 0,
                    "空窗分鐘": int(r["上午_空窗分鐘"]),
                    "空窗時段": r["上午_空窗時段"],
                    "效率_件每小時": float(r["上午_效率_件每小時"]),
                    "命中規則": "上午不扣休",
                })
            if r["下午_筆數"] > 0:
                long_rows.append({
                    user_col: r[user_col], "對應姓名": r["對應姓名"], "日期": r["日期"], "時段": "下午",
                    "第一筆時間": r["下午_第一筆"], "最後一筆時間": r["下午_最後一筆"],
                    "筆數": int(r["下午_筆數"]),
                    "工時_分鐘": int(r["下午_工時_分鐘_扣休"]),
                    "休息分鐘": int(r["下午_休息分鐘"]),
                    "空窗分鐘": int(r["下午_空窗分鐘_扣休"]),
                    "空窗時段": r["下午_空窗時段"],
                    "效率_件每小時": float(r["下午_效率_件每小時"]),
                    "命中規則": str(r["下午_命中規則"]),
                })
        detail_long = pd.DataFrame(long_rows)
        if not detail_long.empty:
            detail_long = detail_long.sort_values([user_col, "日期", "時段", "第一筆時間"])

        # 匯出 bytes
        xlsx_bytes = build_excel_bytes(user_col, summary_out, daily, detail_long)

    # ======================
    # 介面：左右 AM/PM
    # ======================
    plot_df = summary.copy()

    col_l, col_r = st.columns(2)

    with col_l:
        card_open("🌓 AM 班（上午）KPI")
        render_kpis([
            KPI("人數", f"{int(plot_df[user_col].nunique()):,}"),
            KPI("達標門檻", f"效率 ≥ {TARGET_EFF}"),
        ])
        card_close()

        card_open(f"AM 班（上午）效率排行（Top {int(top_n)}）")
        am_rank = plot_df[[user_col, "對應姓名", "上午筆數", "上午工時_分鐘", "上午效率_件每小時"]].copy()
        am_rank = am_rank.rename(columns={"上午效率_件每小時": "效率", "上午筆數": "筆數", "上午工時_分鐘": "工時"})
        am_rank["姓名"] = am_rank["對應姓名"].where(am_rank["對應姓名"].astype(str).str.len() > 0, am_rank[user_col].astype(str))
        bar_topN(
            am_rank[["姓名", "效率", "筆數", "工時"]],
            x_col="姓名", y_col="效率",
            hover_cols=["筆數", "工時"],
            top_n=int(top_n),
            target=float(TARGET_EFF),
        )
        card_close()

    with col_r:
        card_open("🌙 PM 班（下午）KPI")
        pm_met = int((plot_df["下午效率_件每小時"] >= TARGET_EFF).sum())
        pm_total = int(plot_df[user_col].nunique())
        pm_rate = (pm_met / pm_total) if pm_total > 0 else 0.0
        render_kpis([
            KPI("人數", f"{pm_total:,}"),
            KPI("達標人數", f"{pm_met:,}"),
            KPI("達標率", f"{pm_rate:.1%}"),
            KPI("達標門檻", f"效率 ≥ {TARGET_EFF}"),
        ])
        card_close()

        card_open(f"PM 班（下午）效率排行（Top {int(top_n)}）")
        pm_rank = plot_df[[user_col, "對應姓名", "下午筆數", "下午工時_分鐘_扣休", "下午效率_件每小時"]].copy()
        pm_rank = pm_rank.rename(columns={"下午效率_件每小時": "效率", "下午筆數": "筆數", "下午工時_分鐘_扣休": "工時"})
        pm_rank["姓名"] = pm_rank["對應姓名"].where(pm_rank["對應姓名"].astype(str).str.len() > 0, pm_rank[user_col].astype(str))
        bar_topN(
            pm_rank[["姓名", "效率", "筆數", "工時"]],
            x_col="姓名", y_col="效率",
            hover_cols=["筆數", "工時"],
            top_n=int(top_n),
            target=float(TARGET_EFF),
        )
        card_close()

    # 匯出
    card_open("⬇️ 匯出 KPI 報表（Excel）")
    default_name = f"{uploaded.name.rsplit('.', 1)[0]}_上架績效.xlsx"
    download_excel(xlsx_bytes, default_name)
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
            "source_sha256": sha256_bytes(content),
            "params": {
                "top_n": int(top_n),
                "target_eff": TARGET_EFF,
                "filter": "由=QC 且 到不含關鍵字",
                "am_range": "07:00-12:30",
                "pm_range": "13:30-23:59:59",
                "idle_min_threshold": IDLE_MIN_THRESHOLD,
                "idle_exclude_ranges": [(a.strftime("%H:%M"), b.strftime("%H:%M")) for a, b in EXCLUDE_IDLE_RANGES],
            },
            "kpi_am": {"people": total_people, "pass_rate": total_rate},
            "kpi_pm": {"people": int(plot_df[user_col].nunique()), "pass_rate": float(pm_rate)},
            "export_object_path": export_path,
        }
        row = insert_audit_run(payload)
        st.success(f"✅ 已成功留存本次分析（ID：{row.get('id','')}）")
    except Exception as e:
        st.error("❌ 稽核留存發生錯誤")
        st.code(repr(e))


if __name__ == "__main__":
    main()
