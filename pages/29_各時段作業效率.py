# pages/29_各時段作業效率.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import os
from io import StringIO
from datetime import datetime, date, timedelta, time
from zoneinfo import ZoneInfo

import numpy as np
import pandas as pd
import streamlit as st
import altair as alt

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.formatting.rule import FormulaRule

# ---- 套用平台風格（有就用，沒有就退回原生）----
try:
    from common_ui import inject_logistics_theme, set_page, card_open, card_close
    HAS_COMMON_UI = True
except Exception:
    HAS_COMMON_UI = False

TPE = ZoneInfo("Asia/Taipei")

STATUS_PASS = "達標"
STATUS_FAIL = "未達標"
STATUS_NA = "未判斷"


# =========================================================
# 讀檔：CSV/Excel 強韌讀取（bytes）
# ✅ 支援「假 .xls」= TSV/CSV 文字檔
# =========================================================
def read_table_robust(file_name: str, raw: bytes, label: str = "檔案") -> pd.DataFrame:
    ext = os.path.splitext(file_name)[1].lower()

    # 1) 先試 Excel（WMS 的 .xls 常常其實是文字檔，失敗就往下走）
    if ext in (".xlsx", ".xlsm", ".xltx", ".xltm", ".xls"):
        try:
            return pd.read_excel(io.BytesIO(raw))
        except Exception:
            pass

    # 2) 文字檔解析：優先試 tab（TSV），再試常見分隔符
    encodings = ["utf-8-sig", "utf-8", "cp950", "big5", "ms950", "gb18030", "latin1"]
    seps = ["\t", ",", ";", "|"]

    last_err = None
    for enc in encodings:
        for sep in seps:
            try:
                df = pd.read_csv(io.BytesIO(raw), encoding=enc, sep=sep, engine="python")
                if df.shape[1] <= 1:
                    continue
                return df
            except Exception as e:
                last_err = e

    # 3) 最後容錯：讓 pandas 自己猜分隔符
    try:
        text = raw.decode("utf-8", errors="replace")
        df = pd.read_csv(StringIO(text), sep=None, engine="python")
        if df.shape[1] <= 1:
            raise ValueError("偵測不到有效分隔符，請確認檔案是否為真正表格檔。")
        return df
    except Exception as e:
        raise ValueError(f"{label} 讀取失敗（已嘗試多種編碼/分隔符）：{last_err} / 最終：{e}")


def _norm_cols(df: pd.DataFrame) -> pd.DataFrame:
    df = df.copy()
    df.columns = [str(c).strip() for c in df.columns]
    return df


def require_columns(df: pd.DataFrame, required: list, label: str):
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError(f"{label} 缺少欄位：{missing}\n目前欄位：{list(df.columns)}")


def clean_line(series: pd.Series) -> pd.Series:
    return series.astype(str).str.strip()


def clean_zone_1to4(series: pd.Series) -> pd.Series:
    z = pd.to_numeric(series, errors="coerce").astype("Int64")
    return z.where(z.between(1, 4, inclusive="both"))


def _safe_time(s: str) -> str:
    s = str(s).strip()
    if not s:
        return "08:00"
    try:
        datetime.strptime(s, "%H:%M")
        return s
    except Exception:
        return "08:00"


def _time_str_to_time(t: str) -> time:
    t = _safe_time(t)
    hh, mm = t.split(":")
    return time(int(hh), int(mm), 0)


def _ensure_tpe(series: pd.Series) -> pd.Series:
    """
    ✅ 統一時間欄位為 Asia/Taipei
    - 若是 tz-naive：視為台北時間並 localize
    - 若是 tz-aware：轉換為台北時間
    """
    s = pd.to_datetime(series, errors="coerce")
    if s.dt.tz is None:
        return s.dt.tz_localize(TPE)
    return s.dt.tz_convert(TPE)


# =========================================================
# ✅ 午休規則：12:30–13:30 不算工時
# =========================================================
def _is_work_time(ts: pd.Timestamp) -> bool:
    h = int(ts.hour)
    m = int(ts.minute)
    if h == 12 and m >= 30:
        return False
    if h == 13 and m < 30:
        return False
    return True


def _hour_work_segment(day: date, hour: int) -> tuple[datetime, datetime] | None:
    """
    回傳某小時內的「有效工作區間」(start_dt, end_dt)：
    - 12點：12:00–12:30
    - 13點：13:30–14:00
    - 其他：hour:00–hour+1:00
    """
    d = day
    if hour == 12:
        return (
            datetime(d.year, d.month, d.day, 12, 0, 0, tzinfo=TPE),
            datetime(d.year, d.month, d.day, 12, 30, 0, tzinfo=TPE),
        )
    if hour == 13:
        return (
            datetime(d.year, d.month, d.day, 13, 30, 0, tzinfo=TPE),
            datetime(d.year, d.month, d.day, 14, 0, 0, tzinfo=TPE),
        )
    if int(hour) < 23:
        return (
            datetime(d.year, d.month, d.day, int(hour), 0, 0, tzinfo=TPE),
            datetime(d.year, d.month, d.day, int(hour) + 1, 0, 0, tzinfo=TPE),
        )
    return (
        datetime(d.year, d.month, d.day, 23, 0, 0, tzinfo=TPE),
        datetime(d.year, d.month, d.day, 23, 59, 59, tzinfo=TPE),
    )


def _overlap_seconds(a0: datetime, a1: datetime, b0: datetime, b1: datetime) -> float:
    s = max(a0, b0)
    e = min(a1, b1)
    if e <= s:
        return 0.0
    return float((e - s).total_seconds())


def _effective_work_seconds_between(start_dt: datetime, end_dt: datetime) -> float:
    """
    計算 start_dt ~ end_dt 之間，扣除午休後的有效工作秒數（每秒精準）。
    """
    if end_dt <= start_dt:
        return 0.0
    day = start_dt.date()
    if end_dt.date() != day:
        end_dt = datetime(day.year, day.month, day.day, 23, 59, 59, tzinfo=TPE)

    total = 0.0
    for h in range(int(start_dt.hour), int(end_dt.hour) + 1):
        seg = _hour_work_segment(day, h)
        if seg is None:
            continue
        seg_s, seg_e = seg
        total += _overlap_seconds(start_dt, end_dt, seg_s, seg_e)
    return total


# =============================
# KPI 計數
# =============================
def _kpi_counts(dist_df: pd.DataFrame):
    if dist_df is None or dist_df.empty:
        return 0, 0, None
    p = int(dist_df.loc[dist_df["狀態"] == STATUS_PASS, "count"].sum())
    f = int(dist_df.loc[dist_df["狀態"] == STATUS_FAIL, "count"].sum())
    rate = (p / (p + f) * 100.0) if (p + f) > 0 else None
    return p, f, rate


# =============================
# Heatmap（Streamlit 顯示用）
# =============================
def render_hourly_heatmap(df_line_hourly: pd.DataFrame, hour_cols, title: str):
    if df_line_hourly is None or df_line_hourly.empty:
        st.info("沒有可呈現的圖。")
        return

    hour_cols = [int(h) for h in list(hour_cols)]
    plot = df_line_hourly.copy()

    plot["段數"] = pd.to_numeric(plot["段數"], errors="coerce").fillna(0).astype(int)
    plot["小時"] = pd.to_numeric(plot["小時"], errors="coerce").fillna(0).astype(int)
    plot["當小時加權PCS"] = pd.to_numeric(plot["當小時加權PCS"], errors="coerce").fillna(0.0)
    plot["本小時目標"] = pd.to_numeric(plot["本小時目標"], errors="coerce").fillna(0.0)

    plot["label"] = plot["段數"].astype(str) + "段｜" + plot["姓名"].astype(str)
    plot["狀態_色"] = plot["狀態"].fillna(STATUS_NA)
    plot["顯示量"] = plot["當小時加權PCS"].apply(lambda x: "" if abs(float(x)) < 1e-12 else f"{float(x):.0f}")

    order = (
        plot[["label", "段數", "姓名"]]
        .drop_duplicates()
        .sort_values(["段數", "姓名"])["label"]
        .tolist()
    )

    color_enc = alt.Color(
        "狀態_色:N",
        scale=alt.Scale(domain=[STATUS_PASS, STATUS_FAIL, STATUS_NA], range=["#16a34a", "#dc2626", "#d0d5dd"]),
        legend=alt.Legend(title="狀態"),
    )

    base = alt.Chart(plot).encode(
        x=alt.X("小時:O", sort=[str(h) for h in hour_cols], title="每小時"),
        y=alt.Y("label:N", sort=order, title="段數｜姓名"),
        tooltip=[
            alt.Tooltip("線別:N", title="線別"),
            alt.Tooltip("label:N", title="段數｜姓名"),
            alt.Tooltip("小時:O", title="小時"),
            alt.Tooltip("當小時加權PCS:Q", title="當小時加權PCS", format=",.4f"),
            alt.Tooltip("本小時目標:Q", title="本小時目標", format=",.2f"),
            alt.Tooltip("狀態:N", title="狀態"),
        ],
    )

    rect = base.mark_rect(cornerRadius=4).encode(color=color_enc)
    text = base.mark_text(fontSize=12, fontWeight=900).encode(text="顯示量:N")

    n_rows = max(1, plot["label"].nunique())
    height = min(42 * n_rows + 80, 900)
    st.altair_chart((rect + text).properties(title=title, height=height), use_container_width=True)


# =============================
# ✅ 即時看板 UI（每線每段每人）
# ✅ 右側「總量(該線該段累計)」= 當天該線該段累積（不是最近N分鐘）
# =============================
def _board_css():
    st.markdown(
        """
<style>
.board-row { display: grid; grid-template-columns: 160px 170px 1fr 260px; gap: 14px;
             align-items: center; padding: 10px 10px; border-bottom: 1px solid rgba(0,0,0,0.08); }
.board-line { font-size: 40px; font-weight: 900; color:#1e40af; line-height: 1; }
.board-start { font-size: 18px; font-weight: 700; color:#111827; margin-top:6px; }
.board-hourly { font-size: 56px; font-weight: 900; color:#111827; text-align: right; line-height: 1; }
.board-total { font-size: 56px; font-weight: 900; color:#111827; text-align: right; line-height: 1; }
.board-sub { font-size: 18px; font-weight: 900; text-align:right; margin-top:6px; }
.board-bar { width: 100%; height: 34px; background: #e5e7eb; border-radius: 4px; overflow: hidden; position: relative;}
.board-bar > .fill { height: 100%; width: 0%; }
.board-bar > .txt { position:absolute; inset:0; display:flex; align-items:center; justify-content:center;
                    font-size:20px; font-weight:900; color:#111827; }
.board-head { display:grid; grid-template-columns: 160px 170px 1fr 260px; gap: 14px;
              align-items:end; padding: 6px 10px 12px 10px; }
.board-head .h { font-size: 22px; font-weight: 900; color:#1f2937; }
.board-head .hr { font-size: 22px; font-weight: 900; color:#1f2937; text-align:right; }
.badge { display:inline-flex; align-items:center; gap:8px; font-weight:900; }
.dot { width:10px; height:10px; border-radius:99px; display:inline-block; }
</style>
""",
        unsafe_allow_html=True,
    )


def render_person_realtime_board(board_df: pd.DataFrame):
    if board_df is None or board_df.empty:
        st.info("目前沒有可呈現的即時看板資料。")
        return

    _board_css()

    st.markdown(
        """
<div class="board-head">
  <div class="h">線別｜段｜人｜開線</div>
  <div class="h">本小時(實績)</div>
  <div></div>
  <div class="hr">總量(該線該段累計｜當天)</div>
</div>
""",
        unsafe_allow_html=True,
    )

    show = board_df.copy()
    show = show.sort_values(["段數", "排序_未達標優先", "姓名"], ascending=[True, True, True])

    for _, r in show.iterrows():
        line = str(r["線別"])
        zid = int(r["段數"]) if pd.notna(r["段數"]) else 0
        name = str(r["姓名"])
        stime = str(r.get("開始時間", "08:00"))

        hourly = float(r.get("每小時分揀量", 0.0))

        # ✅ 右側總量(累積) → 當天該線該段累計
        total_seg = float(r.get("已分揀總量", 0.0))
        tgt_seg = float(r.get("累積目標", 0.0))
        status_seg = r.get("累積狀態", None)

        tgt_h = float(r.get("本小時目標", 0.0))
        status_h = r.get("狀態", None)

        pct = 0.0
        if tgt_h > 0:
            pct = max(0.0, min(hourly / tgt_h, 1.0)) * 100.0

        if status_h == STATUS_FAIL:
            bar_color = "#dc2626"; dot = "#dc2626"
        elif status_h == STATUS_PASS:
            bar_color = "#16a34a"; dot = "#16a34a"
        else:
            bar_color = "#9ca3af"; dot = "#9ca3af"

        if status_seg == STATUS_FAIL:
            total_color = "#dc2626"
        elif status_seg == STATUS_PASS:
            total_color = "#16a34a"
        else:
            total_color = "#6b7280"

        st.markdown(
            f"""
<div class="board-row">
  <div>
    <div class="board-line">{line}</div>
    <div class="board-start">
      <span class="badge"><span class="dot" style="background:{dot};"></span>{zid}段｜{name}｜{stime}</span>
    </div>
  </div>

  <div class="board-hourly">{hourly:,.0f}</div>

  <div>
    <div class="board-bar">
      <div class="fill" style="width:{pct:.1f}%; background:{bar_color};"></div>
      <div class="txt">本小時目標 {tgt_h:,.0f}</div>
    </div>
  </div>

  <div>
    <div class="board-total">{total_seg:,.0f}</div>
    <div class="board-sub" style="color:{total_color};">累積目標 {tgt_seg:,.0f}</div>
  </div>
</div>
""",
            unsafe_allow_html=True,
        )


def main():
    st.set_page_config(page_title="大豐物流 - 出貨課｜各時段作業效率", page_icon="⏱️", layout="wide")
    if HAS_COMMON_UI:
        inject_logistics_theme()
        set_page("📦 出貨課", "⏱️ 29｜各時段作業效率")

    st.markdown("### ⏱️ 各時段作業效率（每線×每段×每人 即時產能｜段數固定 1→4｜午休 12:30–13:30｜總量=當天）")

    fixed_time_map = {
        "范明俊": "08:00", "阮玉名": "08:00", "李茂銓": "08:00", "河文強": "08:00",
        "蔡麗珠": "08:00", "潘文一": "08:00", "阮伊黃": "08:00", "葉欲弘": "09:00",
        "阮武玉玄": "08:00", "吳黃金珠": "08:30", "潘氏青江": "08:00", "陳國慶": "08:30",
        "楊心如": "08:00", "阮瑞美黃緣": "08:00", "周芸蓁": "08:00", "黎氏瓊": "08:00",
        "王文楷": "08:30", "潘氏慶平": "08:00", "阮氏美麗": "08:00", "岳子恆": "08:30",
        "郭雙燕": "08:30", "阮孟勇": "08:00", "廖永成": "08:30", "楊浩傑": "08:30",
        "黃日康": "08:30", "蔣金妮": "08:30", "柴家欣": "08:30", "邱思捷": "09:00",
        "王建成": "09:00",
    }

    with st.sidebar:
        st.markdown("### 設定")
        target_hr = st.number_input("每人每小時目標（加權PCS/小時）", min_value=1.0, value=790.0, step=10.0)
        hour_min = st.number_input("起始小時（Heatmap 用）", min_value=0, max_value=23, value=8, step=1)
        # ✅ lookback 不再影響「總量」，只用於提示/顯示
        lookback_min = st.number_input("提示：資料最新時間顯示區間（分鐘）", min_value=10, max_value=24 * 60, value=180, step=10)
        st.caption("✅ 看板「總量/本小時」皆以【當天】資料計算；午休 12:30–13:30 不算。")

    prod_files = st.file_uploader(
        "① 上傳『WMS 生產資料』（可多檔：CSV/TXT/XLS/XLSX；可全選上傳）",
        type=["csv", "txt", "xls", "xlsx", "xlsm"],
        accept_multiple_files=True,
    )
    mem_file = st.file_uploader("② 上傳『人員名單』(CSV/Excel)", type=["csv", "xlsx", "xlsm", "xls"])

    if not prod_files or mem_file is None:
        st.info("請上傳：① WMS 生產資料（可多檔） + ② 人員名單")
        return

    try:
        # ✅ 現在時間（tz-aware）
        now = datetime.now(TPE)
        cur_h, cur_m, cur_s = int(now.hour), int(now.minute), int(now.second)
        day = now.date()

        # -------------------------
        # 人員名單解析（每線×段×人）
        # -------------------------
        df_mem_raw = _norm_cols(read_table_robust(mem_file.name, mem_file.getvalue(), label="人員名單檔案"))

        line_col_candidates = ["LINEID", "線別", "LineID", "LINE Id", "Line Id"]
        line_col = next((c for c in line_col_candidates if c in df_mem_raw.columns), None)
        if line_col is None:
            raise ValueError("人員名單找不到線別欄位（需要 LINEID 或 線別）。")

        seg_cols = {1: "第一段", 2: "第二段", 3: "第三段", 4: "第四段"}
        for _, colname in seg_cols.items():
            if colname not in df_mem_raw.columns:
                raise ValueError(f"人員名單缺少欄位：{colname}（需要 第一段～第四段）")

        member_list = []
        for _, row in df_mem_raw.iterrows():
            line_id = str(row.get(line_col, "")).strip()
            if not line_id or line_id.lower() == "nan":
                continue
            for zid, colname in seg_cols.items():
                name = row.get(colname, None)
                if pd.notna(name) and str(name).strip() != "":
                    n_str = str(name).strip()
                    st_time = _safe_time(fixed_time_map.get(n_str, "08:00"))
                    member_list.append({"線別": line_id, "段數": zid, "姓名": n_str, "開始時間": st_time})

        roster_df = pd.DataFrame(member_list)
        if roster_df.empty:
            raise ValueError("人員名單解析後為空：請確認 第一段～第四段 內有姓名。")

        roster_df["線別"] = clean_line(roster_df["線別"])
        roster_df["段數"] = clean_zone_1to4(roster_df["段數"])
        roster_df = roster_df[roster_df["段數"].notna()].copy()
        roster_df["段數"] = roster_df["段數"].astype(int)

        # ✅ 每線每段只留一人（你若要同段多名 → 把這行刪掉）
        roster_df = roster_df.drop_duplicates(["線別", "段數"], keep="first").copy()
        roster_df = roster_df[["線別", "段數", "姓名", "開始時間"]].copy()

        # -------------------------
        # 生產資料：多檔合併
        # -------------------------
        frames = []
        for f in prod_files:
            raw = f.getvalue()
            dfp = read_table_robust(f.name, raw, label=f"生產資料（{f.name}）")
            dfp["__source__"] = f.name
            frames.append(dfp)

        df_raw = _norm_cols(pd.concat(frames, ignore_index=True))
        require_columns(df_raw, ["PICKDATE", "LINEID", "ZONEID", "PACKQTY", "Cweight"], "生產資料（合併）")

        # ✅ 時間統一台北時區（避免 dtype datetime64[ns] vs datetime）
        df_raw["PICKDATE"] = _ensure_tpe(df_raw["PICKDATE"])
        df_raw = df_raw[df_raw["PICKDATE"].notna()].copy()

        # ✅ 只保留「當天」資料（看板總量 = 當天）
        df_raw = df_raw[df_raw["PICKDATE"].dt.tz_convert(TPE).dt.date == day].copy()
        if df_raw.empty:
            st.warning("生產資料中找不到『今天』的資料（台北時區）。")
            return

        max_ts = df_raw["PICKDATE"].max()
        hint_cutoff = now - timedelta(minutes=int(lookback_min))
        hint_min_ts = df_raw.loc[df_raw["PICKDATE"] >= hint_cutoff, "PICKDATE"].min()

        df_raw = df_raw.rename(columns={"LINEID": "線別", "ZONEID": "段數"})
        df_raw["線別"] = clean_line(df_raw["線別"])
        df_raw["段數"] = clean_zone_1to4(df_raw["段數"])
        df_raw = df_raw[df_raw["段數"].notna()].copy()
        df_raw["段數"] = df_raw["段數"].astype(int)

        df_raw["PACKQTY"] = pd.to_numeric(df_raw["PACKQTY"], errors="coerce").fillna(0)
        df_raw["Cweight"] = pd.to_numeric(df_raw["Cweight"], errors="coerce").fillna(0)

        # 去重（避免多檔重疊）— 以「當天資料」為範圍去重
        rid_cols = [c for c in df_raw.columns if c not in ("__rid",)]
        df_raw["__rid"] = pd.util.hash_pandas_object(df_raw[rid_cols], index=False)
        df_raw = df_raw.drop_duplicates("__rid", keep="first").copy()

        # 對 roster 綁定姓名 / 開始時間
        df = pd.merge(df_raw, roster_df, on=["線別", "段數"], how="left", validate="m:1")
        df["姓名"] = df["姓名"].fillna("未設定")
        df["開始時間"] = df["開始時間"].fillna("08:00").map(_safe_time)

        # 拆時間（秒級）
        df["小時"] = df["PICKDATE"].dt.hour
        df["PICK_SEC"] = df["PICKDATE"].dt.hour * 3600 + df["PICKDATE"].dt.minute * 60 + df["PICKDATE"].dt.second

        st_t = df["開始時間"].map(_time_str_to_time)
        df["開始秒"] = st_t.map(lambda x: int(x.hour) * 3600 + int(x.minute) * 60).astype(int)

        # ✅ 納入計算：
        # 1) 不早於開始時間
        # 2) 午休 12:30–13:30 不算
        is_work = df["PICKDATE"].apply(_is_work_time)
        df["納入計算"] = (df["PICK_SEC"] >= df["開始秒"]) & is_work
        df["排除原因"] = np.where(df["納入計算"], "", np.where(~is_work, "午休不算工時", "早於開始時間"))

        df["加權PCS"] = df["PACKQTY"] * df["Cweight"]
        df_today_in = df[df["納入計算"]].copy()

        st.caption(
            f"看板時間：{now.strftime('%Y-%m-%d %H:%M:%S')}｜"
            f"今日資料最新：{max_ts.strftime('%H:%M:%S')}｜"
            f"提示(近{int(lookback_min)}分)最早資料：{(hint_min_ts.strftime('%H:%M:%S') if pd.notna(hint_min_ts) else '—')}｜"
            f"午休：12:30–13:30 不算工時｜"
            f"✅『總量』一律用【今天】資料計算"
        )

        # -------------------------
        # ✅ 本小時 / ✅ 當天累積（該線該段）
        # -------------------------
        hourly_person = (
            df_today_in[df_today_in["小時"] == cur_h]
            .groupby(["線別", "段數"], as_index=False)["加權PCS"].sum()
            .rename(columns={"加權PCS": "每小時分揀量"})
        )
        total_person = (
            df_today_in.groupby(["線別", "段數"], as_index=False)["加權PCS"].sum()
            .rename(columns={"加權PCS": "已分揀總量"})  # ✅ 當天該線該段累計
        )

        person_board = roster_df.merge(total_person, on=["線別", "段數"], how="left").merge(
            hourly_person, on=["線別", "段數"], how="left"
        )
        person_board["已分揀總量"] = pd.to_numeric(person_board["已分揀總量"], errors="coerce").fillna(0.0)
        person_board["每小時分揀量"] = pd.to_numeric(person_board["每小時分揀量"], errors="coerce").fillna(0.0)

        # -------------------------
        # ✅ 目標（每秒）— 累積目標也是「今天起班到現在」(扣午休)
        # -------------------------
        def _calc_targets(row) -> tuple[float, float]:
            stime = _safe_time(row["開始時間"])
            st_dt = datetime.combine(day, _time_str_to_time(stime)).replace(tzinfo=TPE)

            # 本小時：當小時有效工作區間 與 [起班, now] 的交集秒數
            seg = _hour_work_segment(day, cur_h)
            if seg is None:
                sec_h = 0.0
            else:
                seg_s, seg_e = seg
                sec_h = _overlap_seconds(st_dt, now, seg_s, seg_e)

            # 累積：起班到 now 的有效秒數（扣午休）
            sec_total = _effective_work_seconds_between(st_dt, now)

            tgt_h = float(target_hr) * (sec_h / 3600.0)
            tgt_total = float(target_hr) * (sec_total / 3600.0)
            return tgt_h, tgt_total

        tgts = person_board.apply(_calc_targets, axis=1, result_type="expand")
        person_board["本小時目標"] = pd.to_numeric(tgts[0], errors="coerce").fillna(0.0)
        person_board["累積目標"] = pd.to_numeric(tgts[1], errors="coerce").fillna(0.0)

        # ✅ 本小時達標
        person_board["狀態"] = np.where(
            person_board["本小時目標"] <= 1e-12,
            None,
            np.where(person_board["每小時分揀量"] >= person_board["本小時目標"], STATUS_PASS, STATUS_FAIL),
        )

        # ✅ 累積達標（今天）
        person_board["累積狀態"] = np.where(
            person_board["累積目標"] <= 1e-12,
            None,
            np.where(person_board["已分揀總量"] >= person_board["累積目標"], STATUS_PASS, STATUS_FAIL),
        )

        person_board["排序_未達標優先"] = np.where(
            person_board["狀態"] == STATUS_FAIL, 0,
            np.where(person_board["狀態"] == STATUS_PASS, 1, 2)
        )

        # -------------------------
        # UI
        # -------------------------
        st.divider()
        st.markdown("## 即時看板（每線一個視窗｜段數固定 1→4｜右側總量=當天該線該段累計）")

        all_lines = sorted(person_board["線別"].dropna().unique().tolist())
        if not all_lines:
            st.info("沒有線別資料")
            return

        all_zones = [1, 2, 3, 4]

        c1, c2, c3, c4 = st.columns([2.2, 1.3, 2.2, 1.3])
        with c1:
            pick_lines = st.multiselect("線別", all_lines, default=all_lines)
        with c2:
            pick_zones = st.multiselect("段數", all_zones, default=all_zones)
        with c3:
            q = st.text_input("搜尋姓名（可空白）", value="")
        with c4:
            topn = st.number_input("每條線顯示筆數", min_value=20, max_value=800, value=200, step=20)

        show_lines = [x for x in all_lines if x in pick_lines]
        if not show_lines:
            st.warning("你目前沒有選到任何線別")
            return

        tabs = st.tabs(show_lines)

        for tab, line in zip(tabs, show_lines):
            with tab:
                df_line = person_board[
                    (person_board["線別"] == line) &
                    (person_board["段數"].isin(pick_zones))
                ].copy()

                if q.strip():
                    df_line = df_line[df_line["姓名"].astype(str).str.contains(q.strip(), case=False, na=False)].copy()

                dist_now = (
                    df_line[df_line["狀態"].isin([STATUS_PASS, STATUS_FAIL])]
                    .groupby(["狀態"], as_index=False)
                    .size()
                    .rename(columns={"size": "count"})
                )
                p, f, rate = _kpi_counts(dist_now)

                a, b, c, d = st.columns(4)
                a.metric("判斷時間", now.strftime("%H:%M:%S"))
                b.metric("達標 人數(本小時)", p)
                c.metric("未達標 人數(本小時)", f)
                d.metric("達標率(本小時)", (f"{rate:.1f}%" if rate is not None else "—"))

                df_line = df_line.sort_values(["段數", "排序_未達標優先", "姓名"], ascending=[True, True, True]).head(int(topn))
                render_person_realtime_board(df_line)

                with st.expander("（進階）本線各段各人：每小時達標 Heatmap（今天）", expanded=False):
                    df_line_in = df_today_in[df_today_in["線別"] == line].copy()
                    if df_line_in.empty:
                        st.info("本線今天沒有納入計算的生產資料。")
                    else:
                        hour_cols = list(range(int(hour_min), int(cur_h) + 1)) if int(cur_h) >= int(hour_min) else [int(cur_h)]
                        base_cols = ["線別", "段數", "姓名", "開始時間"]

                        hourly_sum = df_line_in.groupby(base_cols + ["小時"], as_index=False)["加權PCS"].sum()
                        hourly_sum = hourly_sum.rename(columns={"加權PCS": "當小時加權PCS"})

                        keys = roster_df[roster_df["線別"] == line][base_cols].drop_duplicates().copy()
                        grid_hours = keys.assign(_k=1).merge(pd.DataFrame({"小時": hour_cols, "_k": 1}), on="_k").drop(columns=["_k"])
                        hourly_full = grid_hours.merge(hourly_sum, on=base_cols + ["小時"], how="left")
                        hourly_full["當小時加權PCS"] = pd.to_numeric(hourly_full["當小時加權PCS"], errors="coerce").fillna(0.0)

                        # 目標分鐘：12點=30(12:00~12:30)，13點=30(13:30~14:00)，其餘=60
                        parts = hourly_full["開始時間"].astype(str).str.split(":", n=1, expand=True)
                        s_h = pd.to_numeric(parts[0], errors="coerce").fillna(8).astype(int)
                        s_m = pd.to_numeric(parts[1], errors="coerce").fillna(0).astype(int)
                        hh = pd.to_numeric(hourly_full["小時"], errors="coerce").fillna(0).astype(int)

                        seg_s = np.where(hh == 13, 30.0, 0.0)
                        seg_e = np.where(hh == 12, 30.0, 60.0)

                        now_min_float = float(cur_m) + float(cur_s) / 60.0
                        eff_e = np.where(hh == cur_h, np.minimum(now_min_float, seg_e), seg_e)

                        p_s = np.where(hh == s_h, s_m.astype(float), 0.0)
                        mins = np.where(
                            hh > cur_h, 0.0,
                            np.where(
                                hh < s_h, 0.0,
                                np.maximum(0.0, eff_e - np.maximum(seg_s, p_s))
                            )
                        )

                        hourly_full["本小時有效分鐘"] = mins
                        hourly_full["本小時目標"] = (mins / 60.0) * float(target_hr)
                        hourly_full["狀態"] = np.where(
                            hourly_full["本小時有效分鐘"] <= 1e-12,
                            None,
                            np.where(hourly_full["當小時加權PCS"] >= hourly_full["本小時目標"], STATUS_PASS, STATUS_FAIL),
                        )

                        render_hourly_heatmap(
                            hourly_full[["線別", "段數", "姓名", "小時", "當小時加權PCS", "本小時目標", "狀態"]].copy(),
                            hour_cols=hour_cols,
                            title=f"{line}｜每小時（今天｜午休 12:30–13:30）"
                        )

    except Exception as e:
        st.error(f"發生錯誤：{e}")


if __name__ == "__main__":
    main()
