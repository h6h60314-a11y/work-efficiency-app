# pages/29_各時段作業效率.py
# -*- coding: utf-8 -*-
from __future__ import annotations

import io
import os
from io import StringIO
from datetime import datetime, date, timedelta
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

# ✅ 特殊工時（分鐘）：12點、13點只有 30 分鐘
WORK_MINUTES_BY_HOUR = {12: 30, 13: 30}


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


def _slot_minutes(hour: int) -> int:
    return int(WORK_MINUTES_BY_HOUR.get(int(hour), 60))


def _time_str_min(t: str) -> int:
    t = _safe_time(t)
    hh, mm = t.split(":")
    return int(hh) * 60 + int(mm)


# =============================
# KPI 計數（某小時）
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
# ✅ Excel：保留公式 + 色塊（條件格式）
# =============================
def build_excel_bytes_with_formulas_and_colors(
    detail_df: pd.DataFrame,
    roster_df: pd.DataFrame,
    hour_cols: list[int],
    target_hr: float,
    now_h: int,
    now_m: int,
) -> bytes:
    wb = Workbook()
    ws_detail = wb.active
    ws_detail.title = "完整明細_去重後"
    ws_mat = wb.create_sheet("時段量體_公式")
    ws_param = wb.create_sheet("參數")

    ws_param["A1"] = "now_h"; ws_param["B1"] = int(now_h)
    ws_param["A2"] = "now_m"; ws_param["B2"] = int(now_m)
    ws_param["A3"] = "target_hr"; ws_param["B3"] = float(target_hr)
    for r in range(1, 4):
        ws_param[f"A{r}"].font = Font(bold=True)

    cols = list(detail_df.columns)
    for c_idx, col in enumerate(cols, start=1):
        ws_detail.cell(row=1, column=c_idx, value=col).font = Font(bold=True)

    col_pack = cols.index("PACKQTY") + 1 if "PACKQTY" in cols else None
    col_w = cols.index("Cweight") + 1 if "Cweight" in cols else None
    col_aw = cols.index("加權PCS") + 1 if "加權PCS" in cols else None

    for r_idx, row in enumerate(detail_df.itertuples(index=False), start=2):
        for c_idx, col in enumerate(cols, start=1):
            v = getattr(row, col) if hasattr(row, col) else None
            ws_detail.cell(row=r_idx, column=c_idx, value=v)

        if col_pack and col_w and col_aw:
            p_cell = f"{get_column_letter(col_pack)}{r_idx}"
            w_cell = f"{get_column_letter(col_w)}{r_idx}"
            ws_detail.cell(row=r_idx, column=col_aw, value=f"={p_cell}*{w_cell}")
            ws_detail.cell(row=r_idx, column=col_aw).number_format = "0.0000"

    detail_header_to_col = {ws_detail.cell(row=1, column=i).value: i for i in range(1, ws_detail.max_column + 1)}
    need = ["線別", "段數", "小時", "加權PCS", "納入計算"]
    for k in need:
        if k not in detail_header_to_col:
            raise ValueError(f"明細缺少欄位「{k}」，無法建立 SUMIFS 公式。")

    d_line = get_column_letter(detail_header_to_col["線別"])
    d_zone = get_column_letter(detail_header_to_col["段數"])
    d_hour = get_column_letter(detail_header_to_col["小時"])
    d_aw = get_column_letter(detail_header_to_col["加權PCS"])
    d_in = get_column_letter(detail_header_to_col["納入計算"])
    d_first, d_last = 2, ws_detail.max_row

    base_cols = ["線別", "段數", "姓名", "開始時間"]
    hour_cols = [int(h) for h in hour_cols]

    headers = base_cols[:]
    for h in hour_cols:
        headers += [str(h), f"{h}_目標", f"{h}_狀態"]
    headers += ["加總", "加總目標", "加總狀態"]

    for c_idx, h in enumerate(headers, start=1):
        ws_mat.cell(row=1, column=c_idx, value=h).font = Font(bold=True)

    fill_ok = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    fill_ng = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    fill_na = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")

    now_h_cell = "參數!$B$1"
    now_m_cell = "參數!$B$2"
    target_hr_cell = "參數!$B$3"

    for r_idx, row in enumerate(roster_df.itertuples(index=False), start=2):
        ws_mat.cell(row=r_idx, column=1, value=row.線別)
        ws_mat.cell(row=r_idx, column=2, value=int(row.段數))
        ws_mat.cell(row=r_idx, column=3, value=str(row.姓名))

        hh, mm = str(row.開始時間).split(":")
        ws_mat.cell(row=r_idx, column=4, value=f"=TIME({int(hh)},{int(mm)},0)")
        start_time_cell = f"$D{r_idx}"

        col_ptr = 5
        sum_cells, tgt_cells = [], []

        for h in hour_cols:
            vol_col = col_ptr
            tgt_col = vol_col + 1
            st_col = vol_col + 2

            vol_cell = f"{get_column_letter(vol_col)}{r_idx}"
            tgt_cell = f"{get_column_letter(tgt_col)}{r_idx}"

            line_cell = f"$A{r_idx}"
            zone_cell = f"$B{r_idx}"
            vol_formula = (
                f'=SUMIFS('
                f'\'{ws_detail.title}\'!${d_aw}${d_first}:${d_aw}${d_last},'
                f'\'{ws_detail.title}\'!${d_line}${d_first}:${d_line}${d_last},{line_cell},'
                f'\'{ws_detail.title}\'!${d_zone}${d_first}:${d_zone}${d_last},{zone_cell},'
                f'\'{ws_detail.title}\'!${d_hour}${d_first}:${d_hour}${d_last},{h},'
                f'\'{ws_detail.title}\'!${d_in}${d_first}:${d_in}${d_last},TRUE)'
            )
            ws_mat.cell(row=r_idx, column=vol_col, value=vol_formula).number_format = "0.0000"

            slot = f'IF(OR({h}=12,{h}=13),30,60)'
            endm = f'IF({h}={now_h_cell},MIN({now_m_cell},{slot}),{slot})'
            sh = f'HOUR({start_time_cell})'
            sm = f'MINUTE({start_time_cell})'
            mins = f'IF({h}>{now_h_cell},0,IF({h}<{sh},0,IF({h}={sh},MAX(0,{endm}-{sm}),{endm})))'
            tgt_formula = f'={target_hr_cell}*({mins})/60'
            ws_mat.cell(row=r_idx, column=tgt_col, value=tgt_formula).number_format = "0.0000"

            st_formula = f'=IF({tgt_cell}<=0,"",IF({vol_cell}>={tgt_cell},"{STATUS_PASS}","{STATUS_FAIL}"))'
            ws_mat.cell(row=r_idx, column=st_col, value=st_formula)

            sum_cells.append(vol_cell)
            tgt_cells.append(tgt_cell)
            col_ptr += 3

        sum_col = col_ptr
        sum_tgt_col = col_ptr + 1
        sum_st_col = col_ptr + 2

        sum_cell = f"{get_column_letter(sum_col)}{r_idx}"
        sum_tgt_cell = f"{get_column_letter(sum_tgt_col)}{r_idx}"

        ws_mat.cell(row=r_idx, column=sum_col, value=f"=SUM({','.join(sum_cells)})").number_format = "0.0000"
        ws_mat.cell(row=r_idx, column=sum_tgt_col, value=f"=SUM({','.join(tgt_cells)})").number_format = "0.0000"
        ws_mat.cell(row=r_idx, column=sum_st_col, value=f'=IF({sum_tgt_cell}<=0,"",IF({sum_cell}>={sum_tgt_cell},"{STATUS_PASS}","{STATUS_FAIL}"))')

    for row in ws_mat.iter_rows(min_row=1, max_row=ws_mat.max_row, min_col=1, max_col=ws_mat.max_column):
        for cell in row:
            cell.alignment = Alignment(horizontal="center", vertical="center")

    start_col = 5
    for i, _h in enumerate(hour_cols):
        vol_col = start_col + i * 3
        tgt_col = vol_col + 1
        st_col = vol_col + 2
        ws_mat.column_dimensions[get_column_letter(tgt_col)].hidden = True
        ws_mat.column_dimensions[get_column_letter(st_col)].hidden = True
        ws_mat.column_dimensions[get_column_letter(vol_col)].width = 10

    sum_col = start_col + len(hour_cols) * 3
    sum_tgt_col = sum_col + 1
    sum_st_col = sum_col + 2
    ws_mat.column_dimensions[get_column_letter(sum_col)].width = 12
    ws_mat.column_dimensions[get_column_letter(sum_tgt_col)].hidden = True
    ws_mat.column_dimensions[get_column_letter(sum_st_col)].hidden = True

    max_r = ws_mat.max_row
    for i, _h in enumerate(hour_cols):
        vol_col = start_col + i * 3
        st_col = vol_col + 2
        vol_letter = get_column_letter(vol_col)
        st_letter = get_column_letter(st_col)
        rng = f"{vol_letter}2:{vol_letter}{max_r}"
        ws_mat.conditional_formatting.add(rng, FormulaRule(formula=[f'${st_letter}2="{STATUS_PASS}"'], fill=fill_ok, stopIfTrue=True))
        ws_mat.conditional_formatting.add(rng, FormulaRule(formula=[f'${st_letter}2="{STATUS_FAIL}"'], fill=fill_ng, stopIfTrue=True))
        ws_mat.conditional_formatting.add(rng, FormulaRule(formula=[f'${st_letter}2=""'], fill=fill_na, stopIfTrue=True))

    sum_letter = get_column_letter(sum_col)
    sum_st_letter = get_column_letter(sum_st_col)
    sum_rng = f"{sum_letter}2:{sum_letter}{max_r}"
    ws_mat.conditional_formatting.add(sum_rng, FormulaRule(formula=[f'${sum_st_letter}2="{STATUS_PASS}"'], fill=fill_ok, stopIfTrue=True))
    ws_mat.conditional_formatting.add(sum_rng, FormulaRule(formula=[f'${sum_st_letter}2="{STATUS_FAIL}"'], fill=fill_ng, stopIfTrue=True))
    ws_mat.conditional_formatting.add(sum_rng, FormulaRule(formula=[f'${sum_st_letter}2=""'], fill=fill_na, stopIfTrue=True))

    out = io.BytesIO()
    wb.save(out)
    return out.getvalue()


# =============================
# ✅ 即時看板 UI（每線每段每人）
# ✅ 段數固定 1→2→3→4（第一順位）
# =============================
def _board_css():
    st.markdown(
        """
<style>
.board-row { display: grid; grid-template-columns: 140px 160px 1fr 220px; gap: 14px;
             align-items: center; padding: 10px 10px; border-bottom: 1px solid rgba(0,0,0,0.08); }
.board-line { font-size: 40px; font-weight: 900; color:#1e40af; line-height: 1; }
.board-start { font-size: 18px; font-weight: 700; color:#111827; margin-top:6px; }
.board-hourly { font-size: 60px; font-weight: 900; color:#111827; text-align: right; line-height: 1; }
.board-total { font-size: 60px; font-weight: 900; color:#111827; text-align: right; line-height: 1; }
.board-bar { width: 100%; height: 34px; background: #e5e7eb; border-radius: 4px; overflow: hidden; position: relative;}
.board-bar > .fill { height: 100%; width: 0%; }
.board-bar > .txt { position:absolute; inset:0; display:flex; align-items:center; justify-content:center;
                    font-size:20px; font-weight:900; color:#111827; }
.board-head { display:grid; grid-template-columns: 140px 160px 1fr 220px; gap: 14px;
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
  <div class="h">每小時分揀量</div>
  <div></div>
  <div class="hr">已分揀總量</div>
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
        total = float(r.get("已分揀總量", 0.0))
        tgt = float(r.get("本小時目標", 0.0))
        status = r.get("狀態", None)

        pct = 0.0
        if tgt > 0:
            pct = max(0.0, min(hourly / tgt, 1.0)) * 100.0

        if status == STATUS_FAIL:
            bar_color = "#dc2626"; dot = "#dc2626"
        elif status == STATUS_PASS:
            bar_color = "#16a34a"; dot = "#16a34a"
        else:
            bar_color = "#9ca3af"; dot = "#9ca3af"

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
      <div class="txt">目標 {tgt:,.0f}</div>
    </div>
  </div>

  <div>
    <div class="board-total">{total:,.0f}</div>
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

    st.markdown("### ⏱️ 各時段作業效率（每線×每段×每人 即時產能｜段數固定 1→4）")

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

    # ✅ Sidebar 只保留必要設定（你要移除的「截止時間/自動刷新/立即刷新」整段已刪掉）
    with st.sidebar:
        st.markdown("### 設定")
        target_hr = st.number_input("每小時目標（加權PCS/小時）", min_value=1.0, value=790.0, step=10.0)
        hour_min = st.number_input("起始小時（Heatmap 用）", min_value=0, max_value=23, value=8, step=1)
        lookback_min = st.number_input("只保留最近N分鐘資料（避免太大）", min_value=10, max_value=24 * 60, value=180, step=10)
        st.caption("本頁固定用『資料最新時間』作為判斷截止（避免沒有資料）。")

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
        # 人員名單解析
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
        roster_df = roster_df.drop_duplicates(["線別", "段數"], keep="first").copy()
        roster_df = roster_df[["線別", "段數", "姓名", "開始時間"]].copy()

        # 生產資料：多檔合併
        frames = []
        for f in prod_files:
            raw = f.getvalue()
            dfp = read_table_robust(f.name, raw, label=f"生產資料（{f.name}）")
            dfp["__source__"] = f.name
            frames.append(dfp)

        df_raw = _norm_cols(pd.concat(frames, ignore_index=True))
        require_columns(df_raw, ["PICKDATE", "LINEID", "ZONEID", "PACKQTY", "Cweight"], "生產資料（合併）")

        df_raw["PICKDATE"] = pd.to_datetime(df_raw["PICKDATE"], errors="coerce")
        df_raw = df_raw[df_raw["PICKDATE"].notna()].copy()

        max_ts = df_raw["PICKDATE"].max()
        cutoff = max_ts - timedelta(minutes=int(lookback_min))
        df_raw = df_raw[df_raw["PICKDATE"] >= cutoff].copy()

        df_raw = df_raw.rename(columns={"LINEID": "線別", "ZONEID": "段數"})
        df_raw["線別"] = clean_line(df_raw["線別"])
        df_raw["段數"] = clean_zone_1to4(df_raw["段數"])
        df_raw = df_raw[df_raw["段數"].notna()].copy()
        df_raw["段數"] = df_raw["段數"].astype(int)

        df_raw["PACKQTY"] = pd.to_numeric(df_raw["PACKQTY"], errors="coerce").fillna(0)
        df_raw["Cweight"] = pd.to_numeric(df_raw["Cweight"], errors="coerce").fillna(0)

        # 去重（避免多檔重疊）
        rid_cols = [c for c in df_raw.columns if c not in ("__rid",)]
        df_raw["__rid"] = pd.util.hash_pandas_object(df_raw[rid_cols], index=False)
        df_raw = df_raw.drop_duplicates("__rid", keep="first").copy()

        df = pd.merge(df_raw, roster_df, on=["線別", "段數"], how="left", validate="m:1")
        df["姓名"] = df["姓名"].fillna("未設定")
        df["開始時間"] = df["開始時間"].fillna("08:00").map(_safe_time)

        df["小時"] = df["PICKDATE"].dt.hour
        df["PICK_MIN"] = df["PICKDATE"].dt.hour * 60 + df["PICKDATE"].dt.minute

        st_parts = df["開始時間"].astype(str).str.split(":", n=1, expand=True)
        st_h = pd.to_numeric(st_parts[0], errors="coerce").fillna(8).astype(int)
        st_m = pd.to_numeric(st_parts[1], errors="coerce").fillna(0).astype(int)
        df["開始分鐘"] = st_h * 60 + st_m

        df["納入計算"] = df["PICK_MIN"] >= df["開始分鐘"]
        df["排除原因"] = np.where(df["納入計算"], "", "早於開始時間")
        df["加權PCS"] = df["PACKQTY"] * df["Cweight"]

        df_in = df[df["納入計算"]].copy()

        # ✅ 固定用資料最新時間當截止（避免沒有資料）
        eval_now = max_ts
        cur_h, cur_m = int(eval_now.hour), int(eval_now.minute)

        st.caption(
            f"資料時間：{cutoff.strftime('%Y-%m-%d %H:%M')} ～ {max_ts.strftime('%Y-%m-%d %H:%M')}｜"
            f"判斷截止：{eval_now.strftime('%Y-%m-%d %H:%M')}（資料最新）｜"
            f"來源檔：{df_in['__source__'].nunique()} 個｜（12/13 小時只算 30 分鐘）"
        )

        # 本小時 / 累積（每線×每段）
        hourly_person = (
            df_in[df_in["小時"] == cur_h]
            .groupby(["線別", "段數"], as_index=False)["加權PCS"].sum()
            .rename(columns={"加權PCS": "每小時分揀量"})
        )
        total_person = (
            df_in.groupby(["線別", "段數"], as_index=False)["加權PCS"].sum()
            .rename(columns={"加權PCS": "已分揀總量"})
        )

        person_board = pd.merge(total_person, hourly_person, on=["線別", "段數"], how="left")
        person_board["每小時分揀量"] = person_board["每小時分揀量"].fillna(0.0)

        person_board = pd.merge(person_board, roster_df, on=["線別", "段數"], how="left")
        person_board["姓名"] = person_board["姓名"].fillna("未設定")
        person_board["開始時間"] = person_board["開始時間"].fillna("08:00").map(_safe_time)

        slot = _slot_minutes(cur_h)
        end_m = min(int(cur_m), int(slot))

        s_min = person_board["開始時間"].map(_time_str_min).astype(int)
        s_hh = (s_min // 60).astype(int)
        s_mm = (s_min % 60).astype(int)

        mins_worked = np.where(
            cur_h < s_hh, 0,
            np.where(cur_h == s_hh, np.maximum(0, end_m - s_mm), end_m)
        ).astype(float)

        person_board["本小時有效分鐘"] = mins_worked
        person_board["本小時目標"] = (mins_worked / 60.0) * float(target_hr)

        person_board["狀態"] = np.where(
            person_board["本小時有效分鐘"] <= 0,
            None,
            np.where(person_board["每小時分揀量"] >= person_board["本小時目標"], STATUS_PASS, STATUS_FAIL)
        )

        person_board["排序_未達標優先"] = np.where(
            person_board["狀態"] == STATUS_FAIL, 0,
            np.where(person_board["狀態"] == STATUS_PASS, 1, 2)
        )

        st.divider()
        st.markdown("## 即時看板（每線一個視窗｜段數固定 1→4）")

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
                a.metric("判斷小時", f"{cur_h:02d} 點")
                b.metric("達標 人數", p)
                c.metric("未達標 人數", f)
                d.metric("達標率", (f"{rate:.1f}%" if rate is not None else "—"))

                df_line = df_line.sort_values(["段數", "排序_未達標優先", "姓名"], ascending=[True, True, True]).head(int(topn))
                render_person_realtime_board(df_line)

                with st.expander("（進階）本線各段各人：每小時達標 Heatmap", expanded=False):
                    hour_cols = list(range(int(hour_min), int(cur_h) + 1)) if int(cur_h) >= int(hour_min) else [int(cur_h)]
                    base_cols = ["線別", "段數", "姓名", "開始時間"]

                    df_line_in = df_in[df_in["線別"] == line].copy()
                    hourly_sum = df_line_in.groupby(base_cols + ["小時"], as_index=False)["加權PCS"].sum()
                    hourly_sum = hourly_sum.rename(columns={"加權PCS": "當小時加權PCS"})

                    keys = roster_df[roster_df["線別"] == line][base_cols].drop_duplicates().copy()
                    grid_hours = keys.assign(_k=1).merge(pd.DataFrame({"小時": hour_cols, "_k": 1}), on="_k").drop(columns=["_k"])
                    hourly_full = grid_hours.merge(hourly_sum, on=base_cols + ["小時"], how="left")
                    hourly_full["當小時加權PCS"] = pd.to_numeric(hourly_full["當小時加權PCS"], errors="coerce").fillna(0.0)

                    parts = hourly_full["開始時間"].astype(str).str.split(":", n=1, expand=True)
                    s_h = pd.to_numeric(parts[0], errors="coerce").fillna(8).astype(int)
                    s_m = pd.to_numeric(parts[1], errors="coerce").fillna(0).astype(int)
                    hh = pd.to_numeric(hourly_full["小時"], errors="coerce").fillna(0).astype(int)

                    slot_arr = hh.map(lambda x: _slot_minutes(int(x))).astype(int)
                    end_m_arr = np.where(hh == cur_h, np.minimum(cur_m, slot_arr), slot_arr).astype(int)

                    minutes_worked = np.where(
                        hh > cur_h, 0,
                        np.where(hh < s_h, 0, np.where(hh == s_h, np.maximum(0, end_m_arr - s_m), end_m_arr))
                    ).astype(float)

                    hourly_full["本小時有效分鐘"] = minutes_worked
                    hourly_full["本小時目標"] = (minutes_worked / 60.0) * float(target_hr)
                    hourly_full["狀態"] = np.where(
                        hourly_full["本小時有效分鐘"] <= 0,
                        None,
                        np.where(hourly_full["當小時加權PCS"] >= hourly_full["本小時目標"], STATUS_PASS, STATUS_FAIL)
                    )

                    render_hourly_heatmap(
                        hourly_full[["線別", "段數", "姓名", "小時", "當小時加權PCS", "本小時目標", "狀態"]].copy(),
                        hour_cols=hour_cols,
                        title=f"{line}｜每小時（12/13=30分）"
                    )

        st.divider()
        st.markdown("## 匯出 Excel（保留公式＋色塊自動更新）")

        detail_df = df.copy().sort_values(["線別", "段數", "PICKDATE"]).reset_index(drop=True)
        if "加權PCS" not in detail_df.columns:
            detail_df["加權PCS"] = np.nan

        hour_cols = list(range(int(hour_min), int(cur_h) + 1)) if int(cur_h) >= int(hour_min) else [int(cur_h)]
        xlsx_bytes = build_excel_bytes_with_formulas_and_colors(
            detail_df=detail_df,
            roster_df=roster_df,
            hour_cols=hour_cols,
            target_hr=float(target_hr),
            now_h=int(cur_h),
            now_m=int(cur_m),
        )
        filename = f"產能時段_公式_色塊_{datetime.now(TPE).strftime('%Y%m%d_%H%M')}.xlsx"
        st.download_button(
            "⬇️ 下載 Excel（保留公式＋色塊自動變色）",
            data=xlsx_bytes,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    except Exception as e:
        st.error(f"發生錯誤：{e}")


if __name__ == "__main__":
    main()
