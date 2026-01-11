# common_ui.py
# -*- coding: utf-8 -*-
import streamlit as st

THEME_CSS = r"""
<style>
/* ========== Base ========== */
:root{
  --bg: #f5f7fb;
  --card: #ffffff;
  --text: rgba(15, 23, 42, 0.92);
  --muted: rgba(15, 23, 42, 0.62);

  --nav-bg: #0b1220;
  --nav-bg2:#0a1020;
  --nav-text: rgba(255,255,255,.84);
  --nav-muted: rgba(255,255,255,.55);
  --nav-hover: rgba(255,255,255,.08);

  --primary: #2563eb;     /* 藍 */
  --primary2:#1d4ed8;
  --border: rgba(15,23,42,.10);
  --shadow: 0 10px 30px rgba(15,23,42,.08);
  --radius: 18px;
  --radius2: 22px;
}

/* font */
html, body, [class*="st-"], .stApp{
  font-family: ui-sans-serif, system-ui, -apple-system, "Segoe UI",
               "Noto Sans TC", "Microsoft JhengHei", Arial, sans-serif !important;
}

/* remove default streamlit chrome */
header[data-testid="stHeader"]{ display:none; }
#MainMenu{ visibility:hidden; }
footer{ visibility:hidden; }

/* App background */
.stApp{
  background: var(--bg);
}

/* ========== Top Bar (custom) ========== */
.df-topbar{
  position: fixed;
  top: 0; left: 0; right: 0;
  height: 52px;
  background: #0b0f1a;
  color: rgba(255,255,255,.92);
  z-index: 1000;
  display:flex;
  align-items:center;
  justify-content:center;
  border-bottom: 1px solid rgba(255,255,255,.06);
}
.df-topbar .wrap{
  width: min(1400px, calc(100% - 24px));
  display:flex;
  align-items:center;
  justify-content:space-between;
}
.df-topbar .title{
  font-weight: 800;
  letter-spacing: .6px;
  font-size: 14px;
}
.df-topbar .right{
  display:flex;
  align-items:center;
  gap: 10px;
  font-size: 13px;
  color: rgba(255,255,255,.80);
}
.df-pill{
  display:inline-flex;
  align-items:center;
  gap: 6px;
  padding: 6px 10px;
  border-radius: 999px;
  background: rgba(34,197,94,.14);
  border: 1px solid rgba(34,197,94,.25);
  color: rgba(255,255,255,.92);
  font-weight: 700;
  font-size: 12px;
}
.df-dot{
  width: 7px; height: 7px;
  border-radius: 50%;
  background: #22c55e;
}

/* push content below fixed topbar */
div[data-testid="stAppViewContainer"]{
  padding-top: 64px;
}

/* ========== Sidebar ========== */
section[data-testid="stSidebar"]{
  background: linear-gradient(180deg, var(--nav-bg), var(--nav-bg2));
  border-right: 1px solid rgba(255,255,255,.06);
}
section[data-testid="stSidebar"] > div{
  padding-top: 12px;
}
.df-brand{
  display:flex;
  align-items:center;
  gap: 10px;
  padding: 6px 10px 14px 10px;
  border-bottom: 1px solid rgba(255,255,255,.06);
  margin-bottom: 10px;
}
.df-brand .logo{
  width: 36px; height: 36px;
  border-radius: 12px;
  background: rgba(37,99,235,.18);
  display:flex;
  align-items:center;
  justify-content:center;
  border: 1px solid rgba(37,99,235,.30);
}
.df-brand .name{
  color: rgba(255,255,255,.92);
  font-weight: 900;
  font-size: 16px;
  line-height: 1.1;
}
.df-brand .sub{
  color: rgba(255,255,255,.55);
  font-weight: 700;
  font-size: 11px;
  letter-spacing: 1px;
}

/* section titles */
.df-nav-title{
  color: var(--nav-muted);
  font-weight: 800;
  font-size: 12px;
  letter-spacing: .8px;
  margin: 14px 10px 6px 10px;
  text-transform: none;
}

/* Streamlit buttons in sidebar -> make them look like nav items */
section[data-testid="stSidebar"] div[data-testid="stButton"] > button{
  width: 100%;
  background: transparent !important;
  border: 1px solid transparent !important;
  color: var(--nav-text) !important;
  border-radius: 12px !important;
  padding: 9px 10px !important;
  justify-content: flex-start !important;
  gap: 10px !important;
  font-weight: 750 !important;
  font-size: 14px !important;
  line-height: 1.15 !important;
}
section[data-testid="stSidebar"] div[data-testid="stButton"] > button:hover{
  background: var(--nav-hover) !important;
}

/* Selected nav item style: you set by adding df-selected class on a container */
.df-selected div[data-testid="stButton"] > button{
  background: rgba(37,99,235,.22) !important;
  border: 1px solid rgba(37,99,235,.55) !important;
  box-shadow: 0 0 0 1px rgba(37,99,235,.25) inset !important;
}

/* ========== Main page title (blue left bar) ========== */
.df-page-title{
  display:flex;
  align-items:center;
  gap: 10px;
  margin: 6px 0 12px 0;
}
.df-page-title .bar{
  width: 4px;
  height: 24px;
  border-radius: 99px;
  background: var(--primary);
}
.df-page-title .txt{
  font-size: 20px;
  font-weight: 900;
  color: var(--text);
}

/* ========== Card ========== */
.df-card{
  background: var(--card);
  border: 1px solid var(--border);
  border-radius: var(--radius2);
  box-shadow: var(--shadow);
  padding: 22px 22px;
}
.df-card-center{
  display:flex;
  flex-direction: column;
  align-items:center;
  justify-content:center;
  gap: 10px;
  text-align:center;
}
.df-cube{
  width: 64px; height: 64px;
  border-radius: 18px;
  background: rgba(37,99,235,.10);
  display:flex;
  align-items:center;
  justify-content:center;
  border: 1px solid rgba(37,99,235,.18);
}
.df-card h2{
  margin: 0;
  font-size: 28px;
  font-weight: 950;
  color: var(--text);
}
.df-card p{
  margin: 0;
  font-size: 14px;
  font-weight: 700;
  color: var(--muted);
}

/* style file uploader dropzone */
div[data-testid="stFileUploaderDropzone"]{
  border: 2px dashed rgba(15,23,42,.20) !important;
  border-radius: 18px !important;
  padding: 24px !important;
  background: rgba(15,23,42,.02) !important;
}
div[data-testid="stFileUploaderDropzone"]:hover{
  border-color: rgba(37,99,235,.55) !important;
  background: rgba(37,99,235,.05) !important;
}

/* make the uploader centered and wide */
.df-uploader-wrap{
  width: min(760px, 92%);
  margin: 12px auto 0 auto;
}
</style>
"""

def inject_df_theme(app_title: str = "大豐物流部-WMS資料自動化助手", status_text: str = "系統連線中"):
    """全站一致風格：Topbar + Sidebar + Card + Uploader"""
    st.markdown(THEME_CSS, unsafe_allow_html=True)
    st.markdown(
        f"""
        <div class="df-topbar">
          <div class="wrap">
            <div class="title">{app_title}</div>
            <div class="right">
              <span>裝置</span>
              <span class="df-pill"><span class="df-dot"></span>{status_text}</span>
            </div>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )

def sidebar_brand():
    st.sidebar.markdown(
        """
        <div class="df-brand">
          <div class="logo">🚚</div>
          <div>
            <div class="name">大豐物流部</div>
            <div class="sub">OPERATIONAL BI PLATFORM</div>
          </div>
        </div>
        """,
        unsafe_allow_html=True
    )

def page_title(text: str):
    st.markdown(
        f"""
        <div class="df-page-title">
          <div class="bar"></div>
          <div class="txt">{text}</div>
        </div>
        """,
        unsafe_allow_html=True
    )

def card_open():
    st.markdown('<div class="df-card">', unsafe_allow_html=True)

def card_close():
    st.markdown("</div>", unsafe_allow_html=True)
