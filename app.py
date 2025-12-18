import streamlit as st
from common_ui import inject_purple_theme

st.set_page_config(
    page_title="工作效率平台",
    page_icon="🏭",
    layout="wide",
)

inject_purple_theme()


def home():
    st.title("🏭 工作效率平台")
    st.markdown(
        """
### 左側選單可切換不同項目

- ✅ **驗收達標效率**（含空窗 / AM-PM / 排除區間）
- 📦 **總上組上架產能**（含空窗 / AM-PM / 報表區塊 / 休息規則）

---

**操作流程：**  
📤 上傳檔案 → ⚙️ 設定參數（如需） → 🚀 開始計算 → ⬇️ 下載 Excel
"""
    )
    st.info("請由左側選單選擇要查看的功能項目。")


# 用官方導航 API 自訂左側頁籤名稱與 icon
pg = st.navigation(
    [
        st.Page(home, title="工作效率平台", icon="🏭", default=True),
        st.Page("pages/1_驗收達標效率.py", title="驗收達標效率", icon="✅"),
        st.Page("pages/2_總上組上架產能.py", title="總上組上架產能", icon="📦"),
    ]
)

pg.run()

