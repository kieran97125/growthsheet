import streamlit as st
import importlib
import json
import os
from datetime import datetime

# 讀取全局 config
CONFIG_PATH = "config.json"
def load_config(path=CONFIG_PATH):
    if not os.path.exists(path):
        st.error(f"找不到設定檔：{path}")
        st.stop()
    with open(path, 'r', encoding='utf-8') as f:
        return json.load(f)
config = load_config()

# Streamlit 頁面配置
st.set_page_config(page_title="GrowthSheet 多品牌報表", layout="wide")

# 側邊欄：品牌選擇
st.sidebar.header("品牌選擇")
brands = config.get('brands', [])
if not brands:
    st.sidebar.error("未在 config.json 中設定任何品牌，請檢查配置。")
    st.stop()
brand_codes = [b['code'] for b in brands]
selected = st.sidebar.selectbox(
    "請選擇品牌",
    brand_codes,
    format_func=lambda c: next((b['name'] for b in brands if b['code']==c), c)
)
brand_cfg = next((b for b in brands if b['code']==selected), None)
if not brand_cfg:
    st.sidebar.error(f"找不到對應的品牌設定：{selected}")
    st.stop()

# 若品牌有子療程類型
subtype = None
if 'subtypes' in brand_cfg:
    st.sidebar.header("療程類型")
    subtype = st.sidebar.radio(
        "", brand_cfg['subtypes']
    )

# 主畫面輸入
st.title(f"GrowthSheet 報表工具 — {brand_cfg['name']}")

hh_path     = st.text_input("CS Booking Excel路徑 例如(XX 客人預約詳細資料 2025) 不包括引號", value=brand_cfg.get('excel_path',''))
visits_path = st.text_input("Show up Excel路徑 例如(202505 IB 每日新客show up 報表及統計) 不包括引號", value=brand_cfg.get('visits_path',''))
target_date = st.date_input("分析日期", datetime.now().date())

st.subheader("手動輸入廣告費")
ad_meta   = st.number_input("Meta 廣告費",   min_value=0.0)
ad_google = st.number_input("Google 廣告費", min_value=0.0)

# 執行按鈕
if st.button("執行並寫入 Google Sheet"):
    # 驗證檔案路徑
    if not hh_path or not os.path.exists(hh_path):
        st.error("Customer Record路徑錯誤"); st.stop()
    if not visits_path or not os.path.exists(visits_path):
        st.error("Show up路徑錯誤"); st.stop()
    if ad_meta + ad_google <= 0:
        st.error("請輸入廣告費"); st.stop()

    # 動態 import 品牌模組
    try:
        mod = importlib.import_module(f"brands.{selected}")
    except ImportError as e:
        st.error(f"無法載入品牌模組: {e}")
        st.stop()

    # 呼叫品牌模組 run() 接口
    try:
        mod.run(
            hh_path=hh_path,
            visits_path=visits_path,
            date=target_date,
            subtype=subtype,
            ad_meta=ad_meta,
            ad_google=ad_google
        )
        st.success("更新完成！")
    except Exception as e:
        st.error(f"執行失敗: {e}")
