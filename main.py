import os
import streamlit as st
from datetime import date
from core import load_config, build_brand_cfg, analyze_hh, parse_visits, run_brand
import traceback

st.set_page_config(page_title="GrowthSheet 多品牌報表", layout="wide")
st.title("GrowthSheet v1.0")

# 讀設定檔、選品牌
cfg_all = load_config()
codes = [b["code"] for b in cfg_all["brands"]]
selected = st.sidebar.selectbox("選擇品牌", codes)
brand_cfg = build_brand_cfg(cfg_all, selected)

# 子療程
subtype = None
if brand_cfg.get("subtypes"):
    opts = ["全部"] + brand_cfg["subtypes"]
    ch = st.sidebar.radio("療程類型", opts)
    subtype = None if ch == "全部" else ch

# 上傳檔案
st.header("請上傳 Excel 檔案")
hh_file = st.file_uploader("CS Booking Excel", type=["xls", "xlsx"])
visits_file = st.file_uploader("Show up Excel", type=["xls", "xlsx"])

# 參數設定
st.header("參數設定")
target_date = st.date_input("分析日期", date.today())
ad_meta = st.number_input("Meta 廣告費", min_value=0.0, value=0.0, step=1.0)
ad_google = st.number_input("Google 廣告費", min_value=0.0, value=0.0, step=1.0)

# 先预览
if hh_file and visits_file:
    st.subheader("📊 預覽結果（不寫入 Google Sheet）")
    try:
        qc, bc, br = analyze_hh(hh_file, brand_cfg["excel"], target_date, subtype)
        visits, revenue = parse_visits(visits_file, brand_cfg["visits"], target_date, subtype)
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("詢訪 QC", qc)
        col2.metric("到店 BC", bc)
        col3.metric("到店率 BR", f"{br}%")
        col4.metric("到店數 VISITS", visits)
        st.metric("實收金額 Revenue", f"{revenue}")
    except Exception as e:
        st.error(f"預覽時計算失敗：{e}")

# 真正執行寫表
if st.button("執行並寫入 Google Sheet"):
    # 驗證
    if hh_file is None:
        st.error("請先上傳 CS Booking Excel")
        st.stop()
    if visits_file is None:
        st.error("請先上傳 Show up Excel")
        st.stop()

    try:
        ok = run_brand(
            cfg=brand_cfg,
            paths={"hh": hh_file, "visits": visits_file},
            date=target_date,
            subtype=subtype,
            ad_meta=ad_meta,
            ad_google=ad_google
        )
        if ok:
            st.success("✅ 更新完成！")
        else:
            st.error("❌ 更新失敗，請檢查日誌或設定。")
    except Exception as e:
        st.error(f"執行時錯誤：{e}")
        st.text(traceback.format_exc())

