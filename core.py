import os
import json
import pandas as pd
from datetime import date as _date
from oauth2client.service_account import ServiceAccountCredentials
from gspread import authorize as gs_authorize
from difflib import get_close_matches
import streamlit as st


def load_config(path: str = "config.json") -> dict:
    if not os.path.exists(path):
        st.error(f"找不到設定檔：{path}")
        st.stop()
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def build_brand_cfg(all_cfg: dict, code: str) -> dict:
    default = all_cfg.get("default", {})
    brand = next(b for b in all_cfg.get("brands", []) if b.get("code") == code)
    merged = {}
    for section in ("excel", "visits", "google"):
        sec = {**default.get(section, {}), **brand.get(section, {})}
        sec["code"] = code
        merged[section] = sec
    merged["code"] = code
    merged["name"] = brand.get("name", code)
    if "subtypes" in brand:
        merged["subtypes"] = brand["subtypes"]
    if "sub_tab" in brand.get("google", {}):
        merged["sub_tab"] = brand["google"]["sub_tab"]
    return merged


def analyze_hh(path, cfg: dict, day: _date, subtype=None):
    raw = pd.read_excel(path, sheet_name=cfg.get("sheet_name", 0), header=None, nrows=50)
    raw = raw.fillna("").astype(str)
    keys = [cfg["contact_date_col"], cfg["last_date_col"], cfg["status_col"]]
    header_idx = None
    for i, row in raw.iterrows():
        titles = row.tolist()
        if all(any(k in cell for cell in titles) for k in keys):
            header_idx = i
            break
    if header_idx is None:
        st.error(f"在前 50 行找不到含 {keys} 的欄位。")
        return 0, 0, 0

    df = pd.read_excel(path, sheet_name=cfg.get("sheet_name", 0), header=header_idx)
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.duplicated()]

    if cfg["code"] != "CB":
        col_name = cfg.get("brand_col")
        bv = cfg.get("brand_value", "").strip()
        if isinstance(col_name, str) and col_name in df.columns and bv:
            df = df.loc[df[col_name].astype(str).str.strip() == bv]
        elif isinstance(col_name, int):
            df = df.loc[df.iloc[:, col_name].astype(str).str.strip() == bv]

    for key in ("contact_date_col", "last_date_col", "status_col"):
        exp = cfg.get(key)
        if exp not in df.columns:
            cand = get_close_matches(exp, df.columns, n=1, cutoff=0.6)
            if cand:
                cfg[key] = cand[0]
            else:
                st.error(f"缺少欄位：{exp}")
                return 0, 0, 0

    df[cfg["contact_date_col"]] = pd.to_datetime(
        df[cfg["contact_date_col"]], errors="coerce", dayfirst=cfg.get("dayfirst", True)
    ).dt.date
    df[cfg["last_date_col"]] = pd.to_datetime(
        df[cfg["last_date_col"]], errors="coerce", dayfirst=cfg.get("dayfirst", True)
    ).dt.date

    if subtype:
        if cfg["code"] == "CB":
            item_col = "項目"
            if item_col in df.columns:
                text = df[item_col].astype(str)
                if subtype == "Eye":
                    df = df.loc[text.str.contains("眼")]
                elif subtype == "Face":
                    df = df.loc[text.str.contains("面")]
        else:
            tcol = cfg.get("treatment_type_col")
            if tcol in df.columns:
                df = df.loc[df[tcol] == subtype]

    qc = len(df.loc[
        (df[cfg["contact_date_col"]] == day) &
        (df[cfg["status_col"]].isin(cfg.get("open_status", [])))
    ])
    bc = len(df.loc[
        (df[cfg["last_date_col"]] == day) &
        (df[cfg["status_col"]].isin(cfg.get("completed_status", [])))
    ])
    br = round((bc / qc) * 100, 2) if qc else 0
    return qc, bc, br


def parse_visits(path, cfg: dict, day: _date, subtype=None):
    xls = pd.ExcelFile(path)
    sheets = xls.sheet_names
    spec = cfg.get("sheet_name", 0)
    sheet = sheets[spec] if isinstance(spec, int) and 0 <= spec < len(sheets) else sheets[0]

    raw = pd.read_excel(path, sheet_name=sheet, header=None, nrows=100)
    raw = raw.fillna("").astype(str)
    keys = [cfg.get("date_col", ""), cfg.get("show_col", ""), cfg.get("no_show_col", ""), cfg.get("brand_col", "")]
    keys = [k for k in keys if k]
    header_idx = None
    for i, row in raw.iterrows():
        titles = row.tolist()
        if all(any(k in cell for cell in titles) for k in keys):
            header_idx = i
            break
    if header_idx is None:
        header_idx = cfg.get("header_row", 1) - 1

    df = pd.read_excel(path, sheet_name=sheet, header=header_idx)
    df.columns = df.columns.str.strip()
    df = df.loc[:, ~df.columns.duplicated()]

    for key in ("brand_col", "show_col", "no_show_col", "date_col", "revenue_col", "treatment_type_col"):
        val = cfg.get(key)
        if isinstance(val, str) and val and val not in df.columns:
            cand = get_close_matches(val, df.columns.astype(str), n=1, cutoff=0.6)
            if cand:
                cfg[key] = cand[0]

    def col(k):
        v = cfg.get(k)
        return df.columns[v] if isinstance(v, int) else v

    bval = str(cfg.get("brand_value", "")).strip()
    if cfg.get("code") != "CB" and bval:
        bcol = cfg.get("brand_col")
        colname = df.columns[bcol] if isinstance(bcol, int) else bcol
        if colname in df.columns:
            df = df.loc[df[colname].astype(str).str.strip() == bval]

    if cfg.get("show_col"):
        df = df.loc[df[col("show_col")] == "P"]
    if cfg.get("no_show_col"):
        df = df.loc[df[col("no_show_col")] != "P"]

    date_col = col("date_col")
    df[date_col] = pd.to_datetime(df[date_col], errors="coerce", dayfirst=True).dt.date
    df = df.loc[df[date_col] == day]

    if subtype:
        if cfg.get("code") == "CB":
            reg = "登記項目"
            if reg in df.columns:
                text = df[reg].astype(str)
                if subtype == "Eye":
                    df = df.loc[text.str.contains("眼")]
                elif subtype == "Face":
                    df = df.loc[text.str.contains("面")]
        else:
            tcol = cfg.get("treatment_type_col")
            if tcol in df.columns:
                df = df.loc[df[tcol] == subtype]

    visits = len(df)
    revenue = int(pd.to_numeric(df[col("revenue_col")], errors="coerce").fillna(0).sum())
    return visits, revenue


def write_to_sheet(gcfg: dict, paths: dict, date: _date, cfg: dict, ad_meta: float, ad_google: float):
    from gspread.exceptions import APIError
    import time

    creds = ServiceAccountCredentials.from_json_keyfile_name(
        gcfg["credentials_json"],
        ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    client = gs_authorize(creds)

    def do_tab(subtype, tab_name):
        qc, bc, br = analyze_hh(paths["hh"], cfg["excel"], date, subtype)
        visits, revenue = parse_visits(paths["visits"], cfg["visits"], date, subtype)
        ws = client.open_by_key(gcfg["sheet_id"]).worksheet(tab_name)
        data = ws.get_all_values()
        header = gcfg.get("date_col_start_row", 3)
        target = header + date.day
        while len(data) < target:
            ws.append_row([""] * ws.col_count)
            data = ws.get_all_values()
        updates = {
            gcfg["query_col"]: qc,
            gcfg["book_col"]: bc,
            gcfg["br_col"]: f"{br}%",
            gcfg["visit_col"]: visits,
            gcfg["revenue_col"]: revenue,
            gcfg["avg_revenue_col"]: round(revenue / visits, 2) if visits else 0,
            gcfg["meta_col"]: ad_meta,
            gcfg["google_col"]: ad_google,
            gcfg["total_ad_col"]: ad_meta + ad_google,
            gcfg["cpl_col"]: round((ad_meta + ad_google) / qc, 2) if qc else "",
            gcfg["cpa_book_col"]: round((ad_meta + ad_google) / bc, 2) if bc else "",
            gcfg["cpa_visit_col"]: round((ad_meta + ad_google) / visits, 2) if visits else ""
        }
        for col_idx, val in updates.items():
            attempts = 0
            while attempts < 3:
                try:
                    ws.update_cell(target, col_idx, val)
                    break
                except APIError:
                    attempts += 1
                    time.sleep(2 ** attempts)

    do_tab(None, gcfg["tab_name"])
    for stype in cfg.get("subtypes", []):
        tab = cfg.get("sub_tab", {}).get(stype, {}).get("tab_name")
        if tab:
            do_tab(stype, tab)

    return True


def run_brand(cfg: dict, paths: dict, date: _date, subtype=None, ad_meta: float = 0.0, ad_google: float = 0.0) -> bool:
    gcfg = cfg["google"].copy()
    return write_to_sheet(gcfg, paths, date, cfg, ad_meta, ad_google)
