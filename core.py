import pandas as pd
from datetime import datetime
from dateutil import parser as date_parser
from gspread import authorize as gs_authorize
from oauth2client.service_account import ServiceAccountCredentials

# ------ 核心函式庫 ------

def analyze_hh(path, cfg, date: datetime.date):
    """
    統一分析 HH Excel 的 functionality
    返回 (qc, bc, br)
    """
    df = pd.read_excel(path, sheet_name=cfg.get('sheet_name'), header=cfg['header_row'])
    df.columns = df.columns.str.strip()
    for col in [cfg['contact_date_col'], cfg['last_date_col'], cfg['status_col']]:
        if col not in df.columns:
            raise ValueError(f"HH Excel 缺少欄位：{col}")
    df[cfg['contact_date_col']] = pd.to_datetime(df[cfg['contact_date_col']], errors='coerce').dt.date
    df[cfg['last_date_col']]    = pd.to_datetime(df[cfg['last_date_col']], errors='coerce').dt.date
    q = df[(df[cfg['contact_date_col']] == date) & df[cfg['status_col']].isin(cfg.get('open_status', []))]
    b = df[(df[cfg['last_date_col']] == date) & df[cfg['status_col']].isin(cfg.get('completed_status', []))]
    qc, bc = len(q), len(b)
    br = round((bc / qc) * 100, 2) if qc else 0
    return qc, bc, br


def parse_visits(path, cfg, date: datetime.date, subtype=None):
    """
    統一分析 Visits Excel 的 functionality
    返回 (visits, revenue)
    """
    sheet = str(date.month) if cfg['sheet_name'] == 'auto' else cfg['sheet_name']
    df = pd.read_excel(path, sheet_name=sheet, header=cfg['header_row'])
    df.columns = df.columns.str.strip()

    # 1. 品牌篩選，可 support list
    bv = cfg.get('brand_value')
    if bv != 'NA':
        col = df.columns[cfg['brand_col']] if isinstance(cfg['brand_col'], int) else cfg['brand_col']
        if isinstance(bv, (list, tuple)):
            df = df[df[col].isin(bv)]
        else:
            df = df[df[col] == bv]

    # 2. 到店 & 非 No Show
    show_col = df.columns[cfg['show_col']] if isinstance(cfg['show_col'], int) else cfg['show_col']
    no_show_col = df.columns[cfg['no_show_col']] if isinstance(cfg['no_show_col'], int) else cfg['no_show_col']
    df = df[df[show_col] == 'P']
    df = df[df[no_show_col] != 'P']

    # 3. 日期篩選
    date_col = df.columns[cfg['date_col']] if isinstance(cfg['date_col'], int) else cfg['date_col']
    df[date_col] = pd.to_datetime(df[date_col], errors='coerce').dt.date
    df = df[df[date_col] == date]

    # 4. subtype
    if subtype and cfg.get('treatment_type_col'):
        df = df[df[cfg['treatment_type_col']] == subtype]

    # 5. 計算
    visits = len(df)
    revenue_col = df.columns[cfg['revenue_col']] if isinstance(cfg['revenue_col'], int) else cfg['revenue_col']
    revenue = int(pd.to_numeric(df[revenue_col], errors='coerce').fillna(0).sum())
    return visits, revenue


def write_to_sheet(gcfg, date: datetime.date, hh, visits, revenue, ad_meta, ad_google):
    """
    統一寫入 Google Sheet
    """
    creds = ServiceAccountCredentials.from_json_keyfile_name(
        gcfg['credentials_json'],
        ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    )
    client = gs_authorize(creds)
    ws = client.open_by_key(gcfg['sheet_id']).worksheet(gcfg['tab_name'])
    vals = ws.get_all_values()

    # 找到行
    row = None
    for i, r in enumerate(vals[1:], start=2):
        cell = r[gcfg['date_col'] - 1].strip()
        try:
            if cell == date.isoformat() or date_parser.parse(cell).date() == date:
                row = i
                break
        except:
            continue
    if not row:
        # 可選自動 append
        row = len(vals) + 1
        ws.append_row([date.isoformat()])

    qc, bc, br = hh
    updates = [
        (gcfg['query_col'], qc),
        (gcfg['book_col'], bc),
        (gcfg['br_col'], f"{br}%"),
        (gcfg['visit_col'], visits),
        (gcfg['revenue_col'], revenue),
        (gcfg['avg_revenue_col'], round(revenue/visits,2) if visits else 0),
        (gcfg['meta_col'], ad_meta),
        (gcfg['google_col'], ad_google),
        (gcfg['total_ad_col'], ad_meta+ad_google),
        (gcfg['cpl_col'], round((ad_meta+ad_google)/qc,2) if qc else ''),
        (gcfg['cpa_book_col'], round((ad_meta+ad_google)/bc,2) if bc else ''),
        (gcfg['cpa_visit_col'], round((ad_meta+ad_google)/visits,2) if visits else '')
    ]
    for col, val in updates:
        ws.update_cell(row, col, val)


def run_brand(cfg, paths, date, subtype=None, ad_meta=0, ad_google=0):
    """
    統一入口，可以被各品牌模組呼叫
    cfg: brand-specific config dict
    paths: dict with 'hh' 和 'visits' 路徑
    """
    hh = analyze_hh(paths['hh'], cfg['excel'], date)
    visits, revenue = parse_visits(paths['visits'], cfg['visits'], date, subtype)
    write_to_sheet(cfg['google'], date, hh, visits, revenue, ad_meta, ad_google)
