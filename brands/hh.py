import json
import os
from datetime import datetime
from core import run_brand

# 讀取全局 config.json，並選取 HH 品牌設定
CONFIG_PATH = os.path.join(os.path.dirname(__file__), '..', 'config.json')
with open(CONFIG_PATH, 'r', encoding='utf-8') as f:
    cfg_all = json.load(f)
brand_cfg = next(b for b in cfg_all['brands'] if b['code'] == 'HH')


def run(hh_path=None, visits_path=None, date=None, subtype=None, ad_meta=0.0, ad_google=0.0):
    """
    執行 HairHealth (HH) 報表：
      - hh_path, visits_path: 若不傳入則使用 config.json 中設定的路徑
      - date: datetime.date, 預設今天
      - subtype: CleanBear 子療程類型 (HH 無)
      - ad_meta, ad_google: 廣告費
    """
    paths = {
        'hh': hh_path or brand_cfg.get('excel_path'),
        'visits': visits_path or brand_cfg.get('visits_path')
    }
    target_date = date or datetime.now().date()
    run_brand(
        cfg=brand_cfg,
        paths=paths,
        date=target_date,
        subtype=subtype,
        ad_meta=ad_meta,
        ad_google=ad_google
    )


if __name__ == '__main__':
    # 本地測試 entry
    run()
