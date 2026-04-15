#!/usr/bin/env python3
"""
===================================================================
爬蟲④：Nikkei Indexes 日經 225 基本面
===================================================================
功能：爬取日經 225 的歷史 P/E、殖利率
用途：
  - ② Model-Driven DCA（P/E 閾值：歷史中位數 ±20%）
  - ③ Threshold DCA（殖利率 > 2.0% 低估門檻）
產出：
 - data/fundamentals/nikkei225_fundamental.csv（date, PER, DividendYield）
===================================================================
"""

import os
from datetime import datetime

import cloudscraper
import pandas as pd
import requests
from bs4 import BeautifulSoup

PROJECT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
FUND_DIR = os.path.join(PROJECT_DIR, 'data', 'fundamentals')
os.makedirs(FUND_DIR, exist_ok=True)
OUT_PATH = os.path.join(FUND_DIR, 'nikkei225_fundamental.csv')

HEADERS = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'}
BASE_URL = 'https://indexes.nikkei.co.jp/nikkeikeiinfo/v2/globalchart/ir/csv'
ARCHIVES_BASE = 'https://indexes.nikkei.co.jp/en/nkave/statistics/dataload'
START_YEAR = 2005


def _previous_month():
    now = datetime.today()
    if now.month == 1:
        return now.year - 1, 12
    return now.year, now.month - 1


def _parse_archives_date(date_str):
    return datetime.strptime(date_str, '%b/%d/%Y')


def _crawl_archives_monthly(scraper, list_name, end_year, end_month):
    """
    從新版 archives dataload 逐月抓取日資料，再回傳日頻 dict。
    list_name: per / dividend / pbr
    """
    data = {}

    for year in range(START_YEAR, end_year + 1):
        month_limit = end_month if year == end_year else 12
        for month in range(1, month_limit + 1):
            url = f'{ARCHIVES_BASE}?list={list_name}&year={year}&month={month}'
            try:
                resp = scraper.get(url, timeout=20)
            except requests.RequestException:
                continue
            if resp.status_code != 200:
                continue

            soup = BeautifulSoup(resp.text, 'html.parser')
            rows = soup.select('table tr')
            if len(rows) <= 1:
                continue

            for row in rows[1:]:
                cols = [c.get_text(strip=True) for c in row.find_all('td')]
                if len(cols) < 2:
                    continue
                date_text = cols[0]
                value_text = cols[1].replace(',', '')
                try:
                    dt = _parse_archives_date(date_text)
                    val = float(value_text)
                except ValueError:
                    continue
                data[dt] = val

    return data


def fetch_nikkei225_fundamental():
    """
    從 Nikkei Indexes API 抓取日經 225 歷史 P/E 和殖利率
    
    Nikkei Indexes 提供 CSV API，直接 HTTP GET 即可取得。
    type=per     → PER（月次）
    type=dividend → 殖利率（月次）
    
    流程：
    1. 爬 PER 和殖利率兩個 CSV
    2. 以日期 merge
    3. 整理成月線格式
    """
    def crawl_csv(chart_type, name):
        url = f'{BASE_URL}?daily=monthly&type={chart_type}'
        resp = requests.get(url, headers=HEADERS, timeout=15)
        if resp.status_code != 200:
            print(f"    ❌ {name} HTTP {resp.status_code}")
            return {}
        data = {}
        for line in resp.text.strip().split('\n')[1:]:
            parts = line.split(',')
            if len(parts) >= 2:
                try:
                    date_str = parts[0].strip('"')
                    val = float(parts[1])
                    date = datetime.strptime(date_str, '%Y/%m/%d')
                    data[date] = val
                except (ValueError, IndexError):
                    pass
        return data

    print("  爬取 Nikkei 225 PER...")
    per_dict = crawl_csv('per', 'PER')
    
    print("  爬取 Nikkei 225 殖利率...")
    div_dict = crawl_csv('dividend', '殖利率')

    if not per_dict and not div_dict:
        print("  改用新版 archives 來源...")
        end_year, end_month = _previous_month()
        scraper = cloudscraper.create_scraper(
            browser={'browser': 'chrome', 'platform': 'darwin', 'mobile': False}
        )

        print("    下載 PER（日頻）...")
        per_dict = _crawl_archives_monthly(scraper, 'per', end_year, end_month)
        print("    下載 殖利率（日頻）...")
        div_dict = _crawl_archives_monthly(scraper, 'dividend', end_year, end_month)
        print("    下載 P/B（日頻）...")
        pbr_dict = _crawl_archives_monthly(scraper, 'pbr', end_year, end_month)

        if per_dict:
            all_dates = sorted(set(per_dict.keys()) | set(div_dict.keys()) | set(pbr_dict.keys()))
            daily = pd.DataFrame({'date': all_dates})
            daily['PE'] = daily['date'].map(per_dict)
            daily['DividendYield'] = daily['date'].map(div_dict)
            daily['PB'] = daily['date'].map(pbr_dict)

            monthly = (
                daily.assign(year_month=daily['date'].dt.to_period('M'))
                .groupby('year_month', as_index=False)[['PE', 'DividendYield', 'PB']]
                .mean()
            )
            monthly['date'] = monthly['year_month'].dt.to_timestamp('M')
            monthly = monthly[['date', 'PE', 'DividendYield', 'PB']]
            monthly = monthly.dropna(subset=['PE']).sort_values('date')
            monthly.to_csv(OUT_PATH, index=False)
            print(f"  ✅ 新版來源成功 {len(monthly)} 筆 → {OUT_PATH}")
            return monthly

        fallback_path = os.path.join(FUND_DIR, 'nikkei225_fundamental_monthly.csv')
        if os.path.exists(fallback_path):
            fallback = pd.read_csv(fallback_path)
            fallback = fallback.rename(columns={
                'PER': 'PE',
                'PBR': 'PB',
                'DIVIDEND': 'DividendYield',
                'Dividend': 'DividendYield',
            })
            if 'DividendYield' in fallback.columns:
                fallback = fallback[['date', 'PE', 'DividendYield'] + ([ 'PB'] if 'PB' in fallback.columns else [])]
            else:
                fallback = fallback[['date', 'PE']]

            fallback.to_csv(OUT_PATH, index=False)
            print(f'  ✅ 用既有月頻檔案補齊 → {OUT_PATH} ({len(fallback)} 筆)')
            return fallback

        if os.path.exists(OUT_PATH):
            existing = pd.read_csv(OUT_PATH)
            if len(existing) > 1:
                print(f'    ✅ 已保留 {len(existing)} 筆現有資料 → {OUT_PATH}')
                return existing

        print('  ⚠️ API 無法取得，且沒有可用既有 Nikkei 歷史檔')
        return pd.DataFrame()

    # 合併兩個指標
    all_dates = sorted(set(per_dict.keys()) | set(div_dict.keys()))
    df = pd.DataFrame({'date': all_dates})
    df['PER'] = df['date'].map(per_dict)
    df['DividendYield'] = df['date'].map(div_dict)
    df = df.dropna(subset=['PER']).sort_values('date')
    df.to_csv(OUT_PATH, index=False)
    print(f"  ✅ {len(df)} 筆 → {OUT_PATH}")
    return df


if __name__ == '__main__':
    print("=" * 60)
    print("爬蟲④：Nikkei Indexes 日經 225 基本面")
    print("=" * 60)
    fetch_nikkei225_fundamental()
    print("=" * 60)
