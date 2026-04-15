#!/usr/bin/env python3
"""
===================================================================
爬蟲③：multpl.com S&P 500 基本面
===================================================================
功能：爬取 S&P 500 的歷史 P/E、殖利率
用途：
  - ② Model-Driven DCA（P/E 閾值：歷史中位數 ±20%）
  - ③ Threshold DCA（殖利率 > 2.5% 低估門檻、CAPE）
產出：
  - data/fundamentals/sp500_fundamental.csv（date, PE, DividendYield, PB）
===================================================================
"""

import requests, os, re, pandas as pd, numpy as np
from bs4 import BeautifulSoup
from datetime import datetime

PROJECT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
FUND_DIR = os.path.join(PROJECT_DIR, 'data', 'fundamentals')
os.makedirs(FUND_DIR, exist_ok=True)
OUT_PATH = os.path.join(FUND_DIR, 'sp500_fundamental.csv')

HEADERS = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36'}
MONTHS = {m.lower()[:3]: i for i, m in enumerate(
    ['January','February','March','April','May','June',
     'July','August','September','October','November','December'], start=1)}


def crawl_multpl(path, name):
    """從 multpl.com 爬取一個指標的歷史數值"""
    url = f'https://www.multpl.com{path}/table/by-month'
    resp = requests.get(url, headers=HEADERS, timeout=15)
    if resp.status_code != 200:
        print(f"    ❌ {name} HTTP {resp.status_code}")
        return None
    
    soup = BeautifulSoup(resp.text, 'html.parser')
    # multpl 新版表格改為 <table id="datatable">，舊版為 class="portable-table"
    table = (
        soup.find('table', {'id': 'datatable'})
        or soup.find('table', {'class': 'portable-table'})
        or soup.find('table')
    )
    if not table:
        return None
    
    data = []
    for row in table.find_all('tr')[1:]:
        cols = row.find_all('td')
        if len(cols) >= 2:
            text = cols[0].text.strip()
            m = re.match(r'([A-Za-z]+)\s+(\d+),?\s+(\d{4})', text)
            if m:
                mon, day, year = m.groups()
                d = datetime(int(year), MONTHS.get(mon.lower()[:3], 1), min(int(day), 28))
                value_text = cols[1].text.strip().replace('%', '').replace(',', '')
                value_match = re.search(r'[-+]?\d*\.?\d+', value_text)
                if not value_match:
                    continue
                try:
                    val = float(value_match.group())
                    data.append({'date': d, 'value': val})
                except ValueError:
                    continue
    return pd.DataFrame(sorted(data, key=lambda x: x['date'])) if data else None


def fetch_sp500_fundamental():
    """
    主函式：爬取 S&P 500 的 P/E、殖利率、P/B
    
    爬蟲流程：
    1. 分別爬取三個指標（P/E、殖利率、P/B）
    2. 以日期 merge 三個指標
    3. groupby 月份取平均（對齊月線）
    4. forward fill 填補月內缺值
    """
    print("  爬取 P/E...")
    pe = crawl_multpl('/s-p-500-pe-ratio', 'P/E')
    
    print("  爬取 殖利率...")
    dy = crawl_multpl('/s-p-500-dividend-yield', '殖利率')
    
    print("  爬取 P/B...")
    pb = crawl_multpl('/s-p-500-price-to-book', 'P/B')
    
    if pe is not None:
        merged = pe.rename(columns={'value': 'PE'})
        if dy is not None:
            merged = merged.merge(dy.rename(columns={'value': 'DividendYield'}), on='date', how='outer')
        if pb is not None:
            merged = merged.merge(pb.rename(columns={'value': 'PB'}), on='date', how='outer')
        
        # 月線對齊：取每月最後一筆資料
        merged['year_month'] = merged['date'].dt.to_period('M')
        monthly = merged.groupby('year_month').agg({
            'PE': 'mean', 'DividendYield': 'mean', 'PB': 'mean'
        }).reset_index()
        monthly['date'] = monthly['year_month'].dt.to_timestamp()
        monthly = monthly[['date', 'PE', 'DividendYield', 'PB']]
        
        # 前向填補（當月缺值用上月）
        monthly['PE'] = monthly['PE'].ffill()
        monthly['DividendYield'] = monthly['DividendYield'].ffill()
        monthly['PB'] = monthly['PB'].ffill()
        
        # 從 1990-01-01 開始
        monthly = monthly[monthly['date'] >= '1990-01-01'].sort_values('date')
        monthly.to_csv(OUT_PATH, index=False)
        print(f"  ✅ {len(monthly)} 筆 → {OUT_PATH}")
        return monthly
    return None


if __name__ == '__main__':
    print("=" * 60)
    print("爬蟲③：multpl.com S&P 500 基本面")
    print("=" * 60)
    fetch_sp500_fundamental()
    print("=" * 60)
