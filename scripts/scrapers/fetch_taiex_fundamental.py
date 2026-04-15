#!/usr/bin/env python3
"""
===================================================================
爬蟲②：Goodinfo 台股 P/E + P/B（Playwright）
===================================================================
功能：抓取台灣加權指數的歷史 P/E 和 P/B
用途：
  - ② Model-Driven DCA（P/E 閾值：歷史中位數 ±20%）
  - ③ Threshold DCA（P/B < 1.5 低估門檻、P/E > 25 高估門檻）
產出：
  - data/fundamentals/taiex_fundamental.csv（date, PE, PB）
===================================================================
"""

import os
import re
from datetime import datetime, timedelta
from urllib.parse import quote

import pandas as pd
from playwright.sync_api import sync_playwright

PROJECT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
FUND_DIR = os.path.join(PROJECT_DIR, 'data', 'fundamentals')
os.makedirs(FUND_DIR, exist_ok=True)
OUT_PATH = os.path.join(FUND_DIR, 'taiex_fundamental.csv')

STOCK_ID = '0000'
START_DT = '2006-01-01'


def _month_end_from_str(dt):
    if dt.month == 12:
        next_month = datetime(dt.year + 1, 1, 1)
    else:
        next_month = datetime(dt.year, dt.month + 1, 1)
    return next_month - timedelta(days=1)


def _roc_to_ad(roc_str):
    roc_year = int(roc_str)
    # 兩位數年代目前用 2000 年代，三位數用民國年
    if roc_year <= 99:
        return 2000 + roc_year
    return 1911 + roc_year


def _parse_chart_rows(rows):
    """將 #tblDetail 的資料列轉成 date/value list。"""
    data = []
    for row in rows:
        cols = [cell.inner_text().strip() for cell in row.locator('td').all()]
        if len(cols) < 5:
            continue

        date_text = cols[0]
        value_text = cols[4]  # 目前 PER / PBR 欄位位置
        if value_text in {'', '-', 'N/A', 'n/a'}:
            continue

        match = re.match(r'^(\d{2,3})M(\d{1,2})$', date_text)
        if not match:
            continue

        try:
            roc_year = int(match.group(1))
            month = int(match.group(2))
            year = _roc_to_ad(roc_year)
            value = float(value_text.replace(',', ''))
        except (TypeError, ValueError):
            continue

        date = f"{year:04d}-{month:02d}-01"
        data.append({'date': date, 'value': value})

    return data


def fetch_chart_data(browser, rpt_cat):
    """抓某一種類別（PER / PBR）的月度歷史值。"""
    start_dt = datetime.strptime(START_DT, '%Y-%m-%d')
    now = datetime.today()
    if now.month == 1:
        prev_month_start = datetime(now.year - 1, 12, 1)
    else:
        prev_month_start = datetime(now.year, now.month - 1, 1)
    end_dt = _month_end_from_str(prev_month_start)

    collected = {}
    while True:
        chart_url = (
            'https://goodinfo.tw/tw/ShowK_ChartFlow.asp'
            f'?RPT_CAT={rpt_cat}'
            f'&STOCK_ID={quote(STOCK_ID)}'
            '&CHT_CAT=MONTH'
            '&PRICE_ADJ=F'
            '&STEP=DATA'
            f'&START_DT={START_DT}'
            f'&END_DT={end_dt.strftime("%Y-%m-%d")}'
        )

        page = browser.new_page()
        page.goto(chart_url, wait_until='domcontentloaded', timeout=60000)
        # `wait_for_function` 在多次導頁時會偶發 context destroyed，改用固定重試 polling。
        rows = []
        for _ in range(60):
            page.wait_for_timeout(500)
            rows = page.locator('#tblDetail tbody tr').all()
            if len(rows) > 20:
                break

        if not rows:
            page.close()
            break

        chunk = _parse_chart_rows(rows)
        if not chunk:
            page.close()
            break

        page.close()

        for item in chunk:
            collected.setdefault(item['date'], item['value'])

        oldest = chunk[-1]['date']
        if oldest <= START_DT:
            break
        if len(chunk) < 240:
            break

        oldest_dt = datetime.strptime(oldest, '%Y-%m-%d')
        if oldest_dt <= start_dt:
            break

        prev_month = oldest_dt.replace(day=1) - timedelta(days=1)
        end_dt = _month_end_from_str(prev_month)
        if end_dt < start_dt:
            break

    if not collected:
        raise RuntimeError(f'{rpt_cat} 抓不到表格資料，請確認頁面樣式是否變更')

    return [{'date': k, 'value': v} for k, v in collected.items()]


def fetch_taiex_fundamental():
    """
    從 Goodinfo 抓取台灣加權指數的歷史 P/E 和 P/B

    為何用 Playwright：
    - Goodinfo 有 Cloudflare 保護，傳統 requests 會被阻擋
    - 本頁面用 `ShowK_ChartFlow.asp` + `STEP=DATA` 可直接回傳資料
    - Playwright headless Chromium 可保持與瀏覽器一致的抓取邏輯
    """
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=True, args=['--disable-blink-features=AutomationControlled'])
        print('  抓取 PER...')
        pe_data = fetch_chart_data(browser, 'PER')

        print('  抓取 PBR...')
        pb_data = fetch_chart_data(browser, 'PBR')

        browser.close()

    pe_df = pd.DataFrame(pe_data).rename(columns={'value': 'PE'})
    pb_df = pd.DataFrame(pb_data).rename(columns={'value': 'PB'})

    df = pd.merge(pe_df, pb_df, on='date', how='outer').sort_values('date')
    df = df.dropna(subset=['PE'])

    df.to_csv(OUT_PATH, index=False)
    print(f"  ✅ {len(df)} 筆 → {OUT_PATH}")
    return df


if __name__ == '__main__':
    print('=' * 60)
    print('爬蟲②：Goodinfo 台股 P/E + P/B（Playwright）')
    print('=' * 60)
    fetch_taiex_fundamental()
    print('=' * 60)
