#!/usr/bin/env python3
"""
===================================================================
爬蟲①：Yahoo Finance 股價
===================================================================
功能：抓取三大市場月收盤價
用途：
  - ① Value Averaging（計算每月組合價值差距）
  - ② Model-Driven DCA（直接用股價投入）
  - ③ Threshold DCA（直接用股價投入）
產出：
  - data/taiex_price.csv     （TAIEX 台灣加權）
  - data/sp500_price.csv     （S&P 500）
  - data/nikkei225_price.csv （Nikkei 225）
  欄位：Price（月最後交易日）、Open/High/Low/Close、Adj Close（無則同 Close）、Volume
===================================================================
"""

import os
import pandas as pd
import yfinance as yf

PROJECT_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', '..'))
DATA_DIR = os.path.join(PROJECT_DIR, 'data')
os.makedirs(DATA_DIR, exist_ok=True)

# 市場對照：名稱 → Yahoo Finance 股票代碼
MARKETS = [
    ('TAIEX', '^TWII', 'taiex_price'),
    ('S&P 500', '^GSPC', 'sp500_price'),
    ('Nikkei 225', '^N225', 'nikkei225_price'),
]


def _normalize_yf_df(df):
    """清理 yfinance 回傳資料，保留 Price、OHLC、Adj Close、Volume。"""
    df = df.reset_index()
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [col[1] if isinstance(col, tuple) else col for col in df.columns]

    if 'Date' in df.columns:
        df = df.rename(columns={'Date': 'Price'})
    if df.columns[0] != 'Price':
        df = df.rename(columns={df.columns[0]: 'Price'})

    # Adj Close：部分指數月線可能無此欄，則沿用 Close
    adj_col = None
    for cand in ('Adj Close', 'AdjClose', 'adjusted_close'):
        if cand in df.columns:
            adj_col = cand
            break
    if adj_col is not None and adj_col != 'Adj Close':
        df = df.rename(columns={adj_col: 'Adj Close'})

    for col in ['Close', 'High', 'Low', 'Open', 'Volume', 'Adj Close']:
        if col not in df.columns:
            df[col] = pd.NA

    # 無還原收盤時以 Close 填（指數常見）
    mask = df['Adj Close'].isna()
    if mask.any():
        df.loc[mask, 'Adj Close'] = df.loc[mask, 'Close']

    df = df[['Price', 'Open', 'High', 'Low', 'Close', 'Adj Close', 'Volume']]

    df['Price'] = pd.to_datetime(df['Price'], errors='coerce')
    df = df.dropna(subset=['Price'])
    df['Price'] = df['Price'].dt.strftime('%Y-%m-%d')
    return df


def fetch_yahoo_prices():
    """
    從 Yahoo Finance 抓取三大市場月線價格

    使用 yfinance library，每月取收盤價（Close）。
    月線 transform：取每月最後一個交易日收盤價。
    """
    results = {}

    for name, symbol, fname in MARKETS:
        print(f"  爬取 {name}（{symbol}）...")
        ticker = yf.Ticker(symbol)
        df = ticker.history(start="2000-01-01", interval="1mo")

        if df is not None and len(df) > 0:
            df = _normalize_yf_df(df)
            out_path = os.path.join(DATA_DIR, f'{fname}.csv')
            df.to_csv(out_path, index=False)
            print(f"    ✅ {len(df)} 筆 → {out_path}")
            results[name] = df
        else:
            print(f"    ❌ 無法取得 {name}")

    return results


if __name__ == '__main__':
    print("=" * 60)
    print("爬蟲①：Yahoo Finance 股價")
    print("=" * 60)
    fetch_yahoo_prices()
    print("=" * 60)
