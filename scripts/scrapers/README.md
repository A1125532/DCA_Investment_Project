# DCA 爬蟲程式說明

## 資料來源對照表

| 市場 | 股價來源 | 基本面來源 | 爬蟲技術 |
|------|---------|-----------|---------|
| TAIEX | Yahoo Finance (^TWII) | TWSE C02001 xlsx（優先）/ Goodinfo（備援） | pandas(openpyxl) / Playwright |
| S&P 500 | Yahoo Finance (^GSPC) | multpl.com | requests + BeautifulSoup |
| Nikkei 225 | Yahoo Finance (^N225) | Nikkei Indexes API | HTTP API |

## 檔案清單

```
scrapers/
├── README.md                    # 本檔案
├── fetch_yahoo_prices.py        # Yahoo Finance 股價（所有市場）
├── import_taiex_from_twse_xlsx.py # TWSE C02001 xlsx 匯入（PE/PB/殖利率）
├── build_twse_monthly_valuation.py # C02001 每月資料下載+整理
├── fetch_twse_march_c02001.py # C02001 三月版 ZIP 抓取+解壓
├── extract_twse_valuation_march.py # 三月 Form1 擷取成 xlsx
├── fetch_taiex_fundamental.py # Goodinfo 台股 P/E + P/B（備援）
├── fetch_sp500_fundamental.py   # multpl.com S&P 500 P/E + 殖利率
└── fetch_nikkei_fundamental.py # Nikkei Indexes 日股 P/E + 殖利率
```

## 使用方式

```bash
# 個別執行
python3 scripts/scrapers/fetch_yahoo_prices.py
python3 scripts/scrapers/import_taiex_from_twse_xlsx.py
python3 scripts/scrapers/fetch_taiex_fundamental.py
python3 scripts/scrapers/fetch_sp500_fundamental.py
python3 scripts/scrapers/fetch_nikkei_fundamental.py

# 或一次執行全部
python3 scripts/scrapers/run_all.py
```

## 產出檔案

```
data/
├── taiex_price.csv             # Yahoo Finance 股價（月線）
├── sp500_price.csv
├── nikkei225_price.csv
└── fundamentals/
    ├── taiex_fundamental.csv  # 優先為 TWSE xlsx 匯入（含殖利率）；否則 Goodinfo
    ├── sp500_fundamental.csv  # multpl.com P/E + 殖利率
    └── nikkei225_fundamental.csv  # Nikkei Indexes P/E + 殖利率
```

台股 C02001 xlsx 預設位置：
`data/twse_c02001_monthly/output`
