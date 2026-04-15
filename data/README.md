# Data Folder Guide

這個資料夾目前分成「主流程必需檔」與「中間產物」。

## 1) 主流程必需檔（不要改路徑）
- `taiex_price.csv`
- `sp500_price.csv`
- `nikkei225_price.csv`
- `fundamentals/taiex_fundamental.csv`
- `fundamentals/sp500_fundamental.csv`
- `fundamentals/nikkei225_fundamental.csv`

## 2) 校準/相容仍會用到的 fundamentals 檔
- `fundamentals/sp500_fundamental_monthly.csv`
- `fundamentals/sp500_fundamental_merged.csv`
- `fundamentals/nikkei225_fundamental_monthly.csv`

## 3) 中間產物（已整理）
- `fundamentals/intermediate/`
  - 歷史抓取中間檔、daily/月度拆分檔、舊版轉換檔

## 4) 台股 TWSE 抓取資料
- `twse_c02001_monthly/`：月度 C02001 ZIP、輸出 xlsx、下載日誌
- `twse_march_C02001/`：三月版資料（保留）
- `twse_reference/`：參考資料（保留）

## 5) 其他
- `DCA_Investment_Data.xlsx`：手動彙整檔（保留）
