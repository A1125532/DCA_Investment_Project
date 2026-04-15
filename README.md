# DCA Investment Project

本專題比較 3 種 DCA 延伸策略於 3 個市場之長期表現，並採用 `Expanding Yearly Walk-Forward`（2006 年起、以前一年資料校準、當年固定參數）進行回測。

## 作業繳交對照

### 1) 結構化資料儲存（CSV / Excel）與來源標註

- 清理後主資料（CSV，已完成日期對齊、欄位統一與缺值處理）
  - `data/taiex_price.csv`
  - `data/sp500_price.csv`
  - `data/nikkei225_price.csv`
  - `data/fundamentals/taiex_fundamental.csv`
  - `data/fundamentals/sp500_fundamental.csv`
  - `data/fundamentals/nikkei225_fundamental.csv`
- TAIEX 原始月報中介檔（Excel）
  - `data/twse_c02001_monthly/output/本益比_PE_大盤_每月_2006-2026.xlsx`
  - `data/twse_c02001_monthly/output/股價淨值比_PB_大盤_每月_2006-2026.xlsx`
  - `data/twse_c02001_monthly/output/殖利率_大盤_每月_2006-2026.xlsx`
- 資料來源說明
  - TAIEX：TWSE C02001
  - SP500：multpl.com
  - Nikkei225：Nikkei Indexes
  - 價格：Yahoo Finance

### 2) 程式碼繳交（Python 原始碼、結構清楚、含必要註解）

- 主要程式碼路徑
  - `scripts/scrapers/`：資料抓取
  - `scripts/analysis/`：策略與指標實作
  - `scripts/experiments/walkforward_yearly_calibration.py`：年度滾動校準
  - `scripts/visualization/plot_charts.py`：圖表與摘要輸出
  - `scripts/run_analysis.py`：整體流程入口

### 3) 資料檔繳交（各市場清理後資料；含整合檔）

- 每市場清理後基本面
  - `data/fundamentals/taiex_fundamental.csv`
  - `data/fundamentals/sp500_fundamental.csv`
  - `data/fundamentals/nikkei225_fundamental.csv`
- 每市場價格資料
  - `data/taiex_price.csv`
  - `data/sp500_price.csv`
  - `data/nikkei225_price.csv`
- 合併後總整理（跨市場、跨策略曲線）
  - `results/performance/walkforward_expanding_curves.csv`
  - `results/performance/walkforward_expanding_3strategies.csv`

### 4) 圖表繳交（至少三類圖表，標示清楚）

執行以下指令可重現圖表輸出：

```bash
python3 scripts/visualization/plot_charts.py
```

圖表輸出於 `figures/`，包含：

- 價格走勢圖
  - `figures/1_price_trend.png`
- 累積投入與投資組合價值比較圖
  - `figures/2_dca_TAIEX.png`
  - `figures/2_dca_SP500.png`
  - `figures/2_dca_Nikkei225.png`
- 最終績效比較圖
  - `figures/3_performance_comparison.png`

## 策略比較範圍

1. `Value Averaging`
2. `Model-Driven DCA`
3. `Threshold DCA (3factor)`

## 市場與資料來源說明

1. 價格資料（Yahoo Finance）
- `TAIEX`: `^TWII`
- `SP500`: `^GSPC`
- `Nikkei225`: `^N225`

2. 基本面資料
- `TAIEX`: TWSE C02001 月報（PE / PB / DividendYield）
- `SP500`: multpl.com（PE / PB / DividendYield）
- `Nikkei225`: Nikkei Indexes（PE / DividendYield，可含 PB）

## 重現流程（建議執行順序）

1. 安裝套件

```bash
pip install -r requirements.txt
```

2. 抓資料（價格 + 基本面）

```bash
python3 scripts/scrapers/run_all.py
```

3. 產生結果（先執行 yearly walk-forward，再輸出表格與圖）

```bash
python3 scripts/visualization/plot_charts.py
```

4. 一次跑完整流程（爬蟲 + 結果）

```bash
python3 scripts/run_analysis.py
```

## TAIEX 前置資料說明

`import_taiex_from_twse_xlsx.py` 會讀取下列三份檔案（預設 `data/twse_c02001_monthly/output/`）：

- `本益比_PE_大盤_每月_2006-2026.xlsx`
- `股價淨值比_PB_大盤_每月_2006-2026.xlsx`
- `殖利率_大盤_每月_2006-2026.xlsx`

若需重建上述三份 xlsx，請執行：

```bash
python3 scripts/scrapers/build_twse_monthly_valuation.py
```

## 核心流程說明

1. `scripts/experiments/walkforward_yearly_calibration.py`
- 每個測試年 `y`
- 校準窗使用 `2006-01-01 ~ (y-1)-12-31`（expanding window）
- 找到該年最佳參數後，固定套用到 `y` 年所有月份

2. `scripts/visualization/plot_charts.py`
- 呼叫 walk-forward 並輸出績效 CSV
- 產出價格圖、DCA 比較圖、績效比較圖

## 主要輸出

1. 績效與參數（`results/performance/`）
- `walkforward_expanding_3strategies.csv`
- `walkforward_expanding_param_log.csv`
- `walkforward_expanding_curves.csv`
- `performance_summary_3strategies_3factor.csv`
- `performance_summary_3strategies_3factor_with_sharpe_rank.csv`

2. 圖表（`figures/`）
- `1_price_trend.png`
- `2_dca_TAIEX.png`
- `2_dca_SP500.png`
- `2_dca_Nikkei225.png`
- `3_performance_comparison.png`

## 專案目錄摘要

```text
scripts/
  scrapers/
  analysis/
  experiments/
    walkforward_yearly_calibration.py
  visualization/
    plot_charts.py

results/performance/
figures/
data/
```

## 常用指令

```bash
# 只抓資料
python3 scripts/scrapers/run_all.py

# 只跑 expanding yearly walk-forward（輸出 CSV）
python3 scripts/experiments/walkforward_yearly_calibration.py

# 產生最新表格 + 圖（內含 walk-forward）
python3 scripts/visualization/plot_charts.py

# 全流程
python3 scripts/run_analysis.py
```
