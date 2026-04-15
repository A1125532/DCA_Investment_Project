# DCA Investment Project

本專案是一個以 **定期定額（Dollar-Cost Averaging, DCA）策略比較與回測** 為核心的研究型專案，針對三大市場（TAIEX、S&P 500、Nikkei 225）進行資料抓取、策略校準、績效評估與視覺化輸出。

---

## 使用程式語言

- **主要語言**：Python（核心策略、資料抓取、回測與視覺化皆以 Python 撰寫）
- **輔助腳本語言**：Bash（`run.sh` 啟動腳本）
- **執行環境常見指令介面**：PowerShell（Windows 使用者）

---

## 專案目標

- 建立可重現的 DCA 策略研究流程（資料抓取 -> 回測 -> 評估 -> 圖表）。
- 比較三種策略在不同市場的長期績效與風險表現。
- 使用估值因子（PE/PB/殖利率）驅動投入倍率，模擬實務中「便宜多買、昂貴少買」行為。
- 透過年度 Walk-Forward 校準，降低參數 overfitting 風險。

---

## 核心功能

- 自動抓取三大市場價格資料（Yahoo Finance）。
- 抓取/整理三市場基本面資料（台股 TWSE/Goodinfo、S&P500、Nikkei225）。
- 執行三種 DCA 策略：
  - Value Averaging（價值平均）
  - Model-Driven DCA（模型驅動）
  - Threshold DCA（門檻分位）
- 年度 expanding walk-forward 參數校準。
- 輸出：
  - 策略逐期明細（投入、持股、資產值、報酬）
  - 績效摘要（報酬、CAGR、Sharpe、Sortino、Max Drawdown）
  - Walk-forward 曲線與參數日誌
  - 比較圖表

---

## 技術棧與依賴

### 語言與主要函式庫

- Python 3.10+（建議）
- pandas、numpy：資料清理與計算
- matplotlib：視覺化
- yfinance：市場價格資料
- requests、beautifulsoup4、lxml：網頁資料擷取
- playwright、cloudscraper：動態頁/反爬備援
- openpyxl：Excel 讀寫（TWSE xlsx）

安裝依賴：

```bash
pip install -r requirements.txt
```

若要使用 Playwright 相關流程，請額外安裝瀏覽器：

```bash
python -m playwright install
```

---

## 專案結構

```text
DCA_Investment_Project/
├─ data/
│  ├─ fundamentals/                  # 各市場基本面資料
│  ├─ taiex_price.csv                # 台股指數價格
│  ├─ sp500_price.csv                # S&P 500 價格
│  ├─ nikkei225_price.csv            # 日經 225 價格
│  └─ ...                            # TWSE 中繼資料/處理結果
├─ scripts/
│  ├─ run_analysis.py                # 主流程入口（推薦）
│  ├─ run_all.py                     # 轉呼叫 scrapers 流程
│  ├─ analysis/
│  │  ├─ dca_strategies.py           # 策略邏輯
│  │  └─ performance_metrics.py      # 績效指標計算
│  ├─ experiments/
│  │  └─ walkforward_yearly_calibration.py  # 年度校準與回測核心
│  ├─ scrapers/                      # 各市場資料抓取與整理
│  └─ visualization/
│     └─ plot_charts.py              # 圖表與輸出整理
├─ results/
│  ├─ *.csv                          # 各策略輸出結果
│  └─ performance/                   # 績效摘要、曲線、參數紀錄
├─ requirements.txt
└─ run.sh
```

---

## 資料來源

- **價格資料**：Yahoo Finance（`^TWII`, `^GSPC`, `^N225`）
- **TAIEX 基本面**：
  - 優先：TWSE C02001 來源（xlsx 匯入/整理）
  - 備援：Goodinfo 抓取
- **S&P500 基本面**：multpl.com
- **Nikkei225 基本面**：Nikkei Indexes API（失敗時 archives 備援）

> 注意：外部網站版型/API 可能改版，抓取腳本需定期維護。

---

## 策略與評估方法

### 1) Value Averaging

透過每期「目標資產值」與「實際資產值」差距決定投入金額，並設定上下限避免過度投入。

### 2) Threshold DCA

根據估值指標（PE/PB/殖利率）分位數判斷低估/高估區間，動態調整每期投入倍率。

### 3) Model-Driven DCA

將多個估值因子轉換為標準化分數（z-score）後合成訊號，映射到投入倍率範圍。

### Walk-Forward 校準

- 以年度為單位做 expanding calibration。
- 每年用「起始年到前一年」資料尋找最佳參數，再套用到當年。
- 目標分數綜合 Sharpe、CAGR 與 Max Drawdown（具體權重見策略校準程式）。

### 主要績效指標

- Total Return
- CAGR
- Sharpe Ratio
- Sortino Ratio
- Max Drawdown

---

## 快速開始

### Windows (PowerShell)

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
python scripts/run_analysis.py
```

### macOS / Linux

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
bash run.sh
```

---

## 執行方式

### 一鍵完整流程（抓資料 + 回測 + 視覺化）

```bash
python scripts/run_analysis.py
```

### 僅資料抓取

```bash
python scripts/scrapers/run_all.py
```

### 單獨抓取價格資料

```bash
python scripts/scrapers/fetch_yahoo_prices.py
```

---

## 輸出結果說明

### `results/`

- 各市場 x 各策略明細 CSV（例如 `TAIEX_value_avg.csv`）
- 可用於自訂分析、再繪圖或報告附錄

### `results/performance/`

- `performance_summary_*.csv`：策略績效總表
- `walkforward_expanding_curves.csv`：策略淨值/績效曲線
- `walkforward_expanding_param_log.csv`：各年度校準參數紀錄

---

## 典型資料流（端到端）

1. `scripts/run_analysis.py` 啟動流程  
2. `scripts/scrapers/run_all.py` 依序抓取價格與基本面  
3. `scripts/experiments/walkforward_yearly_calibration.py` 進行年度校準與回測  
4. `scripts/visualization/plot_charts.py` 匯整結果並輸出 CSV/圖表  
5. 最終產出寫入 `results/` 與 `results/performance/`

---

## 已知限制

- 尚未建立正式測試框架（`pytest`/CI）。
- 多數參數以程式內常數設定，彈性不足。
- 部分資料源依賴網頁爬蟲，對來源站變動敏感。
- 目前以腳本研究用途為主，尚未 package 化。

---

## 建議改進方向

- 加入 `pytest` 測試（策略輸出 schema、績效計算正確性）。
- 將參數抽離為 `config`（YAML/JSON）或 CLI 參數。
- 增加資料品質檢查（缺值率、欄位完整性、日期連續性）。
- 結果輸出加上 timestamp 版本化，避免覆寫歷史結果。
- 建立 CI（lint + test）提升可維護性。

---

## 免責與使用聲明

- 本專案為研究與教學用途，不構成投資建議。
- 回測結果不代表未來績效；實際交易仍需考慮滑價、交易成本、稅費與流動性。

