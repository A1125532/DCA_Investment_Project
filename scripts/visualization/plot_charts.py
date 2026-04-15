#!/usr/bin/env python3
"""
===================================================================
視覺化圖表生成（visualization/plot_charts.py）
===================================================================
功能：生成所有 DCA 分析圖表
===================================================================

圖表對照：
┌──────┬──────────────────────────────────────────┬──────────────┐
│ 圖號   │ 內容                                    │ 輸出檔案      │
├──────┼──────────────────────────────────────────┼──────────────┤
│ 圖1   │ 三大市場價格趨勢                        │ 1_price_trend.png │
│ 圖2   │ 各市場 DCA 疊圖（3策略同圖比較）        │ 2_dca_*.png     │
│ 圖3   │ 績效比較柱狀圖（9策略-市場組合）         │ 3_performance... │
│ 圖4   │ Model DCA 分析：股價 + P/E + 組合價值   │ 4_model_dca_*.png │
│ 圖5   │ Threshold DCA：組合價值 + 月投入信號     │ 5_threshold_*.png │
└──────┴──────────────────────────────────────────┴──────────────┘

重大金融事件標記（所有圖表）：
  💥 Lehman Brothers 破產    → 2008/09/15
  🦠 COVID-19               → 2020/03/11
  📈 Fed 升息               → 2022/03/01
===================================================================
"""

import os
import sys
import shutil
from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib
matplotlib.use('Agg')  # 不開視窗，直接存檔
plt.rcParams['font.family'] = ['DejaVu Sans', 'sans-serif']
plt.rcParams['axes.unicode_minus'] = False

# 路徑設定（repo root）
PROJECT_DIR = Path(__file__).resolve().parents[2]
DATA_DIR = PROJECT_DIR / 'data'
FUND_DIR = DATA_DIR / 'fundamentals'
OUT_DIR = PROJECT_DIR / 'figures'
OUT_DIR.mkdir(parents=True, exist_ok=True)
RESULTS_DIR = PROJECT_DIR / 'results'
PERF_DIR = RESULTS_DIR / 'performance'
RESULTS_DIR.mkdir(parents=True, exist_ok=True)
PERF_DIR.mkdir(parents=True, exist_ok=True)

# 研究期間
DCA_START, DCA_END, MONTHLY = '2006-01-01', '2026-03-01', 1000
YEARS = 20.17
RF_RATE = 0.02

# 資料來源對照（寫入結果表格）
PRICE_SOURCE_MAP = {
    'TAIEX': 'Yahoo Finance (^TWII) -> data/taiex_price.csv',
    'SP500': 'Yahoo Finance (^GSPC) -> data/sp500_price.csv',
    'Nikkei225': 'Yahoo Finance (^N225) -> data/nikkei225_price.csv',
}

FUND_SOURCE_MAP = {
    'TAIEX': 'Goodinfo/TWSE -> data/fundamentals/taiex_fundamental.csv',
    'SP500': 'multpl.com -> data/fundamentals/sp500_fundamental.csv',
    'Nikkei225': 'Nikkei Indexes -> data/fundamentals/nikkei225_fundamental.csv',
}

STRATEGY_SOURCE_MAP = {
    '① Value Averaging': 'scripts/experiments/walkforward_yearly_calibration.py::run_market(Value Averaging)',
    '② Model-Driven DCA': 'scripts/experiments/walkforward_yearly_calibration.py::run_market(Model-Driven DCA)',
    '③ Threshold DCA': 'scripts/experiments/walkforward_yearly_calibration.py::run_market(Threshold DCA 3factor)',
}

# 重大金融事件（所有圖表都會標記）
CRISIS_EVENTS = [
    {'date': '2008-09-15', 'label': 'Lehman', 'color': '#ef4444'},
    {'date': '2020-03-11', 'label': 'COVID-19', 'color': '#f97316'},
    {'date': '2022-03-01', 'label': 'Fed Hike',   'color': '#8b5cf6'},
]


# ============================================================
# 載入資料
# ============================================================

def _clean_price_df(df):
    """清理價格表頭/日期欄位，回傳 Series（index=日期, value=Close）"""
    if isinstance(df.columns, pd.MultiIndex):
        df.columns = [col[1] if isinstance(col, tuple) else col for col in df.columns]

    if 'Date' in df.columns and 'Price' not in df.columns:
        df = df.rename(columns={'Date': 'Price'})
    if df.columns[0] != 'Price':
        # 支援奇怪的欄位順序（保留第一欄為日期欄位）
        df = df.rename(columns={df.columns[0]: 'Price'})

    # 去除 yfinance 輸出的非日期表頭行（如 Ticker/空白行）
    df['Price'] = pd.to_datetime(df['Price'], errors='coerce')
    df = df[df['Price'].notna()]

    close_col = 'Close' if 'Close' in df.columns else None
    if close_col is None:
        lower_cols = {c.lower(): c for c in df.columns}
        close_col = lower_cols.get('close', None)
    if close_col is None and 'Adj Close' in df.columns:
        close_col = 'Adj Close'

    if close_col is None:
        raise ValueError('找不到價格欄位（Close/Adj Close）')

    df = df[['Price', close_col]].rename(columns={close_col: 'Close'})
    return df.set_index('Price')['Close'].sort_index()


def load_prices():
    """載入三大市場月收盤價"""
    prices = {}
    for fname, name in [('taiex_price.csv', 'TAIEX'), ('sp500_price.csv', 'SP500'),
                        ('nikkei225_price.csv', 'Nikkei225')]:
        path = DATA_DIR / fname
        df = pd.read_csv(path)
        if df.empty:
            raise ValueError(f"價格資料為空：{path}")
        prices[name] = _clean_price_df(df)
    return prices


def load_fundamentals():
    """載入三大市場基本面資料"""
    rename_map = {
        'trailingPE': 'PE',
        'per': 'PE',
        'PER': 'PE',
        'priceToBook': 'PB',
        'pb': 'PB',
        'dividendYield': 'DividendYield',
        'dividendyield': 'DividendYield',
    }

    def _load_one(path):
        df = pd.read_csv(path)
        if df.empty:
            raise ValueError(f"基本面資料為空：{path}")
        if 'Unnamed: 0' in df.columns:
            df = df.rename(columns={'Unnamed: 0': 'date'})

        # 將第一欄改為 date 欄位（舊版輸出保險）
        if 'date' not in df.columns and len(df.columns) > 0:
            df = df.rename(columns={df.columns[0]: 'date'})

        df = df.rename(columns=rename_map)
        if 'date' not in df.columns:
            raise ValueError(f"缺少 date 欄位：{path}")
        df['date'] = pd.to_datetime(df['date'], errors='coerce')
        df = df[df['date'].notna()].set_index('date').sort_index()
        return df

    return {
        'TAIEX': _load_one(FUND_DIR / 'taiex_fundamental.csv'),
        'SP500': _load_one(FUND_DIR / 'sp500_fundamental.csv'),
        'Nikkei225': _load_one(FUND_DIR / 'nikkei225_fundamental.csv'),
    }


def _normalize_result_df(df):
    """統一策略輸出欄位，方便保存與績效計算。"""
    if df is None or len(df) == 0:
        return df
    out = df.copy()
    rename_map = {}
    if 'portfolio_value' in out.columns and 'value' not in out.columns:
        rename_map['portfolio_value'] = 'value'
    if 'total_invested' in out.columns and 'invested' not in out.columns:
        rename_map['total_invested'] = 'invested'
    if 'investment' in out.columns and 'monthly_invest' not in out.columns:
        rename_map['investment'] = 'monthly_invest'
    if rename_map:
        out = out.rename(columns=rename_map)
    return out


def _build_dca_results_yearly_walkforward(curves_path):
    """
    從 yearly walk-forward 曲線 CSV 建立繪圖用結果。
    """
    if not curves_path.exists():
        raise FileNotFoundError(f"找不到 yearly walk-forward 曲線檔：{curves_path}")

    wf_curve = pd.read_csv(curves_path)
    wf_curve['Date'] = pd.to_datetime(wf_curve['Date'], errors='coerce')
    wf_curve = wf_curve.dropna(subset=['Date'])

    wf_to_plot = {
        'Value Averaging': '① Value Averaging',
        'Model-Driven DCA': '② Model-Driven DCA',
        'Threshold DCA (3factor)': '③ Threshold DCA',
    }

    dca_results = {}
    for market in ['TAIEX', 'SP500', 'Nikkei225']:
        curve_df = wf_curve[wf_curve['Market'] == market].copy()
        market_res = {}
        for wf_name, plot_name in wf_to_plot.items():
            s = curve_df[curve_df['Strategy'] == wf_name].copy()
            if len(s) == 0:
                continue
            s = s.rename(columns={
                'Date': 'date',
                'Price': 'price',
                'Shares': 'shares',
                'Invested': 'invested',
                'Value': 'value',
                'Monthly Invest': 'monthly_invest',
            })
            s['date'] = pd.to_datetime(s['date'], errors='coerce')
            s = s.dropna(subset=['date']).sort_values('date')
            s['return_pct'] = np.where(
                s['invested'] > 0,
                (s['value'] - s['invested']) / s['invested'] * 100.0,
                0.0,
            )
            market_res[plot_name] = s[['date', 'price', 'shares', 'invested', 'value', 'monthly_invest', 'return_pct']]
        dca_results[market] = market_res
    return dca_results


def _calc_metrics(df, years=YEARS, rf_rate=RF_RATE):
    """計算 Total Return / CAGR / Sharpe / Sortino / MaxDD。"""
    df = _normalize_result_df(df)
    if df is None or len(df) == 0 or 'value' not in df.columns or 'invested' not in df.columns:
        return None

    work = df.copy()
    work['value'] = pd.to_numeric(work['value'], errors='coerce')
    work['invested'] = pd.to_numeric(work['invested'], errors='coerce')
    work = work.dropna(subset=['value', 'invested'])
    if len(work) < 2:
        return None

    final_value = float(work['value'].iloc[-1])
    invested = float(work['invested'].iloc[-1])
    if invested <= 0:
        return None

    total_return = (final_value - invested) / invested * 100.0
    cagr = (((final_value / invested) ** (1.0 / years)) - 1.0) * 100.0 if years > 0 else np.nan

    monthly_returns = work['value'].pct_change().dropna()
    monthly_rf = (1.0 + rf_rate) ** (1.0 / 12.0) - 1.0
    excess = monthly_returns - monthly_rf

    if len(monthly_returns) > 1 and float(monthly_returns.std(ddof=1)) > 0:
        sharpe = float(excess.mean() / monthly_returns.std(ddof=1) * np.sqrt(12))
    else:
        sharpe = np.nan

    downside = excess[excess < 0]
    if len(downside) > 1 and float(downside.std(ddof=1)) > 0:
        sortino = float(excess.mean() / downside.std(ddof=1) * np.sqrt(12))
    else:
        sortino = np.nan

    values = work['value'].to_numpy(dtype=float)
    running_max = np.maximum.accumulate(values)
    max_dd = float(np.max((running_max - values) / np.maximum(running_max, 1e-9)) * 100.0)

    return {
        'Total Return (%)': total_return,
        'CAGR (%)': cagr,
        'Sharpe': sharpe,
        'Sortino': sortino,
        'MaxDD (%)': max_dd,
        'Invested': invested,
        'Final Value': final_value,
    }


def _calc_monthly_irr(cashflows, tol=1e-8, max_iter=200):
    """固定月頻現金流 IRR（回傳月化）。"""
    if not cashflows or not any(cf < 0 for cf in cashflows) or not any(cf > 0 for cf in cashflows):
        return np.nan

    def npv(rate):
        base = 1.0 + rate
        if base <= 0:
            return np.nan
        ln_base = np.log(base)
        total = 0.0
        for t, cf in enumerate(cashflows):
            if t == 0:
                total += cf
                continue
            ln_denom = t * ln_base
            # 避免 exp 下溢為 0 或上溢為 inf 造成除零錯誤
            if ln_denom < -700:
                return np.sign(cf) * np.inf if cf != 0 else total
            if ln_denom > 700:
                continue
            denom = np.exp(ln_denom)
            total += cf / denom
        return total

    def sign_of(x):
        if pd.isna(x):
            return 0
        if x > 0:
            return 1
        if x < 0:
            return -1
        return 0

    # 避免貼近 -1 導致長期折現分母下溢
    low, high = -0.95, 1.0
    npv_low, npv_high = npv(low), npv(high)
    s_low, s_high = sign_of(npv_low), sign_of(npv_high)
    grow = 0
    while s_low * s_high > 0 and grow < 30:
        high *= 2.0
        npv_high = npv(high)
        s_high = sign_of(npv_high)
        grow += 1
    if s_low == 0 or s_high == 0 or s_low * s_high > 0:
        return np.nan

    mid = np.nan
    for _ in range(max_iter):
        mid = (low + high) / 2.0
        npv_mid = npv(mid)
        if not np.isfinite(npv_mid):
            high = mid
            continue
        if abs(npv_mid) < tol:
            return mid
        if sign_of(npv_low) * sign_of(npv_mid) <= 0:
            high = mid
            npv_high = npv_mid
        else:
            low = mid
            npv_low = npv_mid
    return mid


def _calc_irr_annual(df):
    """由每月投入與期末淨值計算年化 IRR。"""
    df = _normalize_result_df(df)
    if df is None or len(df) < 2 or 'value' not in df.columns:
        return np.nan

    work = df.copy()
    if 'monthly_invest' in work.columns:
        monthly = pd.to_numeric(work['monthly_invest'], errors='coerce').fillna(0.0)
    elif 'investment' in work.columns:
        monthly = pd.to_numeric(work['investment'], errors='coerce').fillna(0.0)
    elif 'invested' in work.columns:
        invested = pd.to_numeric(work['invested'], errors='coerce').fillna(0.0)
        monthly = invested.diff().fillna(invested.iloc[0])
    else:
        return np.nan

    final_value = float(pd.to_numeric(work['value'], errors='coerce').iloc[-1])
    cashflows = [-float(x) for x in monthly.to_list()]
    if len(cashflows) == 0:
        return np.nan
    cashflows[-1] += final_value

    r_m = _calc_monthly_irr(cashflows)
    if pd.isna(r_m):
        return np.nan
    return float((1.0 + r_m) ** 12.0 - 1.0)


def export_results(dca_results, threshold_mode):
    """輸出各市場策略 CSV 與績效摘要 CSV。"""
    suffix_map = {
        '① Value Averaging': 'value_avg',
        '② Model-Driven DCA': 'modeldriven',
        '③ Threshold DCA': 'threshold',
    }

    summary_rows = []
    for market, res in dca_results.items():
        for strategy, df in res.items():
            if df is None or len(df) == 0:
                continue

            out_df = _normalize_result_df(df)
            price_source = PRICE_SOURCE_MAP.get(market, '')
            if strategy == '① Value Averaging':
                fund_source = 'N/A（僅使用股價）'
            else:
                fund_source = FUND_SOURCE_MAP.get(market, '')
            strategy_source = STRATEGY_SOURCE_MAP.get(strategy, 'scripts/analysis/dca_strategies.py')

            out_df['資料來源_股價'] = price_source
            out_df['資料來源_基本面'] = fund_source
            out_df['資料來源_策略'] = strategy_source

            suffix = suffix_map.get(strategy, strategy.replace(' ', '_').lower())
            out_path = RESULTS_DIR / f'{market}_{suffix}.csv'
            out_df.to_csv(out_path, index=False, encoding='utf-8-sig')

            metrics = _calc_metrics(out_df)
            if metrics is None:
                continue
            irr_ann = _calc_irr_annual(out_df)
            strategy_name = strategy
            if strategy == '③ Threshold DCA':
                strategy_name = f'③ Threshold DCA ({threshold_mode})'

            summary_rows.append({
                '市場': market,
                '策略': strategy_name,
                '總投入': round(metrics['Invested'], 2),
                '最終價值': round(metrics['Final Value'], 2),
                '總報酬率%': round(metrics['Total Return (%)'], 3),
                'CAGR%': round(metrics['CAGR (%)'], 3),
                'Sharpe': round(metrics['Sharpe'], 3) if pd.notna(metrics['Sharpe']) else np.nan,
                'Sortino': round(metrics['Sortino'], 3) if pd.notna(metrics['Sortino']) else np.nan,
                'MaxDD%': round(metrics['MaxDD (%)'], 3),
                'IRR%': round(irr_ann * 100.0, 3) if pd.notna(irr_ann) else np.nan,
                '資料來源_股價': price_source,
                '資料來源_基本面': fund_source,
                '資料來源_策略': strategy_source,
            })

    if summary_rows:
        summary_df = pd.DataFrame(summary_rows)
        n_strategies = summary_df['策略'].nunique()
        summary_path = PERF_DIR / f'performance_summary_{n_strategies}strategies_{threshold_mode}.csv'
        summary_df.to_csv(summary_path, index=False, encoding='utf-8-sig')
        print(f"  ✅ 績效摘要：{summary_path.name}")


# ============================================================
# 圖1：三大市場價格趨勢圖
# ============================================================

def plot_price_trend(prices):
    """
    圖1：三大市場價格趨勢圖
    
    內容：
    - TAIEX、S&P 500、Nikkei 225 三條月線
    - 加上金融危機標記
    - X軸時間軸：2006-2026
    
    用途：
    - 呈現三個市場 20 年的價格變化
    - 顯示金融危機期間的下跌幅度
    """
    fig, ax = plt.subplots(figsize=(14, 6))
    xmin, xmax = pd.Timestamp(DCA_START), pd.Timestamp(DCA_END)
    
    for name, series in prices.items():
        s = series[(series.index >= xmin) & (series.index <= xmax)]
        ax.plot(s.index, s.values, label=name, linewidth=1.5)
    
    ax.set_title('Figure 1: Price Trend (2006-2026)', fontsize=14, fontweight='bold')
    ax.set_xlabel('Date')
    ax.set_ylabel('Price')
    ax.legend()
    ax.grid(True, alpha=0.3)
    ax.set_xlim(xmin, xmax)
    
    # 加入金融危機標記
    _add_markers(ax, xmin, xmax, prices)
    
    plt.tight_layout()
    plt.savefig(OUT_DIR / '1_price_trend.png', dpi=150)
    plt.close()
    print("  ✅ 圖1：1_price_trend.png")


# ============================================================
# 圖2：每市場一張（3策略同圖）
# ============================================================

def plot_dca_overlay(dca_results, market='TAIEX'):
    """One figure per market, two panels:
    left = portfolio value lines, right = return% lines.
    Both panels overlay the 3 strategies.
    """
    colors = {'① Value Averaging': '#1f77b4',
              '② Model-Driven DCA': '#ff7f0e',
              '③ Threshold DCA': '#2ca02c'}

    xmin_d = pd.Timestamp(DCA_START)
    xmax_d = pd.Timestamp(DCA_END)
    results = dca_results.get(market, {})
    if not results:
        print(f"  ⚠️ 圖2略過：{market} 無策略結果")
        return

    fig, axes = plt.subplots(1, 2, figsize=(15, 5), sharex=True)
    ax_val, ax_ret = axes
    for strat, df in results.items():
        col = colors.get(strat, '#333')
        ax_val.plot(df['date'], df['value'], color=col, linewidth=2.0, label=strat)
        ax_ret.plot(df['date'], df['return_pct'], color=col, linewidth=2.0, label=strat)

    ax_val.set_title(f'{market} - Portfolio Value', fontsize=12, fontweight='bold')
    ax_val.set_ylabel('Value')
    ax_val.set_xlabel('Date')
    ax_val.grid(True, alpha=0.3)
    ax_val.set_xlim(xmin_d, xmax_d)
    ax_val.legend(fontsize=9, loc='upper left')
    _add_markers(ax_val, xmin_d, xmax_d, None)

    ax_ret.set_title(f'{market} - Cumulative Return (%)', fontsize=12, fontweight='bold')
    ax_ret.set_ylabel('Return (%)')
    ax_ret.set_xlabel('Date')
    ax_ret.axhline(0, color='black', linewidth=0.8, alpha=0.7)
    ax_ret.grid(True, alpha=0.3)
    ax_ret.set_xlim(xmin_d, xmax_d)
    ax_ret.legend(fontsize=9, loc='upper left')
    _add_markers(ax_ret, xmin_d, xmax_d, None)

    plt.tight_layout()
    safe = market.replace(' ', '_').replace('/', '_')
    plt.savefig(OUT_DIR / f'2_dca_{safe}.png', dpi=150)
    plt.close()
    print(f"  ✅ 圖2：2_dca_{safe}.png")


# ============================================================
# 圖3：績效比較柱狀圖
# ============================================================

def plot_performance_comparison(dca_results):
    """
    圖3：績效比較柱狀圖
    
    內容：
    - X軸：三種策略
    - Y軸：總報酬率（%）
    - 三組柱子：TAIEX、S&P 500、Nikkei 225
    - 每個市場用不同顏色
    
    用途：
    - 一眼看出哪個市場哪個策略報酬最高
    - 橫向比較三個市場的差異
    """
    summary = []
    for market, results in dca_results.items():
        for strat, df in results.items():
            if df is not None and len(df) > 0:
                final = df.iloc[-1]
                summary.append({
                    'Market': market,
                    'Strategy': strat,
                    'Total Return (%)': float(final['return_pct']),
                })
    
    if not summary: return
    df_sum = pd.DataFrame(summary)
    
    fig, ax = plt.subplots(figsize=(14, 6))
    markets = df_sum['Market'].unique()
    strategies = df_sum['Strategy'].unique()
    x = np.arange(len(strategies))
    width = 0.8 / len(markets)
    colors_mkt = ['#1f77b4', '#ff7f0e', '#2ca02c']
    
    for i, market in enumerate(markets):
        subset = df_sum[df_sum['Market'] == market].set_index('Strategy').reindex(strategies)
        offset = (i - len(markets)/2 + 0.5) * width
        ax.bar(x + offset, subset['Total Return (%)'], width,
               label=market, color=colors_mkt[i % len(colors_mkt)],
               edgecolor='white', linewidth=0.5)
    
    ax.set_title('Figure 3: Total Return (%) Comparison', fontsize=14, fontweight='bold')
    ax.set_xticks(x)
    ax.set_xticklabels([s.replace('DCA', '\nDCA').replace('Value Averaging', 'Value\nAveraging')
                       for s in strategies], fontsize=10)
    ax.set_ylabel('Total Return (%)')
    ax.legend()
    ax.grid(True, alpha=0.3, axis='y')
    plt.tight_layout()
    plt.savefig(OUT_DIR / '3_performance_comparison.png', dpi=150)
    plt.close()
    print("  ✅ 圖3：3_performance_comparison.png")


# ============================================================
# 圖4：Model-Driven DCA 分析（股價 + P/E + 組合價值）
# ============================================================

def plot_model_dca(dca_results, prices, fundamentals):
    """
    圖4：Model-Driven DCA 詳細分析
    
    內容（三個市場，各三層）：
    - 上層：股價（藍線）
    - 中層：P/E（橙線）+ 歷史中位數參考帶（虛線）
    - 下層：組合價值 vs 累計投入
    
    用途：
    - 呈現 Model 策略在估值區間的行為
    - 中位數參考帶僅供視覺輔助，實際投入由 walk-forward 校準決定
    """
    for market in ['SP500', 'TAIEX', 'Nikkei225']:
        dyn_df = dca_results.get(market, {}).get('② Model-Driven DCA')
        fund_df = fundamentals.get(market)
        price_s = prices.get(market)
        
        if dyn_df is None or fund_df is None: continue
        
        fig, axes = plt.subplots(3, 1, figsize=(12, 10))
        xmin_d, xmax_d = dyn_df['date'].min(), dyn_df['date'].max()
        
        # 上：股價
        if price_s is not None:
            p_s = price_s[(price_s.index >= xmin_d) & (price_s.index <= xmax_d)]
            axes[0].plot(p_s.index, p_s.values, 'b-', linewidth=1, label='Price', alpha=0.8)
            axes[0].set_ylabel('Price')
            axes[0].set_title(f'{market} - Model-Driven DCA Analysis', fontsize=12, fontweight='bold')
            axes[0].legend()
            axes[0].grid(True, alpha=0.3)
            axes[0].set_xlim(xmin_d, xmax_d)
            _add_markers(axes[0], xmin_d, xmax_d, p_s)
        
        # 中：P/E
        pe_col = 'PE' if 'PE' in fund_df.columns else 'PER'
        pe_vals = []
        for d in dyn_df['date']:
            avail = fund_df[fund_df.index <= d]
            if len(avail) > 0:
                pe_vals.append((d, avail[pe_col].iloc[-1]))
            else:
                pe_vals.append((d, np.nan))
        pe_df = pd.DataFrame(pe_vals, columns=['date', 'PE']).set_index('date')
        axes[1].plot(pe_df.index, pe_df['PE'], 'orange', linewidth=1.2, label='P/E', alpha=0.8)
        med_pe = pe_df['PE'].median()
        if not np.isnan(med_pe) and med_pe > 0:
            axes[1].axhline(med_pe * 0.80, color='green', linestyle=':', linewidth=1,
                           label=f'Undervalued ({med_pe*0.80:.1f})')
            axes[1].axhline(med_pe * 1.20, color='red', linestyle=':', linewidth=1,
                           label=f'Overvalued ({med_pe*1.20:.1f})')
        axes[1].set_ylabel('P/E')
        axes[1].legend(fontsize=8)
        axes[1].grid(True, alpha=0.3)
        axes[1].set_xlim(xmin_d, xmax_d)
        _add_markers(axes[1], xmin_d, xmax_d, pe_df)
        
        # 下：組合價值
        axes[2].plot(dyn_df['date'], dyn_df['value'], 'b-', linewidth=1.5, label='Portfolio Value')
        axes[2].plot(dyn_df['date'], dyn_df['invested'], 'gray', linewidth=1,
                     linestyle='--', label='Cumulative Invested')
        axes[2].fill_between(dyn_df['date'], dyn_df['invested'], dyn_df['value'],
                              alpha=0.1, color='blue')
        axes[2].set_ylabel('Value ($)')
        axes[2].set_xlabel('Date')
        axes[2].legend()
        axes[2].grid(True, alpha=0.3)
        axes[2].set_xlim(xmin_d, xmax_d)
        _add_markers(axes[2], xmin_d, xmax_d, dyn_df)
        
        plt.tight_layout()
        safe = market.replace(' ', '_').replace('/', '_')
        plt.savefig(OUT_DIR / f'4_model_dca_{safe}.png', dpi=150)
        plt.close()
        print(f"  ✅ 圖4：4_model_dca_{safe}.png")


# ============================================================
# 圖5：Threshold DCA 分析（組合價值 + 月投入信號）
# ============================================================

def plot_threshold_dca(dca_results):
    """
    圖5：Threshold DCA 詳細分析
    
    內容（三個市場，各兩層）：
    - 上層：組合價值 vs 累計投入
    - 下層：月投入柱狀圖（顏色區分信號）
      - 綠色：多買（滿足低估條件）
      - 紅色：暫停（滿足高估條件）
      - 藍色：正常
    
    用途：
    - 呈現多指標門檻觸發的時機
    - 綠色柱子多的區間代表低估（買到便宜）
    - 紅色柱子代表高估暂停（避開高點）
    """
    for market in ['SP500', 'TAIEX', 'Nikkei225']:
        thresh_df = dca_results.get(market, {}).get('③ Threshold DCA')
        if thresh_df is None: continue
        
        fig, axes = plt.subplots(2, 1, figsize=(12, 7))
        xmin_d, xmax_d = thresh_df['date'].min(), thresh_df['date'].max()
        
        # 上：組合價值
        axes[0].plot(thresh_df['date'], thresh_df['value'], 'b-', linewidth=1.5, label='Portfolio Value')
        axes[0].plot(thresh_df['date'], thresh_df['invested'], 'gray', linewidth=1,
                     linestyle='--', label='Cumulative Invested')
        axes[0].fill_between(thresh_df['date'], thresh_df['invested'], thresh_df['value'],
                             alpha=0.1, color='blue')
        axes[0].set_title(f'{market} - Threshold DCA (3 factors)', fontsize=12, fontweight='bold')
        axes[0].set_ylabel('Value ($)')
        axes[0].legend()
        axes[0].grid(True, alpha=0.3)
        axes[0].set_xlim(xmin_d, xmax_d)
        _add_markers(axes[0], xmin_d, xmax_d, thresh_df)
        
        # 下：月投入柱狀圖
        bar_colors = []
        for inv in thresh_df['monthly_invest']:
            if inv > MONTHLY * 1.1:
                bar_colors.append('green')   # 多買
            elif inv == 0:
                bar_colors.append('red')     # 暂停
            else:
                bar_colors.append('steelblue')  # 正常
        axes[1].bar(thresh_df['date'], thresh_df['monthly_invest'],
                     color=bar_colors, alpha=0.7, width=20, label='Monthly Invest')
        axes[1].axhline(MONTHLY, color='orange', linewidth=1.5,
                        label=f'Base ${MONTHLY}')
        axes[1].set_title(f'{market} - Threshold DCA Monthly Signal', fontsize=12, fontweight='bold')
        axes[1].set_ylabel('Monthly Invest ($)')
        axes[1].set_xlabel('Date')
        axes[1].legend()
        axes[1].grid(True, alpha=0.3, axis='y')
        axes[1].set_xlim(xmin_d, xmax_d)
        _add_markers(axes[1], xmin_d, xmax_d, thresh_df)
        
        plt.tight_layout()
        safe = market.replace(' ', '_').replace('/', '_')
        plt.savefig(OUT_DIR / f'5_threshold_{safe}.png', dpi=150)
        plt.close()
        print(f"  ✅ 圖5：5_threshold_{safe}.png")


# ============================================================
# 輔助函式
# ============================================================

def _add_markers(ax, xmin, xmax, data=None):
    """在圖上加入金融危機標記"""
    for ev in CRISIS_EVENTS:
        d = pd.Timestamp(ev['date'])
        if xmin <= d <= xmax:
            ax.axvline(d, color=ev['color'], linestyle='--', linewidth=1.2, alpha=0.6)
            y_max = ax.get_ylim()[1]
            ax.annotate(ev['label'],
                       xy=(d, y_max),
                       xytext=(d, y_max * 0.95),
                       fontsize=7, ha='center', color='white', fontweight='bold',
                       bbox=dict(boxstyle='round,pad=0.3',
                                facecolor=ev['color'], alpha=0.85))


# ============================================================
# 主程式
# ============================================================

if __name__ == '__main__':
    # 讓 `from experiments...` 能從 scripts/experiments 匯入
    scripts_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), '..'))
    if scripts_dir not in sys.path:
        sys.path.insert(0, scripts_dir)
    
    print("=" * 60)
    print("視覺化圖表生成")
    print("=" * 60)
    
    # 載入資料
    print("\n[1] 載入資料...")
    prices = load_prices()
    fundamentals = load_fundamentals()
    threshold_mode = '3factor'
    print("  模式：Yearly Walk-Forward（2006~前一年校準，當年固定參數）")
    print(f"  Threshold 模式：{threshold_mode}")
    
    # 執行 yearly walk-forward（產生參數/曲線 CSV）
    print("\n[2] 執行 DCA 策略...")
    from experiments.walkforward_yearly_calibration import main as wf_main
    wf_main()

    curves_path = PERF_DIR / 'walkforward_expanding_curves.csv'
    dca_results = _build_dca_results_yearly_walkforward(curves_path)
    for market in ['TAIEX', 'SP500', 'Nikkei225']:
        res = dca_results.get(market, {})
        for strategy in ['① Value Averaging', '② Model-Driven DCA', '③ Threshold DCA']:
            df = res.get(strategy)
            if df is None or len(df) == 0:
                print(f"  {strategy} {market} ❌: 無資料")
            else:
                print(f"  {strategy} {market}: {df.iloc[-1]['return_pct']:+.1f}%")

    # 匯出策略結果與績效摘要（覆寫舊檔，避免讀到過期結果）
    print("\n[2.5] 匯出結果...")
    export_results(dca_results, threshold_mode)
    
    # 生成圖表（每市場一張；每張左值右報酬，三策略同圖）
    print("\n[3] 生成圖表...")
    plot_price_trend(prices)

    for market in ['TAIEX', 'SP500', 'Nikkei225']:
        plot_dca_overlay(dca_results, market=market)

    plot_performance_comparison(dca_results)
    
    # 複製到伺服器資料夾
    server_dir = PROJECT_DIR / 'figures'
    server_dir.mkdir(parents=True, exist_ok=True)
    for f in os.listdir(OUT_DIR):
        if f.endswith('.png'):
            src = OUT_DIR / f
            dst = server_dir / f
            if src.resolve() != dst.resolve():
                shutil.copy2(src, dst)
    
    print(f"\n✅ 完成！{len(os.listdir(OUT_DIR))} 張圖表 → {OUT_DIR}")
    print("=" * 60)
