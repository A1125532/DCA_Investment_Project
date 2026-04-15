#!/usr/bin/env python3
"""
===================================================================
績效指標計算（analysis/performance_metrics.py）
===================================================================
功能：自寫函數計算績效指標（不使用 library）
===================================================================

指標說明：
┌──────────┬──────────────────────────────────┬──────────────────┐
│ 指標       │ 公式                              │ 意義              │
├──────────┼──────────────────────────────────┼──────────────────┤
│ CAGR     │ (Final/Invested)^(1/y)-1         │ 年均複合報酬率     │
│ Sharpe   │ (月均報酬-月無風險)/月標準差×√12 │ 風險調整報酬      │
│ Sortino  │ (月均報酬-月無風險)/下行標準差×√12│ 只看下跌風險     │
│ MaxDD    │ max((Peak-Trough)/Peak)            │ 最大回撤          │
└──────────┴──────────────────────────────────┴──────────────────┘
===================================================================
"""

import pandas as pd
import numpy as np

RF_RATE = 0.02  # 無風險利率：年化 2%（可調整）


def calc_cagr(total_invested, final_value, years):
    """
    CAGR（年均化報酬率）
    
    公式：CAGR = (Final Value / Total Invested)^(1/years) - 1
    
    意義：衡量每年平均複合成長率，適用於不同期間的投資比較
    """
    if total_invested <= 0 or final_value <= 0:
        return None
    if years <= 0:
        return None
    return (final_value / total_invested) ** (1 / years) - 1


def calc_sharpe(monthly_returns, rf_rate=RF_RATE):
    """
    Sharpe Ratio（夏普比例）
    
    公式：Sharpe = (月均報酬 - 月無風險利率) / 月標準差 × √12
    
    意義：每承擔一單位風險所獲得的超額報酬
    - Sharpe > 1：風險調整後報酬佳
    - Sharpe < 0：不值得投資（連無風險利率都跑輸）
    
    設定：
    - 無風險利率：年化 2%（參考各研究文獻標準）
    - 月無風險利率 = (1+年)^(1/12) - 1
    """
    if monthly_returns is None or len(monthly_returns) < 2:
        return None
    
    # 月無風險利率（年化 → 月化）
    monthly_rf = (1 + rf_rate) ** (1/12) - 1
    
    # 超額報酬（月報酬 - 月無風險利率）
    excess_returns = monthly_returns - monthly_rf
    
    # 月標準差（自由度 1）
    std = np.std(excess_returns, ddof=1)
    if std == 0 or np.isnan(std):
        return None
    
    # 年化 Sharpe Ratio
    return np.mean(excess_returns) / std * np.sqrt(12)


def calc_sortino(monthly_returns, rf_rate=RF_RATE):
    """
    Sortino Ratio（索提諾比率）
    
    公式：Sortino = (月均報酬 - 月無風險利率) / 下行標準差 × √12
    
    意義：只考慮下行風險（負報酬），對投資人更直觀
    - 只看負報酬的波動，而非所有波動
    - 對有持續正報酬的策略更友好
    
    與 Sharpe 的差異：
    - Sharpe：所有波動都是風險
    - Sortino：只有負報酬才是風險
    """
    if monthly_returns is None or len(monthly_returns) < 2:
        return None
    
    monthly_rf = (1 + rf_rate) ** (1/12) - 1
    excess_returns = monthly_returns - monthly_rf
    
    # 下行標準差：只看負的超額報酬
    downside_returns = excess_returns[excess_returns < 0]
    if len(downside_returns) == 0:
        # 全部正報酬： Sortino = 無限大（理想情況）
        return None
    
    downside_std = np.std(downside_returns, ddof=1)
    if downside_std == 0 or np.isnan(downside_std):
        return None
    
    return np.mean(excess_returns) / downside_std * np.sqrt(12)


def calc_max_drawdown(portfolio_values):
    """
    Maximum Drawdown（最大回撤）
    
    公式：MaxDD = max((歷史高點 - 當前低點) / 歷史高點)
    
    意義：從歷史高點到低點的最大跌幅
    - 反映策略在最糟時期的損失
    - 對投資人心理承受力很重要
    - 可用於設定停損或資產配置調整
    """
    if portfolio_values is None or len(portfolio_values) < 2:
        return None
    
    values = np.array(portfolio_values, dtype=float)
    
    # 累計歷史高點
    running_max = np.maximum.accumulate(values)
    
    # 回撤 = (高點 - 當前) / 高點
    drawdowns = (running_max - values) / running_max
    
    return np.max(drawdowns)


def compute_all_metrics(df, years=20.17, rf_rate=RF_RATE):
    """
    一次性計算所有績效指標
    
    輸入：
    - df: DCA 結果 DataFrame（需要有 invested、value 欄位）
    - years: 研究期間（年）
    - rf_rate: 無風險利率（預設 2%）
    
    輸出：dict，包含所有指標
    """
    if df is None or len(df) == 0:
        return None
    
    final = df.iloc[-1]
    invested = float(final['invested'])
    final_value = float(final['value'])
    total_return = (final_value - invested) / invested * 100
    
    # 月報酬率（用於 Sharpe / Sortino）
    monthly_returns = df['value'].pct_change().dropna()
    
    return {
        'Total Invested': invested,
        'Final Value': final_value,
        'Total Return (%)': total_return,
        'CAGR': calc_cagr(invested, final_value, years) * 100,
        'Sharpe Ratio': calc_sharpe(monthly_returns, rf_rate),
        'Sortino Ratio': calc_sortino(monthly_returns, rf_rate),
        'Max Drawdown (%)': calc_max_drawdown(df['value'].values) * 100,
    }


if __name__ == '__main__':
    # 測試：使用虛構數據驗證公式正確性
    print("測試績效指標計算...")
    
    # 簡單測試：投資 $1000 → $2000（1年）
    print(f"  CAGR $1000→$2000, 1yr: {calc_cagr(1000, 2000, 1)*100:.1f}%")
    
    # 月報酬測試
    returns = pd.Series([0.01, -0.02, 0.03, 0.01, 0.02] * 12)
    print(f"  Sharpe (sample): {calc_sharpe(returns):.3f}")
    print(f"  Sortino (sample): {calc_sortino(returns):.3f}")
    
    # 最大回撤測試
    values = [100, 120, 90, 110, 80, 130]
    print(f"  MaxDD [100,120,90,110,80,130]: {calc_max_drawdown(values)*100:.1f}%")
    
    print("✅ 測試完成")
