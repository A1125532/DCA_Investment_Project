#!/usr/bin/env python3
"""
DCA 策略實作（精簡版）

目前專案主流程改為 experiments/walkforward_yearly_calibration.py，
此檔僅保留 Value Averaging 的基礎實作。
"""

import pandas as pd

MONTHLY = 1000


def _prepare_prices(prices: pd.Series) -> pd.Series:
    p = prices.copy()
    if not isinstance(p.index, pd.DatetimeIndex):
        p.index = pd.to_datetime(p.index, errors='coerce')
    p = p.dropna()
    return p


def dca_value_averaging(prices, monthly_base=MONTHLY, max_invest=5000, months=None):
    """
    Value Averaging（基礎版）

    I_t = clip(Gap_t, 0.3 * base, max_invest)
    """
    prices = _prepare_prices(prices)
    if months is not None:
        prices = prices[(prices.index >= pd.Timestamp(months[0])) &
                        (prices.index <= pd.Timestamp(months[1]))]

    monthly = prices.resample('ME').last().dropna()
    if len(monthly) == 0:
        return pd.DataFrame(columns=[
            'date', 'price', 'shares', 'invested', 'value', 'return_pct', 'monthly_invest'
        ])

    results = []
    total_shares = 0.0
    total_invested = 0.0

    for i, (date, price) in enumerate(monthly.items()):
        target_value = monthly_base * (i + 1)
        current_value = total_shares * price
        gap = target_value - current_value

        investment = max(min(gap, max_invest), monthly_base * 0.3)

        if investment > 0:
            total_shares += investment / price
            total_invested += investment

        value = total_shares * price
        results.append({
            'date': date,
            'price': price,
            'shares': total_shares,
            'invested': total_invested,
            'value': value,
            'return_pct': (value - total_invested) / max(total_invested, 1) * 100,
            'monthly_invest': investment,
        })

    return pd.DataFrame(results)
