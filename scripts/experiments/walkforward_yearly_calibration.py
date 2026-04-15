#!/usr/bin/env python3
import itertools
import math
from pathlib import Path

import numpy as np
import pandas as pd

PROJECT = Path(__file__).resolve().parents[2]
DATA_DIR = PROJECT / "data"
FUND_DIR = DATA_DIR / "fundamentals"
OUT_DIR = PROJECT / "results" / "performance"
OUT_DIR.mkdir(parents=True, exist_ok=True)

START = "2006-01-01"
END = "2026-03-31"
BASE = 1000.0
RF = 0.02
FIRST_TEST_YEAR = 2006
MIN_TRAIN_MONTHS = 10

MARKETS = {
    "TAIEX": {
        "price": DATA_DIR / "taiex_price.csv",
        "fund_candidates": [
            FUND_DIR / "taiex_fundamental.csv",
        ],
    },
    "SP500": {
        "price": DATA_DIR / "sp500_price.csv",
        "fund_candidates": [
            FUND_DIR / "sp500_fundamental.csv",
            FUND_DIR / "sp500_fundamental_monthly.csv",
            FUND_DIR / "sp500_fundamental_merged.csv",
        ],
    },
    "Nikkei225": {
        "price": DATA_DIR / "nikkei225_price.csv",
        "fund_candidates": [
            FUND_DIR / "nikkei225_fundamental.csv",
            FUND_DIR / "nikkei225_fundamental_monthly.csv",
        ],
    },
}

STRATEGIES = [
    "Value Averaging",
    "Threshold DCA (3factor)",
    "Model-Driven DCA",
]

DEFAULT_PARAMS = {
    "Value Averaging": {"min_ratio": 0.3, "max_mult": 5.0},
    "Threshold DCA (3factor)": {"q": 0.2, "boost": 1.5, "reduce": 0.5},
    "Model-Driven DCA": {"alpha": 0.35, "clip_low": 0.5, "clip_high": 1.5, "w_pb": 1.0, "w_dy": 1.0},
}


def pick_existing(paths):
    for p in paths:
        if p.exists():
            return p
    raise FileNotFoundError(f"No candidate file exists: {paths}")


def load_price(path: Path) -> pd.Series:
    df = pd.read_csv(path)
    date_col = df.columns[0]
    close_col = "Close" if "Close" in df.columns else ("Adj Close" if "Adj Close" in df.columns else df.columns[-1])
    out = df[[date_col, close_col]].copy()
    out.columns = ["date", "close"]
    out["date"] = pd.to_datetime(out["date"], errors="coerce")
    out["close"] = pd.to_numeric(out["close"], errors="coerce")
    out = out.dropna().set_index("date").sort_index()
    out = out.loc[(out.index >= pd.Timestamp(START)) & (out.index <= pd.Timestamp(END))]
    return out["close"].resample("ME").last().dropna()


def load_fund(path: Path) -> pd.DataFrame:
    df = pd.read_csv(path)
    date_col = df.columns[0]
    df = df.rename(columns={date_col: "date"})
    rename = {
        "PER": "PE",
        "per": "PE",
        "trailingPE": "PE",
        "trailingpe": "PE",
        "PBR": "PB",
        "pb": "PB",
        "priceToBook": "PB",
        "price_to_book": "PB",
        "dividendYield": "DividendYield",
        "dividendyield": "DividendYield",
    }
    df = df.rename(columns=rename)
    keep = [c for c in ["date", "PE", "PB", "DividendYield"] if c in df.columns]
    df = df[keep].copy()
    df["date"] = pd.to_datetime(df["date"], errors="coerce")
    for c in ["PE", "PB", "DividendYield"]:
        if c in df.columns:
            df[c] = pd.to_numeric(df[c], errors="coerce")
    for c in ["PE", "PB", "DividendYield"]:
        if c not in df.columns:
            df[c] = np.nan
    df = df.dropna(subset=["date"]).sort_values("date")
    # Use one-month publication lag to avoid look-ahead.
    df["period"] = df["date"].dt.to_period("M") + 1
    df = df[["period", "PE", "PB", "DividendYield"]].drop_duplicates("period", keep="last")
    df.index = df["period"].dt.to_timestamp("M")
    return df[["PE", "PB", "DividendYield"]].sort_index()


def safe_monthly_irr(cashflows):
    lo, hi = -0.95, 5.0

    def npv(r):
        total = 0.0
        base = 1.0 + r
        for t, cf in enumerate(cashflows):
            denom = base ** t
            if denom == 0.0 or not np.isfinite(denom):
                return np.nan
            total += cf / denom
        return total

    f_lo = npv(lo)
    f_hi = npv(hi)
    if not np.isfinite(f_lo) or not np.isfinite(f_hi):
        return np.nan
    if f_lo * f_hi > 0:
        return np.nan

    for _ in range(120):
        mid = (lo + hi) / 2.0
        f_mid = npv(mid)
        if abs(f_mid) < 1e-10:
            return mid
        if f_lo * f_mid <= 0:
            hi = mid
            f_hi = f_mid
        else:
            lo = mid
            f_lo = f_mid
    return (lo + hi) / 2.0


def compute_metrics(df: pd.DataFrame, include_irr: bool = True) -> dict:
    if len(df) == 0:
        return {
            "Total Invested": np.nan,
            "Final Value": np.nan,
            "Total Return %": np.nan,
            "CAGR %": np.nan,
            "Sharpe": np.nan,
            "Sortino": np.nan,
            "Max Drawdown %": np.nan,
            "IRR %": np.nan,
        }

    invested = float(df["invested"].iloc[-1])
    final_value = float(df["value"].iloc[-1])
    total_return = (final_value - invested) / invested * 100 if invested > 0 else np.nan
    years = max((len(df) - 1) / 12.0, 1e-9)
    cagr = ((final_value / invested) ** (1.0 / years) - 1.0) * 100 if invested > 0 else np.nan

    v = df["value"].to_numpy(dtype=float)
    rets = pd.Series(v).pct_change().dropna()
    m_rf = (1.0 + RF) ** (1.0 / 12.0) - 1.0
    excess = rets - m_rf
    vol = rets.std(ddof=1)
    sharpe = float(excess.mean() / vol * math.sqrt(12.0)) if len(rets) > 1 and vol > 0 else np.nan

    downside = excess[excess < 0]
    dvol = downside.std(ddof=1)
    sortino = float(excess.mean() / dvol * math.sqrt(12.0)) if len(downside) > 1 and dvol > 0 else np.nan

    peak = np.maximum.accumulate(v)
    maxdd = float(np.max((peak - v) / np.maximum(peak, 1e-9)) * 100.0)

    irr = np.nan
    if include_irr:
        cashflows = (-df["monthly_invest"]).to_list()
        cashflows[-1] = cashflows[-1] + final_value
        m_irr = safe_monthly_irr(cashflows)
        irr = ((1.0 + m_irr) ** 12.0 - 1.0) * 100.0 if pd.notna(m_irr) else np.nan

    return {
        "Total Invested": invested,
        "Final Value": final_value,
        "Total Return %": total_return,
        "CAGR %": cagr,
        "Sharpe": sharpe,
        "Sortino": sortino,
        "Max Drawdown %": maxdd,
        "IRR %": irr,
    }


def objective_value(metric_dict):
    sharpe = metric_dict["Sharpe"] if pd.notna(metric_dict["Sharpe"]) else -9.0
    cagr = metric_dict["CAGR %"] if pd.notna(metric_dict["CAGR %"]) else -99.0
    maxdd = metric_dict["Max Drawdown %"] if pd.notna(metric_dict["Max Drawdown %"]) else 100.0
    return float(sharpe + 0.05 * cagr - 0.02 * maxdd)


def amount_value_averaging(date, price, shares, va_target, params):
    min_ratio = float(params["min_ratio"])
    max_mult = float(params["max_mult"])
    new_target = va_target + BASE
    gap = new_target - shares * price
    amount = float(np.clip(gap, BASE * min_ratio, BASE * max_mult))
    return amount, {"va_target": new_target}


def amount_threshold(date, fund, params):
    q = float(params["q"])
    boost = float(params["boost"])
    reduce = float(params["reduce"])

    row = fund.loc[date]
    train = fund.loc[fund.index < date].dropna()
    if len(train) < MIN_TRAIN_MONTHS or row.isna().any():
        return BASE

    pe_l, pe_h = train["PE"].quantile(q), train["PE"].quantile(1.0 - q)
    pb_l, pb_h = train["PB"].quantile(q), train["PB"].quantile(1.0 - q)
    dy_l, dy_h = train["DividendYield"].quantile(q), train["DividendYield"].quantile(1.0 - q)

    underval = int(row["PE"] <= pe_l) + int(row["PB"] <= pb_l) + int(row["DividendYield"] >= dy_h)
    overval = int(row["PE"] >= pe_h) + int(row["PB"] >= pb_h) + int(row["DividendYield"] <= dy_l)

    if underval >= 2:
        return BASE * boost
    if overval >= 2:
        return BASE * reduce
    return BASE


def amount_model(date, fund, params):
    alpha = float(params["alpha"])
    clip_low = float(params["clip_low"])
    clip_high = float(params["clip_high"])
    w_pb = float(params["w_pb"])
    w_dy = float(params["w_dy"])

    row = fund.loc[date]
    train = fund.loc[fund.index < date].dropna()
    if len(train) < MIN_TRAIN_MONTHS or row.isna().any():
        return BASE

    mu = train.mean()
    sd = train.std(ddof=1).replace(0, 1.0)
    z = (row - mu) / sd
    score = -z["PE"] - w_pb * z["PB"] + w_dy * z["DividendYield"]
    mult = float(np.clip(1.0 + alpha * score, clip_low, clip_high))
    return BASE * mult


def run_strategy(price, fund, strategy, params_by_year):
    shares = 0.0
    invested = 0.0
    va_target = 0.0
    rows = []

    for date, px in price.items():
        params = params_by_year.get(date.year)
        if params is None:
            continue

        if strategy == "Value Averaging":
            amount, state = amount_value_averaging(date, px, shares, va_target, params)
            va_target = state["va_target"]
        elif strategy == "Threshold DCA (3factor)":
            amount = amount_threshold(date, fund, params)
        elif strategy == "Model-Driven DCA":
            amount = amount_model(date, fund, params)
        else:
            raise ValueError(f"Unknown strategy: {strategy}")

        amount = max(float(amount), 0.0)
        if amount > 0:
            shares += amount / px
            invested += amount
        value = shares * px
        rows.append((date, px, shares, invested, value, amount))

    out = pd.DataFrame(rows, columns=["date", "price", "shares", "invested", "value", "monthly_invest"])
    return out


def calibrate_one_year(train_price, fund, strategy):
    if strategy == "Value Averaging":
        grid = [
            {"min_ratio": mn, "max_mult": mx}
            for mn, mx in itertools.product([0.0, 0.3], [2.0, 3.0, 5.0])
            if mx >= mn + 0.5
        ]
    elif strategy == "Threshold DCA (3factor)":
        grid = [
            {"q": q, "boost": b, "reduce": r}
            for q, b, r in itertools.product([0.2, 0.3], [1.3, 1.5], [0.5])
        ]
    elif strategy == "Model-Driven DCA":
        grid = [
            {"alpha": a, "clip_low": lo, "clip_high": hi, "w_pb": wpb, "w_dy": wdy}
            for a, lo, hi, wpb, wdy in itertools.product(
                [0.25, 0.35, 0.45],
                [0.5, 0.7],
                [1.3, 1.5],
                [1.0],
                [1.0],
            )
            if lo < hi
        ]
    else:
        raise ValueError(f"Unknown strategy: {strategy}")

    best_params = grid[0]
    best_score = -1e18
    best_metrics = None

    year = int(train_price.index[-1].year) + 1
    params_by_year = {d.year: best_params for d in train_price.index}
    params_by_year[year] = best_params

    for params in grid:
        params_by_year = {d.year: params for d in train_price.index}
        df = run_strategy(train_price, fund, strategy, params_by_year)
        m = compute_metrics(df, include_irr=False)
        score = objective_value(m)
        if score > best_score:
            best_score = score
            best_params = params
            best_metrics = m

    return best_params, best_score, best_metrics


def run_market(market, price, fund):
    idx = price.index
    years = sorted(idx.year.unique())
    test_years = [y for y in years if y >= FIRST_TEST_YEAR]

    results_rows = []
    param_rows = []
    curve_rows = []

    for strategy in STRATEGIES:
        params_by_year = {}
        for y in test_years:
            if y == FIRST_TEST_YEAR:
                params = DEFAULT_PARAMS[strategy]
                params_by_year[y] = params
                param_rows.append(
                    {
                        "Market": market,
                        "Strategy": strategy,
                        "Year": y,
                        "Train End": "",
                        "Train Score": np.nan,
                        "Train Sharpe": np.nan,
                        "Train CAGR %": np.nan,
                        "Train MaxDD %": np.nan,
                        "Params": str(params),
                        "Calibration Mode": "default_start_year",
                    }
                )
                continue

            train_start = pd.Timestamp(START)
            train_end = pd.Timestamp(f"{y-1}-12-31")
            train_price = price.loc[(price.index >= train_start) & (price.index <= train_end)]

            if len(train_price) < MIN_TRAIN_MONTHS:
                params = DEFAULT_PARAMS[strategy]
                params_by_year[y] = params
                param_rows.append(
                    {
                        "Market": market,
                        "Strategy": strategy,
                        "Year": y,
                        "Train End": train_end.date().isoformat(),
                        "Train Score": np.nan,
                        "Train Sharpe": np.nan,
                        "Train CAGR %": np.nan,
                        "Train MaxDD %": np.nan,
                        "Params": str(params),
                        "Calibration Mode": "fallback_default_insufficient_data",
                    }
                )
                continue

            params, score, m = calibrate_one_year(train_price, fund, strategy)
            params_by_year[y] = params
            param_rows.append(
                {
                    "Market": market,
                    "Strategy": strategy,
                    "Year": y,
                    "Train End": train_end.date().isoformat(),
                    "Train Score": score,
                    "Train Sharpe": m["Sharpe"] if m is not None else np.nan,
                    "Train CAGR %": m["CAGR %"] if m is not None else np.nan,
                    "Train MaxDD %": m["Max Drawdown %"] if m is not None else np.nan,
                    "Params": str(params),
                        "Calibration Mode": "expanding_from_start",
                    }
                )

        test_price = price.loc[price.index >= pd.Timestamp(f"{FIRST_TEST_YEAR}-01-01")]
        out = run_strategy(test_price, fund, strategy, params_by_year)
        m = compute_metrics(out)
        results_rows.append(
            {
                "Market": market,
                "Strategy": strategy,
                **m,
            }
        )
        for _, r in out.iterrows():
            curve_rows.append(
                {
                    "Market": market,
                    "Strategy": strategy,
                    "Date": r["date"],
                    "Price": r["price"],
                    "Shares": r["shares"],
                    "Invested": r["invested"],
                    "Value": r["value"],
                    "Monthly Invest": r["monthly_invest"],
                }
            )

    return pd.DataFrame(results_rows), pd.DataFrame(param_rows), pd.DataFrame(curve_rows)


def main():
    all_results = []
    all_params = []
    all_curves = []

    for market, cfg in MARKETS.items():
        price = load_price(cfg["price"])
        fund_path = pick_existing(cfg["fund_candidates"])
        fund = load_fund(fund_path).reindex(price.index)

        result_df, param_df, curve_df = run_market(market, price, fund)
        all_results.append(result_df)
        all_params.append(param_df)
        all_curves.append(curve_df)

    summary = pd.concat(all_results, ignore_index=True)
    params = pd.concat(all_params, ignore_index=True)
    curves = pd.concat(all_curves, ignore_index=True)

    summary["Sharpe Rank (1=best)"] = summary.groupby("Market")["Sharpe"].rank(ascending=False, method="min").astype(int)
    summary["Total Return Rank (1=best)"] = summary.groupby("Market")["Total Return %"].rank(ascending=False, method="min").astype(int)
    summary["CAGR Rank (1=best)"] = summary.groupby("Market")["CAGR %"].rank(ascending=False, method="min").astype(int)

    summary_path = OUT_DIR / "walkforward_expanding_3strategies.csv"
    params_path = OUT_DIR / "walkforward_expanding_param_log.csv"
    curves_path = OUT_DIR / "walkforward_expanding_curves.csv"

    summary.to_csv(summary_path, index=False)
    params.to_csv(params_path, index=False)
    curves.to_csv(curves_path, index=False)

    print("Saved:")
    print(summary_path)
    print(params_path)
    print(curves_path)
    print()
    print(summary.round(4).to_string(index=False))


if __name__ == "__main__":
    main()
