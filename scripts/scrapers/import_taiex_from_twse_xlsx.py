#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
從證交所 C02001 月報輸出（三份 xlsx）合併為 taiex_fundamental.csv

預期檔名（與 scripts/scrapers/build_twse_monthly_valuation.py 一致）：
  - 本益比_PE_大盤_每月_2006-2026.xlsx
  - 股價淨值比_PB_大盤_每月_2006-2026.xlsx
  - 殖利率_大盤_每月_2006-2026.xlsx

工作表：「資料」；欄位：西元年、月份、數值_*、狀態（可選）

輸出：data/fundamentals/taiex_fundamental.csv（date, PE, PB, DividendYield）

路徑優先順序：
  1. 環境變數 TWSE_C02001_OUTPUT
  2. <專案根>/data/twse_c02001_monthly/output
  3. <專案根>/scripts/twse_c02001_workspace/data/twse_c02001_monthly/output（舊位置，相容）
  4. <專案根>/DCA/data/twse_c02001_monthly/output（舊位置，相容）
  5. <專案根>/../DCA/data/twse_c02001_monthly/output（舊位置，相容）
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

import pandas as pd

PROJECT_DIR = Path(__file__).resolve().parents[2]
FUND_DIR = PROJECT_DIR / "data" / "fundamentals"
OUT_CSV = FUND_DIR / "taiex_fundamental.csv"

PE_NAME = "本益比_PE_大盤_每月_2006-2026.xlsx"
PB_NAME = "股價淨值比_PB_大盤_每月_2006-2026.xlsx"
DY_NAME = "殖利率_大盤_每月_2006-2026.xlsx"


def resolve_input_dir(explicit: str | None) -> Path:
    if explicit:
        return Path(explicit)
    env = os.environ.get("TWSE_C02001_OUTPUT", "").strip()
    if env:
        return Path(env)
    candidates = [
        PROJECT_DIR / "data" / "twse_c02001_monthly" / "output",
        PROJECT_DIR / "scripts" / "twse_c02001_workspace" / "data" / "twse_c02001_monthly" / "output",
        PROJECT_DIR / "DCA" / "data" / "twse_c02001_monthly" / "output",
        PROJECT_DIR.parent / "DCA" / "data" / "twse_c02001_monthly" / "output",
    ]
    for p in candidates:
        if p.is_dir() and (p / PE_NAME).is_file():
            return p
    return candidates[0]


def _read_metric(path: Path, value_col: str, rename_to: str) -> pd.DataFrame:
    if not path.is_file():
        raise FileNotFoundError(f"找不到檔案：{path}")
    df = pd.read_excel(path, sheet_name="資料", engine="openpyxl")
    if "西元年" not in df.columns or "月份" not in df.columns:
        raise ValueError(f"{path.name} 缺少 西元年 或 月份 欄位")
    if value_col not in df.columns:
        raise ValueError(f"{path.name} 缺少欄位 {value_col}")
    if "狀態" in df.columns:
        df = df[df["狀態"].astype(str) == "ok"].copy()
    out = df[["西元年", "月份", value_col]].copy()
    out = out.rename(columns={value_col: rename_to})
    for c in ("西元年", "月份"):
        out[c] = pd.to_numeric(out[c], errors="coerce")
    out[rename_to] = pd.to_numeric(out[rename_to], errors="coerce")
    out = out.dropna(subset=["西元年", "月份"])
    out["西元年"] = out["西元年"].astype(int)
    out["月份"] = out["月份"].astype(int)
    return out.drop_duplicates(subset=["西元年", "月份"], keep="last")


def merge_fundamentals(input_dir: Path) -> pd.DataFrame:
    pe_path = input_dir / PE_NAME
    pb_path = input_dir / PB_NAME
    dy_path = input_dir / DY_NAME

    pe = _read_metric(pe_path, "數值_倍", "PE")
    pb = _read_metric(pb_path, "數值_倍", "PB")
    dy = _read_metric(dy_path, "數值_百分比", "DividendYield")

    m = pe.merge(pb, on=["西元年", "月份"], how="outer")
    m = m.merge(dy, on=["西元年", "月份"], how="outer")
    m["date"] = pd.to_datetime(
        dict(year=m["西元年"], month=m["月份"], day=1), errors="coerce"
    )
    m = m.dropna(subset=["date"]).sort_values("date")
    out = m[["date", "PE", "PB", "DividendYield"]].copy()
    out["date"] = out["date"].dt.strftime("%Y-%m-%d")
    return out


def main() -> int:
    ap = argparse.ArgumentParser(description="合併證交所 C02001 三份 xlsx → taiex_fundamental.csv")
    ap.add_argument(
        "--input-dir",
        type=str,
        default=None,
        help="含三份 xlsx 的目錄（預設依 TWSE_C02001_OUTPUT 或 data/twse_c02001_monthly/output）",
    )
    args = ap.parse_args()
    input_dir = resolve_input_dir(args.input_dir)

    missing = [
        n
        for n in (PE_NAME, PB_NAME, DY_NAME)
        if not (input_dir / n).is_file()
    ]
    if missing:
        print(
            f"略過台股 xlsx 匯入：目錄 {input_dir} 缺少：{', '.join(missing)}",
            file=sys.stderr,
        )
        return 2

    try:
        df = merge_fundamentals(input_dir)
    except Exception as e:
        print(f"台股 xlsx 合併失敗：{e}", file=sys.stderr)
        return 1

    FUND_DIR.mkdir(parents=True, exist_ok=True)
    df.to_csv(OUT_CSV, index=False)
    print(f"已寫入 {OUT_CSV}（{len(df)} 筆），來源：{input_dir}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
