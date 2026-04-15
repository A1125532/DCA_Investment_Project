#!/usr/bin/env python3
"""
===================================================================
一次執行所有爬蟲（便利腳本）
===================================================================
台股基本面來源優先序：
1) 證交所 C02001 匯出 xlsx（含 PE/PB/殖利率）
2) Goodinfo（PE/PB，當 xlsx 不可用時 fallback）
===================================================================
"""

import os
import subprocess
import sys


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def _run(script_name):
    path = os.path.join(SCRIPT_DIR, script_name)
    result = subprocess.run([sys.executable, path], check=False)
    if result.returncode != 0:
        raise RuntimeError(f"{script_name} 執行失敗（exit code={result.returncode}）")


if __name__ == '__main__':
    print("=" * 60)
    print("DCA 爬蟲程式（一次性執行）")
    print("=" * 60)

    import_twse = os.path.join(SCRIPT_DIR, "import_taiex_from_twse_xlsx.py")

    scripts = [
        ('[1/4] Yahoo Finance 股價...', 'fetch_yahoo_prices.py'),
        ('[3/4] multpl.com S&P 500 基本面...', 'fetch_sp500_fundamental.py'),
        ('[4/4] Nikkei Indexes 日經 225 基本面...', 'fetch_nikkei_fundamental.py'),
    ]

    print("\n[2/4] 台股基本面（優先證交所 C02001 xlsx）...")
    twse_result = subprocess.run([sys.executable, import_twse], check=False)
    if twse_result.returncode == 0:
        print("  ✅ 已從 C02001 xlsx 匯入（含 DividendYield）")
    elif twse_result.returncode == 2:
        print("  ⚠️ 未找到完整 xlsx，改用 Goodinfo（PE/PB）")
        _run("fetch_taiex_fundamental.py")
    else:
        raise RuntimeError(f"台股 xlsx 匯入失敗（exit code={twse_result.returncode}）")

    for step, script in scripts:
        print(f"\n{step}")
        _run(script)

    print("\n" + "=" * 60)
    print("✅ 所有資料爬蟲完成！")
    print("=" * 60)
