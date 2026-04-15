#!/usr/bin/env python3
"""
===================================================================
DCA 投資策略分析 - 主程式
===================================================================
執行：爬蟲 + 視覺化
"""

import os
import subprocess
import sys


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
PROJECT_DIR = os.path.dirname(SCRIPT_DIR)
PYTHON = sys.executable


def run_step(command, description):
    print(f"\n{description}...")
    result = subprocess.run(command, check=False)
    if result.returncode != 0:
        raise RuntimeError(f"{description}失敗，請先修正後再重新執行")


def main():
    print("=" * 60)
    print("DCA 投資策略分析系統")
    print("=" * 60)

    # Step 1: 爬蟲
    run_step(
        [PYTHON, os.path.join(SCRIPT_DIR, 'scrapers', 'run_all.py')],
        "[Step 1/2] 爬蟲（抓資料）"
    )

    # Step 2: 視覺化
    run_step(
        [PYTHON, os.path.join(SCRIPT_DIR, 'visualization', 'plot_charts.py')],
        "[Step 2/2] 圖表生成"
    )

    print("\n" + "=" * 60)
    print("✅ 完成！")
    print(f"  圖表：{os.path.join(PROJECT_DIR, 'figures')}/")
    print(f"  資料：{os.path.join(PROJECT_DIR, 'data')}/")
    print("=" * 60)


if __name__ == '__main__':
    try:
        main()
    except Exception as exc:
        print(f"\n❌ {exc}")
        sys.exit(1)
