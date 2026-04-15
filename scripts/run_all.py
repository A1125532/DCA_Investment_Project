#!/usr/bin/env python3
"""
===================================================================
一次執行所有爬蟲（便利腳本）
===================================================================
"""

import os
import subprocess
import sys


if __name__ == '__main__':
    # 保留舊入口：統一轉交給 scripts/scrapers/run_all.py
    script_dir = os.path.dirname(os.path.abspath(__file__))
    target = os.path.join(script_dir, "scrapers", "run_all.py")
    result = subprocess.run(
        [sys.executable, target],
        check=False,
    )
    if result.returncode != 0:
        raise SystemExit(result.returncode)
