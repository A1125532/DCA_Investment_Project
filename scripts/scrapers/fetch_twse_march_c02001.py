# -*- coding: utf-8 -*-
"""
下載證交所「市場交易月報」中：
證券市場統計概要與市場總市值、投資報酬率、本益比、殖利率一覽表（C02001）
每年 3 月版本（ZIP），並解壓到本機。

檔案命名規則（證交所 staticFiles）：
  https://www.twse.com.tw/staticFiles/inspection/inspection/02/001/YYYY03_C02001.zip

來源頁面（可查詢同系列其他月份）：
  https://www.twse.com.tw/zh/trading/statistics/index02.html

請勿過於頻繁請求；預設每次下載間隔數秒。
"""
from __future__ import annotations

import json
import sys
import time
import urllib.request
import zipfile
from pathlib import Path

BASE = "https://www.twse.com.tw/staticFiles/inspection/inspection/02/001"
USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
    "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
)
SLEEP_SEC = 4.0


def download(url: str, dest: Path) -> bool:
    req = urllib.request.Request(url, headers={"User-Agent": USER_AGENT})
    try:
        with urllib.request.urlopen(req, timeout=120) as r:
            data = r.read()
        dest.write_bytes(data)
        return True
    except Exception as e:
        print(f"  [失敗] {url}\n         {e}", file=sys.stderr)
        return False


def main() -> None:
    project_dir = Path(__file__).resolve().parents[2]
    out_zip = project_dir / "data" / "twse_march_C02001" / "zip"
    out_x = project_dir / "data" / "twse_march_C02001" / "extracted"
    manifest = project_dir / "data" / "twse_march_C02001" / "manifest.jsonl"
    out_zip.mkdir(parents=True, exist_ok=True)
    out_x.mkdir(parents=True, exist_ok=True)
    if manifest.exists():
        manifest.unlink()

    years = range(2006, 2027)
    rows = []

    for i, y in enumerate(years):
        if i > 0:
            time.sleep(SLEEP_SEC)
        name = f"{y}03_C02001.zip"
        url = f"{BASE}/{name}"
        zpath = out_zip / name
        print(f"[{y}] 下載 {name} …")
        ok = download(url, zpath)
        row = {"year": y, "month": 3, "code": "C02001", "url": url, "zip_ok": ok}
        if ok and zipfile.is_zipfile(zpath):
            target = out_x / f"{y}_03"
                    target.mkdir(parents=True, exist_ok=True)
            try:
                with zipfile.ZipFile(zpath, "r") as zf:
                    zf.extractall(target)
                    row["extracted_to"] = str(target.relative_to(project_dir))
                    row["members"] = zf.namelist()
            except zipfile.BadZipFile as e:
                row["extract_error"] = str(e)
        rows.append(row)
        with manifest.open("a", encoding="utf-8") as f:
            f.write(json.dumps(row, ensure_ascii=False) + "\n")

    ok_n = sum(1 for r in rows if r.get("zip_ok"))
    print(f"\n完成：成功 {ok_n} / {len(rows)} 。輸出目錄：{out_x}")


if __name__ == "__main__":
    main()
