"""
从 BI 系统下载学情积分核算所需的两张报表。

使用 smartbi-data-cli 方式（Playwright browser export），直接通过 report_id
打开报表、设置筛选条件、导出 Excel。

报表1: 海外思维续费规划表_新版_26年启用
  report_id: I2c928087019b236723675f9c019b353f6027505b
  筛选: 当前课包签单年月开始/结束 = 开始日期~结束日期
        当前课包签单时间开始/结束 = 开始日期~结束日期

报表2: 海外思维学员上课明细
  report_id: I2c9280870198767976798e4f0198889e7cc27654
  筛选: 开始日期 = 当月1号, 结束日期 = 当月15号(或月底)

用法:
  python fetch_reports.py --start 2026-05-01 --end 2026-05-15
"""

import argparse
import asyncio
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from fetch_reports_smartbi import fetch_all, OUTPUT_DIR


def main():
    parser = argparse.ArgumentParser(description="下载学情积分核算所需BI报表")
    parser.add_argument("--start", required=True, help="开始日期 YYYY-MM-DD (当月1号)")
    parser.add_argument("--end", required=True, help="结束日期 YYYY-MM-DD (当月15号)")
    parser.add_argument("--headful", action="store_true", help="显示浏览器窗口（调试用）")
    args = parser.parse_args()

    Path(OUTPUT_DIR).mkdir(parents=True, exist_ok=True)

    results = asyncio.run(fetch_all(args.start, args.end, headless=not args.headful))

    failed = [k for k, v in results.items() if v["status"] != "ok"]
    if failed:
        print(f"\n[WARN] 失败的报表: {failed}")
        sys.exit(1)


if __name__ == "__main__":
    main()
