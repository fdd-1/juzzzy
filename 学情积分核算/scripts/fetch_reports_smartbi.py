#!/usr/bin/env python3
"""从 BI 系统下载学情积分核算所需的两张报表（smartbi-data-cli 方式）。

使用 SmartBI browser export 方法，通过 Playwright 打开 SIMPLE_REPORT，
设置筛选条件后导出 Excel。

报表1: 海外思维续费规划表_新版_26年启用
  report_id: I2c928087019b236723675f9c019b353f6027505b
  筛选: 当前课包签单年月开始/结束, 当前课包签单时间开始/结束

报表2: 海外思维学员上课明细
  report_id: I2c9280870198767976798e4f0198889e7cc27654
  筛选: 开始日期, 结束日期

用法:
  python fetch_reports_smartbi.py --start 2026-05-01 --end 2026-05-15

环境变量:
  SMARTBI_USERNAME  BI 账号
  SMARTBI_PASSWORD  BI 密码
"""

from __future__ import annotations

import argparse
import asyncio
import json
import os
import sys
import time
from datetime import datetime
from pathlib import Path
from typing import Any

SCRIPT_DIR = Path(__file__).resolve().parent
PROJECT_ROOT = SCRIPT_DIR.parent
sys.path.insert(0, str(SCRIPT_DIR))

from smartbi_browser_export import export_simple_report, SmartbiBrowserExportError

CONFIG_PATH = PROJECT_ROOT / "configs" / "smartbi_simple_report_tasks.json"
OUTPUT_DIR = PROJECT_ROOT / "01_bi_exports"

USERNAME = os.environ.get("SMARTBI_USERNAME", "76218")
PASSWORD = os.environ.get("SMARTBI_PASSWORD", "123456")


def load_config() -> dict:
    with CONFIG_PATH.open("r", encoding="utf-8") as f:
        return json.load(f)


def build_filters_for_task(task: dict, start_date: str, end_date: str) -> list[list[str]]:
    """根据 task config 的 date_mapping 构建 [alias, value, displayValue] 列表。"""
    mapping = task.get("filters", {}).get("date_mapping", {})
    filters: list[list[str]] = []
    for field_name, role in mapping.items():
        if role == "start_date":
            filters.append([field_name, start_date, start_date])
        elif role == "end_date":
            filters.append([field_name, end_date, end_date])
    return filters


async def fetch_one_report(
    task_name: str,
    task: dict,
    start_date: str,
    end_date: str,
    output_dir: Path,
    headless: bool = True,
) -> dict[str, Any]:
    """下载单张报表，返回结果 dict。"""
    report = task["report"]
    report_id = report["id"]
    filename = task.get("output", {}).get("filename", f"{task_name}.xlsx")
    max_rows = task.get("max_rows", 50000)
    filters = build_filters_for_task(task, start_date, end_date)

    output_path = output_dir / filename
    print(f"  报表: {report.get('path', task_name)}")
    print(f"  report_id: {report_id}")
    print(f"  筛选: {filters}")
    print(f"  输出: {output_path}")

    t0 = time.time()
    result = await export_simple_report(
        username=USERNAME,
        password=PASSWORD,
        report_id=report_id,
        output_path=output_path,
        max_rows=max_rows,
        browser_channel="chrome",
        headless=headless,
        filters=filters,
    )
    elapsed = time.time() - t0
    result["elapsed_sec"] = round(elapsed, 1)
    print(f"  完成! {result.get('bytes', 0)} bytes, {elapsed:.1f}s")
    return result


async def fetch_all(start_date: str, end_date: str, headless: bool = True) -> dict:
    """下载所有启用的报表任务。"""
    config = load_config()
    tasks = config.get("tasks", {})
    output_dir = OUTPUT_DIR
    output_dir.mkdir(parents=True, exist_ok=True)

    results = {}
    for task_name, task in tasks.items():
        if not task.get("enabled", True):
            continue
        print(f"\n{'='*60}")
        print(f"下载: {task.get('description', task_name)}")
        print(f"{'='*60}")
        try:
            result = await fetch_one_report(
                task_name, task, start_date, end_date, output_dir, headless
            )
            results[task_name] = {"status": "ok", **result}
        except SmartbiBrowserExportError as e:
            print(f"  [ERROR] {e}")
            results[task_name] = {"status": "error", "error": str(e)}
        except Exception as e:
            print(f"  [ERROR] {type(e).__name__}: {e}")
            results[task_name] = {"status": "error", "error": str(e)}

    success = sum(1 for r in results.values() if r["status"] == "ok")
    total = len(results)
    print(f"\n{'='*60}")
    print(f"结果: {success}/{total} 成功")
    if success == total:
        print(f"[OK] 所有报表下载完成，文件在: {output_dir}")
    else:
        print("[WARN] 部分报表下载失败，请检查上方日志")
    return results


def main():
    parser = argparse.ArgumentParser(
        description="下载学情积分核算所需 BI 报表（smartbi-data-cli 方式）"
    )
    parser.add_argument("--start", required=True, help="开始日期 YYYY-MM-DD")
    parser.add_argument("--end", required=True, help="结束日期 YYYY-MM-DD")
    parser.add_argument("--headful", action="store_true", help="显示浏览器窗口（调试用）")
    parser.add_argument("--json", action="store_true", help="输出 JSON 结果")
    args = parser.parse_args()

    results = asyncio.run(fetch_all(args.start, args.end, headless=not args.headful))

    if args.json:
        print(json.dumps(results, ensure_ascii=False, indent=2, default=str))

    failed = [k for k, v in results.items() if v["status"] != "ok"]
    sys.exit(1 if failed else 0)


if __name__ == "__main__":
    main()
