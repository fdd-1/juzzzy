#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""导出海外思维续费规划表_新版_26年启用报表

调用 smartbi_cli 下载报表，填充开课M计算时间和退费结束时间。

工作流：
  1. 根据 --month 参数（YYYY-MM）计算：
     - 开课M计算时间 = 当月1号 00:00:00
     - 退费结束时间 = 上月最后一天 23:59:59
  2. 动态构造 smartbi_tasks.json 里的 filters.overrides
  3. 调用 smartbi_cli run --config smartbi_tasks.json --task xuewei_warning_renewal_plan
  4. 报表下载到 exports/学位预警_{YYYYMMDD}/
"""
import sys, io, json, subprocess, argparse, datetime as dt
from pathlib import Path
from calendar import monthrange

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent
CONFIG_PATH = SCRIPT_DIR / "smartbi_tasks.json"
SMARTBI_CLI_DIR = Path(r"C:\Users\fengjianyi\Desktop\smartbi-data-cli-internal-20260526\smartbi-data-cli-internal-20260526")
SMARTBI_CLI = SMARTBI_CLI_DIR / "scripts" / "smartbi_cli.py"
BROWSER_EXPORT = SMARTBI_CLI_DIR / "scripts" / "smartbi_browser_export.py"
REPORT_ID = "I2c928087019b236723675f9c019b353f6027505b"  # 海外思维续费规划表_新版_26年启用 (SIMPLE_REPORT)

def log(m): print(m, flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--month", required=True, help="目标月份 YYYY-MM（例：2026-06）")
    args = ap.parse_args()

    # 解析月份
    try:
        y, m = map(int, args.month.split("-"))
        if m < 1 or m > 12:
            raise ValueError
    except Exception:
        log(f"[ERROR] --month 格式错误，需 YYYY-MM：{args.month}")
        sys.exit(1)

    # 计算日期
    # 开课M计算时间 = 当月1号
    kaike_date = dt.date(y, m, 1)
    # 退费结束时间 = 上月最后一天
    if m == 1:
        prev_y, prev_m = y - 1, 12
    else:
        prev_y, prev_m = y, m - 1
    last_day = monthrange(prev_y, prev_m)[1]
    tuifei_date = dt.date(prev_y, prev_m, last_day)

    kaike_str = kaike_date.strftime("%Y-%m-%d")
    tuifei_str = tuifei_date.strftime("%Y-%m-%d 23:59:59")

    log(f"[INFO] 月份：{args.month}")
    log(f"  开课M计算时间：{kaike_str}")
    log(f"  退费结束时间：{tuifei_str}")

    # 改用 smartbi_browser_export.py（Playwright 浏览器渲染），HTTP-only 路径已失效
    if not BROWSER_EXPORT.exists():
        log(f"[ERROR] 找不到 smartbi_browser_export.py: {BROWSER_EXPORT}")
        sys.exit(2)

    # filters: [alias, value, displayValue] 三元组
    filters = [
        ["开课M计算时间", kaike_str, kaike_str],
        ["退费结束时间", tuifei_str, tuifei_str],
        ["学管大区", "海外教学服务部", "海外教学服务部"],
        ["池子节点3", "服务月", "服务月"],
    ]

    # 输出目录（每次新建）
    today_tag = dt.date.today().strftime("%Y%m%d")
    out_dir = SCRIPT_DIR / "exports" / f"学位预警_{today_tag}"
    out_dir.mkdir(parents=True, exist_ok=True)
    timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_path = out_dir / f"海外思维续费规划表_新版_26年启用_{timestamp}.xlsx"

    log(f"[STEP] 调用 smartbi_browser_export.py 下载到 {out_path}")
    cmd = [
        sys.executable,
        str(BROWSER_EXPORT),
        "--report-id", REPORT_ID,
        "--output", str(out_path),
        "--max-rows", "500000",
        "--filters-json", json.dumps(filters, ensure_ascii=False),
        "--json",
    ]
    log(f"  -> filters={filters}")
    result = subprocess.run(cmd, cwd=str(SMARTBI_CLI_DIR), capture_output=True, text=True, encoding="utf-8", errors="replace")
    if result.stdout:
        log(result.stdout)
    if result.stderr:
        log(f"[STDERR] {result.stderr}")
    if result.returncode != 0:
        log(f"[ERROR] smartbi_browser_export 返回 {result.returncode}")
        sys.exit(3)

    if not out_path.exists():
        log(f"[ERROR] 输出文件未生成：{out_path}")
        sys.exit(4)

    log(f"[OK] 报表导出完成：{out_path}")


if __name__ == "__main__":
    main()
