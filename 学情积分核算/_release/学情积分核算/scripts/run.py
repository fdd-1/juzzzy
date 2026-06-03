"""
学情积分核算 - 一键运行脚本

整合 BI 取数 + 数据处理的完整流程。

用法:
  python run.py --start 2026-05-01 --end 2026-05-15
  python run.py --auto  (根据当前日期自动计算区间)
  python run.py --start 2026-05-01 --end 2026-05-15 --skip-fetch  (跳过BI取数，直接处理)

定时任务逻辑:
  每月1号触发: 计算上月16号 ~ 上月最后一天
  每月16号触发: 计算本月1号 ~ 本月15号
"""

import argparse
import subprocess
import sys
import glob
import calendar
from pathlib import Path
from datetime import datetime, date, timedelta

BASE_DIR = Path(__file__).resolve().parent.parent
SCRIPTS_DIR = BASE_DIR / "scripts"
EXPORTS_DIR = BASE_DIR / "01_bi_exports"
OUTPUT_DIR = BASE_DIR / "03_output"


def calc_date_range_auto() -> tuple[str, str]:
    """根据当前日期自动计算统计区间"""
    today = date.today()
    if today.day <= 15:
        # 1号触发: 上月16号 ~ 上月最后一天
        last_month_last = today.replace(day=1) - timedelta(days=1)
        start = last_month_last.replace(day=16)
        end = last_month_last
    else:
        # 16号触发: 本月1号 ~ 本月15号
        start = today.replace(day=1)
        end = today.replace(day=15)
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def find_latest_xlsx(directory: Path, keyword: str) -> Path | None:
    """在目录中找到包含关键字的最新xlsx文件"""
    matches = []
    for f in directory.glob("*.xlsx"):
        if keyword in f.name:
            matches.append(f)
    if not matches:
        return None
    return max(matches, key=lambda p: p.stat().st_mtime)


def main():
    parser = argparse.ArgumentParser(description="学情积分核算 - 一键运行")
    parser.add_argument("--start", help="开始日期 YYYY-MM-DD")
    parser.add_argument("--end", help="结束日期 YYYY-MM-DD")
    parser.add_argument("--auto", action="store_true", help="根据当前日期自动计算区间")
    parser.add_argument("--pool", help="历史课包池文件路径(可选，不传则仅用本期数据)")
    parser.add_argument("--pool-sheet", help="从指定Excel的sheet读取课包池")
    parser.add_argument("--pool-source", help="包含pool-sheet的Excel文件")
    parser.add_argument("--skip-fetch", action="store_true", help="跳过BI取数步骤")
    args = parser.parse_args()

    if args.auto:
        args.start, args.end = calc_date_range_auto()
        print(f"[自动模式] 计算区间: {args.start} ~ {args.end}")
    elif not args.start or not args.end:
        parser.error("请指定 --start 和 --end，或使用 --auto 自动计算")

    start_short = args.start.replace("-", "")
    end_short = args.end.replace("-", "")
    output_folder = OUTPUT_DIR / f"{start_short}-{end_short}学情积分发放明细"

    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    print("=" * 60)
    print(f"学情积分核算  {args.start} ~ {args.end}")
    print("=" * 60)

    # Step 1: BI 取数
    if not args.skip_fetch:
        print("\n>>> 步骤1: 从BI下载报表...")
        rc = subprocess.run(
            [sys.executable, str(SCRIPTS_DIR / "fetch_reports.py"),
             "--start", args.start, "--end", args.end],
            encoding="utf-8", errors="replace"
        ).returncode
        if rc != 0:
            print("[ERROR] BI取数失败，请检查日志")
            sys.exit(1)
    else:
        print("\n>>> 步骤1: 跳过BI取数 (--skip-fetch)")

    # Step 2: 定位下载的文件
    print("\n>>> 步骤2: 定位下载文件...")
    report1_file = find_latest_xlsx(EXPORTS_DIR, "续费规划")
    report2_file = find_latest_xlsx(EXPORTS_DIR, "上课明细")

    if not report1_file:
        report1_file = find_latest_xlsx(EXPORTS_DIR, "续费")
    if not report2_file:
        report2_file = find_latest_xlsx(EXPORTS_DIR, "学员")

    if not report1_file or not report2_file:
        print(f"[ERROR] 在 {EXPORTS_DIR} 中未找到报表文件")
        print(f"  报表1(续费规划表): {report1_file}")
        print(f"  报表2(上课明细): {report2_file}")
        print("\n请确认文件已下载，或手动指定路径后用 process_xueqing.py 处理")
        sys.exit(1)

    print(f"  报表1: {report1_file.name}")
    print(f"  报表2: {report2_file.name}")

    # Step 3: 数据处理
    print("\n>>> 步骤3: 数据处理...")
    process_cmd = [
        sys.executable, str(SCRIPTS_DIR / "process_xueqing.py"),
        "--report1", str(report1_file),
        "--report2", str(report2_file),
        "--output", str(output_folder),
    ]
    if args.pool:
        process_cmd += ["--pool", args.pool]
    elif args.pool_sheet:
        process_cmd += ["--pool-sheet", args.pool_sheet]
        if args.pool_source:
            process_cmd += ["--pool-source", args.pool_source]

    rc = subprocess.run(process_cmd, encoding="utf-8", errors="replace").returncode
    if rc != 0:
        print("[ERROR] 数据处理失败")
        sys.exit(1)

    # Step 4: 生成发放豌豆币模板（文件名带当天日期）
    print("\n>>> 步骤4: 生成发放豌豆币填写模板...")
    rc = subprocess.run(
        [sys.executable, str(SCRIPTS_DIR / "fill_wandou_template.py"),
         "--output-dir", str(output_folder)],
        encoding="utf-8", errors="replace",
    ).returncode
    if rc != 0:
        print("[WARN] 模板生成失败，可手动后续处理")

    print("\n" + "=" * 60)
    print(f"完成! 输出文件夹: {output_folder}")
    print("=" * 60)


if __name__ == "__main__":
    main()
