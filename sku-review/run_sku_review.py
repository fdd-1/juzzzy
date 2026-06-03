# -*- coding: utf-8 -*-
"""
SKU复盘自动化 - 主入口脚本
下载BI报表 → 提取数据 → 匹配人群 → 聚合计算 → 生成报告

用法:
  python run_sku_review.py --month 4 --year 2026
  python run_sku_review.py --start 2026-04-01 --end 2026-04-30
"""

import argparse
import subprocess
import sys
import os
from datetime import date, timedelta, datetime
from pathlib import Path
from calendar import monthrange

sys.stdout.reconfigure(encoding='utf-8')

SKILL_DIR = Path(__file__).parent
sys.path.insert(0, str(SKILL_DIR))

from config import (
    BI_SKILL_PATH, DATA_DIR, OUTPUT_DIR,
    BI_REPORT, FILTERS
)
from extract_data import extract_bi_data, extract_pool_data, match_orders_with_pool
from analyze import aggregate_by_node, aggregate_by_package, extract_budget_from_sku, extract_package_budget_ratio
from generate_report import (
    generate_html_report, generate_template_excel, generate_csv_detail
)


def determine_date_range(args):
    """根据参数确定日期范围"""
    if args.start and args.end:
        return args.start, args.end

    today = date.today()
    if args.month and args.year:
        year, month = args.year, args.month
    else:
        first_of_this_month = today.replace(day=1)
        last_month = first_of_this_month - timedelta(days=1)
        year, month = last_month.year, last_month.month

    start = f"{year}-{month:02d}-01"
    _, last_day = monthrange(year, month)
    end = f"{year}-{month:02d}-{last_day:02d}"
    return start, end


def get_month_label(start_date):
    """从开始日期提取月份标签"""
    parts = start_date.split("-")
    return f"{int(parts[1])}月"


def find_file_in_data(pattern, data_dir):
    """在data目录中查找匹配的文件（排除Excel临时锁文件 ~$*）"""
    import glob
    matches = [f for f in data_dir.glob(pattern)
               if not f.name.startswith("~$")]
    if matches:
        return max(matches, key=lambda f: f.stat().st_mtime)
    return None


def download_bi_report(start_date, end_date, output_dir, max_retries=3):
    """调用bi_skill下载BI报表"""
    output_dir.mkdir(parents=True, exist_ok=True)

    for attempt in range(1, max_retries + 1):
        print(f"\n{'─'*50}")
        print(f"  下载: {BI_REPORT['profile_name']}")
        print(f"  日期: {start_date} ~ {end_date}")
        if attempt > 1:
            print(f"  重试: {attempt}/{max_retries}")
        print(f"{'─'*50}")

        cmd = [
            sys.executable, str(BI_SKILL_PATH), "search",
            "--profile-name", BI_REPORT["profile_name"],
            "--start-date-field", BI_REPORT["start_date_field"],
            "--start-date", start_date,
            "--end-date-field", BI_REPORT["end_date_field"],
            "--end-date", end_date,
            "--output", str(output_dir),
        ]

        env = dict(os.environ)
        env["PYTHONIOENCODING"] = "utf-8"

        try:
            result = subprocess.run(
                cmd, capture_output=True, text=True,
                encoding="utf-8", timeout=300, env=env
            )
            output = result.stdout + result.stderr
            print(output[-500:] if len(output) > 500 else output)

            if result.returncode == 0:
                xlsx_files = list(output_dir.glob("*.xlsx"))
                if xlsx_files:
                    latest = max(xlsx_files, key=lambda f: f.stat().st_mtime)
                    print(f"  ✓ 下载成功: {latest.name}")
                    return latest
        except subprocess.TimeoutExpired:
            print(f"  ✗ 超时")
        except Exception as e:
            print(f"  ✗ 错误: {e}")

        if attempt < max_retries:
            import time
            time.sleep(5)

    print("  ✗ 下载失败，请手动下载后放入data目录")
    return None


def main():
    parser = argparse.ArgumentParser(description="SKU复盘自动化")
    parser.add_argument("--month", type=int, help="月份")
    parser.add_argument("--year", type=int, help="年份")
    parser.add_argument("--start", help="开始日期 YYYY-MM-DD")
    parser.add_argument("--end", help="结束日期 YYYY-MM-DD")
    parser.add_argument("--skip-download", action="store_true",
                        help="跳过BI下载，使用data目录中已有文件")
    args = parser.parse_args()

    print("=" * 60)
    print("  SKU复盘自动化")
    print("=" * 60)

    # 1. 确定日期范围
    start_date, end_date = determine_date_range(args)
    month_label = get_month_label(start_date)
    print(f"\n  日期范围: {start_date} ~ {end_date}")
    print(f"  月份标签: {month_label}")

    # 目录：输出文件夹按生成日期+复盘月份命名
    month_data_dir = DATA_DIR / month_label
    today = datetime.now().strftime("%Y%m%d")
    output_folder_name = f"{today}_{month_label}复盘"
    month_output_dir = OUTPUT_DIR / output_folder_name
    month_data_dir.mkdir(parents=True, exist_ok=True)
    month_output_dir.mkdir(parents=True, exist_ok=True)

    # 2. 下载BI报表
    print(f"\n[1/5] 下载BI报表...")
    if args.skip_download:
        bi_file = find_file_in_data("*主订单宽表*", month_data_dir)
        if not bi_file:
            bi_file = find_file_in_data("*主订单宽表*", DATA_DIR)
        if not bi_file:
            print("  ✗ data目录中未找到主订单宽表文件")
            print("  请将BI报表放入:", month_data_dir)
            return
        print(f"  使用已有文件: {bi_file.name}")
    else:
        bi_file = download_bi_report(start_date, end_date, month_data_dir)
        if not bi_file:
            return

    # 3. 查找正式池和SKU测算文件
    print(f"\n[2/5] 查找正式池和SKU测算文件...")
    pool_file = find_file_in_data("*正式池*", month_data_dir)
    if not pool_file:
        pool_file = find_file_in_data("*正式池*", DATA_DIR)
    if not pool_file:
        print("  ✗ 未找到正式池文件")
        print("  请将正式池文件放入:", month_data_dir)
        return
    print(f"  正式池: {pool_file.name}")

    sku_file = find_file_in_data("*SKU复盘*", month_data_dir)
    if not sku_file:
        sku_file = find_file_in_data("*SKU*测算*", month_data_dir)
    if not sku_file:
        sku_file = find_file_in_data("*SKU*", DATA_DIR)
    print(f"  SKU测算: {sku_file.name if sku_file else '未找到（将跳过预算对标）'}")

    # 4. 提取数据
    print(f"\n[3/5] 提取数据...")
    orders = extract_bi_data(bi_file, filters=FILTERS)
    pool = extract_pool_data(pool_file)
    matched = match_orders_with_pool(orders, pool)

    if not matched:
        print("  ✗ 无匹配数据，请检查文件")
        return

    # 5. 聚合计算
    print(f"\n[4/5] 聚合计算...")
    node_results = aggregate_by_node(matched)
    pkg_results, cohort_totals = aggregate_by_package(matched)

    print(f"  按节点聚合: {len(node_results)} 个组合 (含'综合'汇总)")
    print(f"  按套餐聚合: {len(pkg_results)} 个套餐")
    print(f"  一续: {cohort_totals.get('一续', 0)} 单")
    print(f"  多续: {cohort_totals.get('多续', 0)} 单")

    # 预算数据
    budget_data = {}
    pkg_budget_ratio = {}
    if sku_file:
        try:
            budget_data = extract_budget_from_sku(sku_file)
            pkg_budget_ratio = extract_package_budget_ratio(sku_file)
        except Exception as e:
            print(f"  ⚠ 预算数据提取失败: {e}")

    # 6. 生成报告
    print(f"\n[5/5] 生成报告...")
    generate_template_excel(
        node_results, budget_data, pkg_results, cohort_totals,
        pkg_budget_ratio, month_label, month_output_dir, FILTERS
    )
    generate_html_report(
        node_results, budget_data, pkg_results, cohort_totals,
        pkg_budget_ratio, month_label, month_output_dir, FILTERS
    )
    generate_csv_detail(pkg_results, month_label, month_output_dir)

    print(f"\n{'=' * 60}")
    print(f"  ✓ SKU复盘完成！")
    print(f"  输出目录: {month_output_dir}")
    print(f"{'=' * 60}")


if __name__ == "__main__":
    main()
