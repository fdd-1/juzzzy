# -*- coding: utf-8 -*-
"""
SKU复盘自动化 - 主入口脚本
下载BI报表 → 提取数据 → 匹配人群 → 聚合计算 → 生成报告

用法:
  python run_sku_review.py --region gangao --month 4 --year 2026
  python run_sku_review.py --region taiwan --start 2026-04-01 --end 2026-04-30
  python run_sku_review.py --region all --month 4 --year 2026
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
    BI_REPORT, REGION_FILTERS, ALL_REGIONS
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
    parts = start_date.split("-")
    return f"{int(parts[1])}月"


def find_file_in_data(pattern, data_dir):
    """在data目录中查找匹配的文件（排除Excel临时锁文件 ~$*）"""
    matches = [f for f in data_dir.glob(pattern)
               if not f.name.startswith("~$")]
    if matches:
        if len(matches) > 1:
            print(f"  ⚠ 命中多份 {pattern}：")
            for m in sorted(matches, key=lambda f: f.stat().st_mtime, reverse=True):
                ts = datetime.fromtimestamp(m.stat().st_mtime).strftime("%Y-%m-%d %H:%M")
                print(f"    - {m.name} ({ts})")
            print(f"  → 取最新")
        return max(matches, key=lambda f: f.stat().st_mtime)
    return None


def download_bi_report(start_date, end_date, output_dir, region_label, max_retries=3):
    """调用bi_skill下载BI报表（按区域命名）"""
    output_dir.mkdir(parents=True, exist_ok=True)

    for attempt in range(1, max_retries + 1):
        print(f"\n{'─'*50}")
        print(f"  下载: {BI_REPORT['profile_name']} - {region_label}")
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

    print(f"  ✗ 下载失败（{region_label}），请手动下载后放入 {output_dir}")
    return None


def run_one_region(region_key, start_date, end_date, month_label,
                   month_data_dir, output_root, skip_download):
    """跑单个区域。返回 (success: bool, error_msg: str|None)"""
    region_cfg = REGION_FILTERS[region_key]
    region_label = region_cfg["label"]
    filters = region_cfg["filters"]
    sku_keyword = region_cfg["sku_keyword"]

    print(f"\n{'='*60}")
    print(f"  区域: {region_label}  ({region_key})")
    print(f"{'='*60}")

    region_output_dir = output_root / region_label
    region_output_dir.mkdir(parents=True, exist_ok=True)

    # BI报表（按区域子目录隔离，避免文件互相覆盖）
    region_data_dir = month_data_dir / region_label
    region_data_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n[1/5] 下载BI报表...")
    if skip_download:
        bi_file = find_file_in_data("*主订单宽表*", region_data_dir)
        if not bi_file:
            bi_file = find_file_in_data(f"*主订单宽表*{region_label}*", month_data_dir)
        if not bi_file:
            # fallback: 月份根目录下的全量宽表（不分区域,通过filters筛选区域）
            bi_file = find_file_in_data("*主订单宽表*", month_data_dir)
        if not bi_file:
            return False, f"{region_label}: data目录中未找到主订单宽表文件"
        print(f"  使用已有文件: {bi_file.name}")
    else:
        bi_file = download_bi_report(start_date, end_date, region_data_dir, region_label)
        if not bi_file:
            return False, f"{region_label}: BI下载失败"

    # 正式池：所有区域共用同一份
    print(f"\n[2/5] 查找正式池和SKU测算文件...")
    pool_file = find_file_in_data("*正式池*", month_data_dir)
    if not pool_file:
        pool_file = find_file_in_data("*正式池*", DATA_DIR)
    if not pool_file:
        return False, f"{region_label}: 未找到正式池文件"
    print(f"  正式池: {pool_file.name}")

    # SKU测算：按区域关键字匹配
    sku_file = find_file_in_data(f"*{sku_keyword}*SKU复盘*", month_data_dir)
    if not sku_file:
        sku_file = find_file_in_data(f"*{sku_keyword}*SKU*", month_data_dir)
    print(f"  SKU测算: {sku_file.name if sku_file else '未找到（将跳过预算对标）'}")

    # 数据提取
    print(f"\n[3/5] 提取数据...")
    try:
        orders = extract_bi_data(bi_file, filters=filters)
    except Exception as e:
        return False, f"{region_label}: BI数据提取失败 - {e}"

    if not orders:
        return False, f"{region_label}: 筛选后订单为0，请检查筛选字段值是否变更"

    pool = extract_pool_data(pool_file)
    matched = match_orders_with_pool(orders, pool)

    if not matched:
        return False, f"{region_label}: 无匹配数据，请检查正式池月份是否对齐"

    # 聚合
    print(f"\n[4/5] 聚合计算...")
    node_results = aggregate_by_node(matched)
    pkg_results, cohort_totals = aggregate_by_package(matched)

    print(f"  按节点聚合: {len(node_results)} 个组合")
    print(f"  按套餐聚合: {len(pkg_results)} 个套餐")
    print(f"  一续: {cohort_totals.get('一续', 0)} 单")
    print(f"  多续: {cohort_totals.get('多续', 0)} 单")

    # 预算
    budget_data = {}
    pkg_budget_ratio = {}
    if sku_file:
        try:
            budget_data = extract_budget_from_sku(sku_file)
            pkg_budget_ratio = extract_package_budget_ratio(sku_file)
        except Exception as e:
            print(f"  ⚠ 预算数据提取失败: {e}")
            print(f"  → 本期 {region_label} 报告将不含对标，需人工确认是否补预算")

    # 报告
    print(f"\n[5/5] 生成报告...")
    region_month_label = f"{region_label}_{month_label}"
    generate_template_excel(
        node_results, budget_data, pkg_results, cohort_totals,
        pkg_budget_ratio, region_month_label, region_output_dir, filters
    )
    generate_html_report(
        node_results, budget_data, pkg_results, cohort_totals,
        pkg_budget_ratio, region_month_label, region_output_dir, filters
    )
    generate_csv_detail(pkg_results, region_month_label, region_output_dir)

    print(f"\n  ✓ {region_label} 完成 → {region_output_dir}")
    return True, None


def main():
    parser = argparse.ArgumentParser(description="SKU复盘自动化")
    parser.add_argument("--region", choices=ALL_REGIONS + ["all"],
                        default="gangao",
                        help="区域: gangao(港澳) / oumeiao(欧美澳) / taiwan(台湾) / all(全部)")
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

    start_date, end_date = determine_date_range(args)
    month_label = get_month_label(start_date)
    print(f"\n  日期范围: {start_date} ~ {end_date}")
    print(f"  月份标签: {month_label}")
    print(f"  区域: {args.region}")

    month_data_dir = DATA_DIR / month_label
    today = datetime.now().strftime("%Y%m%d")
    output_root = OUTPUT_DIR / f"{today}_{month_label}复盘"
    month_data_dir.mkdir(parents=True, exist_ok=True)
    output_root.mkdir(parents=True, exist_ok=True)

    regions = ALL_REGIONS if args.region == "all" else [args.region]

    results = []
    for region_key in regions:
        ok, err = run_one_region(
            region_key, start_date, end_date, month_label,
            month_data_dir, output_root, args.skip_download
        )
        results.append((region_key, ok, err))

    # 汇总
    print(f"\n{'='*60}")
    print(f"  执行汇总")
    print(f"{'='*60}")
    success = [r for r in results if r[1]]
    failed = [r for r in results if not r[1]]
    for region_key, ok, err in results:
        label = REGION_FILTERS[region_key]["label"]
        if ok:
            print(f"  ✓ {label}")
        else:
            print(f"  ✗ {label}: {err}")
    print(f"\n  成功: {len(success)} / 失败: {len(failed)}")
    print(f"  输出根目录: {output_root}")

    if failed:
        sys.exit(1)


if __name__ == "__main__":
    main()
