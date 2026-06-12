# -*- coding: utf-8 -*-
"""
服务激励核算 - 全自动月度执行脚本
下载BI报表 → 提取指标 → 计算激励 → 输出Excel

用法:
  python run_monthly.py                    # 计算上月（适用于每月1号定时任务）
  python run_monthly.py --month 5 --year 2026  # 指定月份
  python run_monthly.py --start 2026-05-01 --end 2026-05-27  # 指定日期范围
"""

import argparse
import os
import subprocess
import sys
import time
from datetime import date, datetime, timedelta
from pathlib import Path
from calendar import monthrange

SKILL_DIR = Path(__file__).parent
BI_SKILL_PATH = Path(r"C:\Users\fengjianyi\.workbuddy\skills\bi_skill\bi_skill.py")
DATA_DIR = SKILL_DIR / "data"
OUTPUT_DIR = SKILL_DIR / "output"
REFERENCE_DIR = SKILL_DIR / "reference"
ENV_FILE = SKILL_DIR / ".env"

# 指标合理值范围（与 reference/metric_rules.md 一致）
METRIC_VALUE_RANGES = {
    "首通及时跟进率": (0, 1),
    "首课及时跟进率": (0, 1),
    "首专及时跟进率": (0, 1),
    "语义点执行率加和": (0, 3),
    "外呼跟进率": (0, 1),
}
INCENTIVE_TOTAL_CAP = 2000


def load_dotenv(env_path: Path = ENV_FILE) -> None:
    """轻量加载 .env 到 os.environ，缺失则跳过（避免硬依赖 python-dotenv）。"""
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        k, v = k.strip(), v.strip().strip('"').strip("'")
        if k and k not in os.environ:
            os.environ[k] = v

REPORTS = [
    {
        "name": "海外思维学管服务指标统计表",
        "profile_name": "海外思维学管服务指标统计表",
        "type": "report",
        "date_args": lambda s, e: [
            "--start-date-field", "开始日期", "--start-date", s,
            "--end-date-field", "结束时间", "--end-date", e,
            "--extra-dates", f"LP做工开始时间={s},LP做工结束时间={e},结束时间={e} 23:59:00",
        ],
    },
    {
        "name": "海外思维服务SOP执行情况",
        "profile_name": "海外思维服务SOP执行情况",
        "type": "report",
        "date_args": lambda s, e: [
            "--extra-dates", f"日期={e},做工开始时间={s},做工结束时间={e}",
            "--filters", "海外思维团队=全部",
        ],
    },
    {
        "name": "益智海外新生首通监控",
        "profile_name": "益智海外新生首通监控",
        "type": "monitor",
        "date_args": lambda s, e: [
            "--start-date-field", "首次分配开始时间", "--start-date", s,
            "--end-date-field", "首次分配结束时间", "--end-date", e,
        ],
    },
    {
        "name": "思维停课学员执行监控",
        "search_name": "思维停课学员执行监控",
        "type": "monitor",
        "date_args": lambda s, e: [
            "--start-date-field", "开始时间", "--start-date", s,
            "--end-date-field", "结束时间", "--end-date", e,
        ],
    },
]


def determine_date_range(args) -> tuple:
    """根据参数确定日期范围。"""
    if args.start and args.end:
        return args.start, args.end

    today = date.today()
    if args.month and args.year:
        year, month = args.year, args.month
    else:
        # 默认：计算上月
        first_of_this_month = today.replace(day=1)
        last_month = first_of_this_month - timedelta(days=1)
        year, month = last_month.year, last_month.month

    start = f"{year}-{month:02d}-01"
    _, last_day = monthrange(year, month)
    end = f"{year}-{month:02d}-{last_day:02d}"
    return start, end


def get_month_label(start_date: str) -> str:
    """从开始日期提取月份标签，如 '5月'。"""
    # 处理带时间的日期格式 "YYYY-MM-DD HH:MM:SS"
    date_part = start_date.split()[0] if ' ' in start_date else start_date
    dt = datetime.strptime(date_part, "%Y-%m-%d")
    return f"{dt.month}月"


def download_reports(start_date: str, end_date: str, output_dir: Path, max_retries: int = 3) -> list:
    """下载4张BI报表到指定目录，失败时重试。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    results = []
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")

    for report in REPORTS:
        success = False
        for attempt in range(1, max_retries + 1):
            print(f"\n{'─'*50}")
            print(f"  下载: {report['name']}")
            print(f"  日期: {start_date} ~ {end_date}")
            if attempt > 1:
                print(f"  重试: {attempt}/{max_retries}")
            print(f"{'─'*50}")

            cmd = [sys.executable, str(BI_SKILL_PATH), "search"]

            if "profile_name" in report:
                cmd += ["--profile-name", report["profile_name"]]
            elif "search_name" in report:
                cmd += ["--name", report["search_name"]]

            cmd += report["date_args"](start_date, end_date)
            cmd += ["--output", str(output_dir)]

            env = dict(__import__("os").environ)
            env["PYTHONIOENCODING"] = "utf-8"

            try:
                result = subprocess.run(
                    cmd, capture_output=True, text=True, encoding="utf-8",
                    timeout=300, env=env
                )
                output = result.stdout + result.stderr
                if "下载完成" in output:
                    print(f"  ✓ 下载成功")
                    # 重命名文件加上时间戳
                    for file in output_dir.glob(f"{report['name']}*.xlsx"):
                        new_name = file.stem + f"_{timestamp}" + file.suffix
                        new_path = file.parent / new_name
                        file.rename(new_path)
                        print(f"  📁 文件: {new_name}")
                    results.append({"name": report["name"], "status": "success"})
                    success = True
                    break
                else:
                    print(f"  ✗ 下载可能失败")
                    if attempt < max_retries:
                        print(f"  ⏳ 等待后重试...")
                        time.sleep(5)
                    else:
                        print(f"    {output[-200:]}")
                        results.append({"name": report["name"], "status": "failed", "msg": output[-200:]})
            except subprocess.TimeoutExpired:
                print(f"  ✗ 超时")
                if attempt < max_retries:
                    print(f"  ⏳ 等待后重试...")
                    time.sleep(5)
                else:
                    results.append({"name": report["name"], "status": "timeout"})
            except Exception as e:
                print(f"  ✗ 异常: {e}")
                if attempt < max_retries:
                    print(f"  ⏳ 等待后重试...")
                    time.sleep(5)
                else:
                    results.append({"name": report["name"], "status": "error", "msg": str(e)})

        if not success:
            print(f"\n  ❌ {report['name']} 下载失败，无法继续")
            return results

    return results


def calculate_incentive(data_dir: Path, reference_file: Path, output_dir: Path, month: str) -> dict:
    """提取指标并计算激励金额。"""
    sys.path.insert(0, str(SKILL_DIR))
    from extract_metrics import extract_all_metrics
    from config import METRIC_SOURCES, INCENTIVE_THRESHOLDS

    print(f"\n{'═'*60}")
    print(f"  提取指标 & 计算激励 - {month}")
    print(f"{'═'*60}")

    # 提取指标
    metrics = extract_all_metrics(data_dir, METRIC_SOURCES)
    if not metrics:
        return {"error": "未能提取任何指标"}

    # 计算激励
    incentive_items = [
        ("首通及时跟进率", "首通", 200),
        ("首课及时跟进率", "首课", 200),
        ("首专及时跟进率", "首专", 400),
        ("语义点执行率加和", "语义点", 600),
        ("外呼跟进率", "外呼", 600),
    ]

    results = []
    total = 0
    for metric_name, threshold_key, amount in incentive_items:
        actual = metrics.get(metric_name)
        if actual is None:
            results.append({"name": metric_name, "status": "missing"})
            continue

        threshold = INCENTIVE_THRESHOLDS[threshold_key]
        ratio = min(actual / threshold, 1.0)
        incentive = round(amount * ratio, 2)
        total += incentive

        results.append({
            "name": metric_name,
            "actual": actual,
            "target": threshold,
            "ratio": ratio,
            "amount": amount,
            "incentive": incentive,
        })
        status = "封顶" if ratio >= 1.0 else f"{ratio*100:.1f}%"
        print(f"  {metric_name}: {actual:.4f} / {threshold} = {status} → {incentive}元")

    print(f"\n  总计: {total}元 / 2000元")

    # 生成输出Excel
    try:
        from build_incentive import build_incentive_excel
        month_dir = output_dir / month
        month_dir.mkdir(parents=True, exist_ok=True)
        output_file = build_incentive_excel(
            bi_dir=data_dir,
            reference_file=reference_file,
            output_dir=output_dir,
            month=month,
        )
        print(f"\n  ✓ Excel已生成: {output_file}")
    except Exception as e:
        print(f"\n  ⚠ Excel生成失败: {e}")
        output_file = None

    return {
        "month": month,
        "total": total,
        "items": results,
        "output_file": str(output_file) if output_file else None,
    }


def print_validation_checklist(result: dict) -> bool:
    """输出校验清单。返回是否全部通过。"""
    print(f"\n{'═'*60}")
    print(f"  📋 输出校验清单")
    print(f"{'═'*60}")

    checks = []
    items = result.get("items", [])
    extracted = [i for i in items if i.get("status") != "missing"]
    missing = [i for i in items if i.get("status") == "missing"]

    # 1. 5 个指标全部提取成功
    ok_extract = len(extracted) == 5 and len(missing) == 0
    checks.append(("5 个指标全部提取成功", ok_extract,
                   f"{len(extracted)}/5" + (f"（缺失: {[m['name'] for m in missing]}）" if missing else "")))

    # 2. 每个指标值在合理范围
    range_fails = []
    for it in extracted:
        name = it["name"]
        actual = it.get("actual")
        lo, hi = METRIC_VALUE_RANGES.get(name, (None, None))
        if lo is None or actual is None:
            continue
        if not (lo <= actual <= hi):
            range_fails.append(f"{name}={actual:.4f} 超出 [{lo},{hi}]")
    ok_range = not range_fails
    checks.append(("每个指标在合理范围", ok_range,
                   "全部正常" if ok_range else "; ".join(range_fails)))

    # 3. 激励总额 ≤ 2000
    total = result.get("total", 0)
    ok_cap = total <= INCENTIVE_TOTAL_CAP
    checks.append((f"激励总额 ≤ {INCENTIVE_TOTAL_CAP} 元", ok_cap, f"{total} 元"))

    # 4. 输出 Excel 已生成
    out_file = result.get("output_file")
    ok_file = bool(out_file) and Path(out_file).exists() if out_file else False
    checks.append(("输出 Excel 文件已生成", ok_file,
                   Path(out_file).name if ok_file else "未生成"))

    for desc, ok, detail in checks:
        mark = "✓" if ok else "✗"
        print(f"  [{mark}] {desc}: {detail}")

    all_pass = all(ok for _, ok, _ in checks)
    print(f"{'─'*60}")
    print(f"  结果: {'✅ 全部通过' if all_pass else '❌ 存在异常，需人工排查'}")
    print(f"{'═'*60}")
    return all_pass


def find_reference_file(month: str) -> Path:
    """查找参考激励文件（优先完整模板，其次方案文件）。"""
    # 优先找当月完整模板（含激励项、激励汇总sheet）
    template = REFERENCE_DIR / f"{month}服务激励-模板.xlsx"
    if template.exists():
        return template

    # 找任意最新的完整模板
    templates = sorted(REFERENCE_DIR.glob("*模板*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if templates:
        return templates[0]

    # 兜底：方案文件（只有激励方案sheet，Excel生成会跳过）
    plan_file = REFERENCE_DIR / f"{month}服务激励方案.xlsx"
    if plan_file.exists():
        return plan_file

    refs = sorted(REFERENCE_DIR.glob("*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if refs:
        return refs[0]

    raise FileNotFoundError(f"未找到参考激励文件，请在 {REFERENCE_DIR} 中放入模板")


def main():
    parser = argparse.ArgumentParser(description="服务激励核算自动化")
    parser.add_argument("--month", type=int, help="月份（如 5）")
    parser.add_argument("--year", type=int, help="年份（如 2026）")
    parser.add_argument("--start", default="", help="开始日期 YYYY-MM-DD")
    parser.add_argument("--end", default="", help="结束日期 YYYY-MM-DD")
    parser.add_argument("--skip-download", action="store_true", help="跳过下载，直接用已有报表计算")
    args = parser.parse_args()

    # 加载 .env（凭据），缺失也不阻塞，仅在下载时报错
    load_dotenv()
    if not args.skip_download and not (os.environ.get("SMARTBI_USERNAME") and os.environ.get("SMARTBI_PASSWORD")):
        print("❌ 缺少凭据：请在项目根目录创建 .env 并填写 SMARTBI_USERNAME / SMARTBI_PASSWORD")
        print("   或手动设置环境变量后重试。模板见 .env.example")
        return {"status": "failed", "reason": "missing credentials"}

    start_date, end_date = determine_date_range(args)
    month = get_month_label(start_date)

    print(f"{'═'*60}")
    print(f"  服务激励核算 - {month}")
    print(f"  日期范围: {start_date} ~ {end_date}")
    print(f"{'═'*60}")

    # 数据目录按月份隔离
    month_data_dir = DATA_DIR / month
    month_data_dir.mkdir(parents=True, exist_ok=True)

    # 步骤1: 下载报表
    if not args.skip_download:
        download_results = download_reports(start_date, end_date, month_data_dir)
        success_count = sum(1 for r in download_results if r["status"] == "success")
        print(f"\n  下载完成: {success_count}/4 成功")
        if success_count < 4:
            print(f"\n  ❌ 有报表下载失败，停止核算")
            return {"status": "failed", "reason": "报表下载失败"}

    # 步骤2: 计算激励
    reference_file = find_reference_file(month)
    print(f"\n  参考文件: {reference_file.name}")

    result = calculate_incentive(month_data_dir, reference_file, OUTPUT_DIR, month)

    # 步骤3: 输出汇总
    print(f"\n{'═'*60}")
    print(f"  核算完成")
    print(f"  李文韬 {month}服务激励: {result.get('total', 0)}元")
    if result.get("output_file"):
        print(f"  输出文件: {result['output_file']}")
    print(f"{'═'*60}")

    # 步骤4: 输出校验清单
    print_validation_checklist(result)

    return result


if __name__ == "__main__":
    main()
