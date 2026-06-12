"""4.1 板块报表下载器
调用 bi_skill 下载 5 份报表到 exports/weekly_{开始}_{结束}/4_1/ 目录。
"""
from __future__ import annotations
import sys
import os
import json
import subprocess
from pathlib import Path
from datetime import date, datetime

# 强制UTF-8输出(Windows GBK问题)
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, ValueError):
        pass

sys.path.insert(0, str(Path(__file__).parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from date_utils import last_week_range, end_of_day, hours_before, fmt_date, month_start
from _paths import PROJECT_ROOT, resolve_bi_skill_path  # noqa: E402

BI_SKILL_PATH = str(resolve_bi_skill_path())


def _build_filter_args(filters_template: dict, dates: dict) -> list[str]:
    """根据日期参数和模板,构建 bi_skill search 的命令行参数列表。"""
    args = []

    # start_date / end_date
    if "start_date" in filters_template:
        v = filters_template["start_date"].format(**dates)
        args.extend(["--start-date", v])
    if "end_date" in filters_template:
        v = filters_template["end_date"].format(**dates)
        args.extend(["--end-date", v])

    # extra_dates
    extra = filters_template.get("extra_dates", [])
    if extra:
        parts = []
        for ed in extra:
            field = ed["field"]
            value = ed["value"].format(**dates)
            parts.append(f"{field}={value}")
        args.extend(["--extra-dates", ",".join(parts)])

    # filters (combobox)
    flt = filters_template.get("filters", [])
    if flt:
        parts = []
        for f in flt:
            parts.append(f"{f['field']}={f['value']}")
        args.extend(["--filters", ",".join(parts)])

    return args


def download_report(report_config: dict, dates: dict, output_dir: Path) -> Path | None:
    """下载单份报表,返回文件路径。"""
    output_dir.mkdir(parents=True, exist_ok=True)
    sub_dir = output_dir / report_config["id"]
    sub_dir.mkdir(parents=True, exist_ok=True)

    cmd = [
        sys.executable,
        BI_SKILL_PATH,
        "search",
        "--profile-name", report_config["name"],
        "--output", str(sub_dir),
    ]
    cmd.extend(_build_filter_args(report_config["filters_template"], dates))

    print(f"[下载] {report_config['name']} -> {sub_dir}")
    print(f"  cmd: {' '.join(cmd)}")

    try:
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=300)
        if result.returncode != 0:
            print(f"[错误] 下载失败 (exit={result.returncode})")
            print(f"  stderr: {result.stderr[:500]}")
            return None

        # 找到下载的xlsx文件
        xlsx_files = list(sub_dir.glob("*.xlsx"))
        if not xlsx_files:
            print(f"[错误] 未找到导出的xlsx文件")
            return None

        latest = max(xlsx_files, key=lambda p: p.stat().st_mtime)
        print(f"[完成] {latest}")
        return latest

    except subprocess.TimeoutExpired:
        print(f"[错误] 下载超时(>300s)")
        return None


def download_all(start_date: date, end_date: date, base_dir: Path) -> dict:
    """下载 4.1 所有报表,返回 {报表id: 文件路径} 映射。"""
    config_path = Path(__file__).parent.parent / "config" / "reports_4_1.json"
    config = json.loads(config_path.read_text(encoding="utf-8"))

    # 构造日期上下文
    last_friday = start_date.replace(day=start_date.day) + (end_date - start_date - (end_date - start_date))
    # 简化:直接用 weekday 计算
    from datetime import timedelta
    last_friday = end_date - timedelta(days=2)  # 上周日 - 2天 = 上周五
    last_thursday = end_date - timedelta(days=3)  # 上周日 - 3天 = 上周四

    dates = {
        "last_monday": fmt_date(start_date),
        "last_sunday": fmt_date(end_date),
        "last_friday": fmt_date(last_friday),
        "last_thursday": fmt_date(last_thursday),
        "month_start": fmt_date(month_start(end_date)),
    }

    output_dir = base_dir / f"weekly_{start_date.strftime('%Y%m%d')}_{end_date.strftime('%Y%m%d')}" / "4_1"
    output_dir.mkdir(parents=True, exist_ok=True)

    print(f"\n=== 4.1 板块报表下载 ===")
    print(f"日期范围: {dates['last_monday']} ~ {dates['last_sunday']}")
    print(f"输出目录: {output_dir}\n")

    results = {}
    for report in config["reports"]:
        path = download_report(report, dates, output_dir)
        results[report["id"]] = path

    return results


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--start-date", help="上周一 YYYY-MM-DD")
    parser.add_argument("--end-date", help="上周日 YYYY-MM-DD")
    parser.add_argument("--output", default=str(PROJECT_ROOT / "exports"), help="输出根目录")
    args = parser.parse_args()

    if args.start_date and args.end_date:
        start = datetime.strptime(args.start_date, "%Y-%m-%d").date()
        end = datetime.strptime(args.end_date, "%Y-%m-%d").date()
    else:
        start, end = last_week_range()

    results = download_all(start, end, Path(args.output))

    print(f"\n=== 下载结果 ===")
    for k, v in results.items():
        status = "✓" if v else "✗"
        print(f"  {status} {k}: {v}")
