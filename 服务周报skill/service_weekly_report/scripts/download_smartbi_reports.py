"""
服务周报自动化 - SmartBI 报表下载器

功能：
1. 自动计算时间窗口（当月1号到上周日 / 上周一到上周日）
2. 批量下载 12 个报表
3. 按筛选项重命名文件（首课/首专）
"""
import sys
import json
import subprocess
from pathlib import Path
from datetime import datetime, timedelta

# 让本脚本可独立运行：把项目根目录加入 sys.path 以便 import _paths
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _paths import resolve_smartbi_cli_dir  # noqa: E402

# SmartBI CLI 路径（环境变量解析，不写绝对路径）
SMARTBI_CLI_DIR = resolve_smartbi_cli_dir()
CONFIG_TEMPLATE = Path(__file__).parent.parent / "configs" / "service_weekly_smartbi_tasks.json"


def calculate_date_windows():
    """计算时间窗口

    返回:
        previous_week: (上周一, 上周日)
        month_to_last_sunday: (当月1号, 上周日)
    """
    today = datetime.today()

    # 上周一到上周日
    days_since_monday = (today.weekday() + 7) % 7 or 7
    last_monday = today - timedelta(days=days_since_monday + 7)
    last_sunday = last_monday + timedelta(days=6)

    # 当月1号到上周日
    month_start = datetime(today.year, today.month, 1)

    return {
        "previous_week": (last_monday, last_sunday),
        "month_to_last_sunday": (month_start, last_sunday),
        "last_sunday": last_sunday,
        "last_thursday": last_sunday - timedelta(days=3),
        "last_friday": last_sunday - timedelta(days=2),
    }


def generate_dynamic_config(base_config_path: Path, output_dir: Path, date_windows: dict, custom_output_dir: Path = None) -> Path:
    """生成动态配置文件（将 date_window 替换为 extra_params）

    Args:
        base_config_path: 模板配置文件路径
        output_dir: 配置文件输出目录
        date_windows: 时间窗口字典
        custom_output_dir: 自定义报表输出目录（默认为 smartbi-data-cli 的 outputs 目录）

    Returns:
        生成的配置文件路径
    """
    with open(base_config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)

    # 格式化函数
    def fmt_date(dt):
        return dt.strftime("%Y-%m-%d")

    prev_week_start, prev_week_end = date_windows["previous_week"]
    month_start, last_sunday = date_windows["month_to_last_sunday"]

    # 为每个任务设置正确的日期参数
    for task_name, task_config in config["tasks"].items():
        filters = task_config.get("filters", {})
        date_window = filters.get("date_window")

        # 初始化 extra_params
        if "extra_params" not in filters:
            filters["extra_params"] = []

        # 根据 date_window 添加日期参数
        if date_window == "previous_week":
            # 上周一到上周日
            filters["extra_params"].extend([
                {"key": "开始时间", "value": fmt_date(prev_week_start)},
                {"key": "结束时间", "value": fmt_date(prev_week_end)}
            ])

            # 特殊处理：4.1 首课和首专的 LP做工开始/结束时间
            if "shouke" in task_name:
                filters["extra_params"].extend([
                    {"key": "LP做工开始时间", "value": fmt_date(prev_week_start)},
                    {"key": "LP做工结束时间", "value": fmt_date(prev_week_end)}
                ])
            elif "shouzhuan" in task_name:
                filters["extra_params"].extend([
                    {"key": "LP做工开始时间", "value": fmt_date(prev_week_start)},
                    {"key": "LP做工结束时间", "value": fmt_date(prev_week_end)}
                ])

            # 4.1 SOP 和 4.5 SOP：用 overrides 设置参数（不用 extra_params）
            # 关键：海外思维团队 = 空字符串 表示全部数据
            if "sop" in task_name:
                # 4.5 sop 的 日期=上周一，4.1 sop 的 日期=上周日
                if "4_5" in task_name:
                    sop_date = fmt_date(prev_week_start)  # 上周一
                else:
                    sop_date = fmt_date(last_sunday)  # 上周日

                filters["overrides"] = [
                    {"key": "做工开始时间", "value": fmt_date(prev_week_start), "displayValue": fmt_date(prev_week_start)},
                    {"key": "做工结束时间", "value": fmt_date(prev_week_end), "displayValue": fmt_date(prev_week_end)},
                    {"key": "日期", "value": sop_date, "displayValue": sop_date},
                    {"key": "海外思维团队", "value": "", "displayValue": ""}
                ]
                # 清空 extra_params
                filters["extra_params"] = []

            # 4.1 首通：首次分配开始/结束时间
            if "shoutong" in task_name:
                filters["extra_params"] = [
                    {"key": "首次分配开始时间", "value": fmt_date(prev_week_start)},
                    {"key": "首次分配结束时间", "value": fmt_date(prev_week_end)}
                ]

        elif date_window == "month_to_last_sunday":
            # 当月1号到上周日
            # 特殊处理：4.5 转介绍报表使用特殊参数名
            if "4_5_fuwuyue" in task_name:
                filters["extra_params"].extend([
                    {"key": "做工开始时间_周维度", "value": fmt_date(month_start)},
                    {"key": "结束日期", "value": fmt_date(last_sunday)},
                    {"key": "日期", "value": fmt_date(last_sunday)}
                ])
            else:
                # 其他报表使用通用参数名
                filters["extra_params"].extend([
                    {"key": "开始时间", "value": fmt_date(month_start)},
                    {"key": "结束时间", "value": fmt_date(last_sunday)}
                ])

        elif date_window == "current_week_snapshot":
            # LP架构表：日期=上周日
            filters["extra_params"] = [
                {"key": "日期", "value": fmt_date(last_sunday)}
            ]

        # 删除 date_window（smartbi-data-cli 不识别自定义值）
        if "date_window" in filters:
            del filters["date_window"]

    # 如果指定了自定义输出目录，修改所有任务的输出路径
    if custom_output_dir:
        run_date = datetime.now().strftime("%Y-%m-%d")
        for task_name, task_config in config["tasks"].items():
            output = task_config.get("output", {})
            # 提取原始的子目录名（如 4_1_shoutong）
            original_dir = output.get("dir", "").split("/")[-1]
            # 修改为自定义路径
            output["dir"] = str(custom_output_dir / run_date / original_dir)

    # 保存动态配置
    output_config = output_dir / f"service_weekly_tasks_dynamic_{datetime.now().strftime('%Y%m%d')}.json"
    output_config.parent.mkdir(parents=True, exist_ok=True)

    with open(output_config, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

    print(f"✓ 动态配置已生成: {output_config}")
    return output_config


def download_reports(config_path: Path, tasks: list = None, max_workers: int = 3, dry_run: bool = False):
    """批量下载报表（顺序模式，避免 smartbi-data-cli 批量模式的编码问题）

    Args:
        config_path: 配置文件路径
        tasks: 要下载的任务列表（None=全部）
        max_workers: 并发数（顺序模式下忽略）
        dry_run: 是否仅测试
    """
    if tasks is None:
        # 默认下载所有任务（除了 LP架构表 和 4.5 SOP）
        # 4.5 SOP 与 4.1 SOP 共用同一份报表（同筛选项），处理时直接复用
        tasks = [
            "service_weekly_4_1_shoutong",
            "service_weekly_4_1_shouke",
            "service_weekly_4_1_shouzhuan",
            "service_weekly_4_1_sop",
            # "service_weekly_4_1_lp_arch",  # SIMPLE_REPORT，跳过
            "service_weekly_4_2",
            "service_weekly_4_3",
            "service_weekly_4_4",
            "service_weekly_4_5_fuwuyue",
            # "service_weekly_4_5_sop",  # 与 4.1 SOP 同源，处理时复用
            "service_weekly_4_6_waihu",
            "service_weekly_4_6_qiwei",
        ]

    print(f"\n{'=' * 60}")
    print(f"{'DRY RUN' if dry_run else '开始下载'} - 顺序下载报表")
    print(f"{'=' * 60}")
    print(f"任务数: {len(tasks)}")
    print(f"配置文件: {config_path}")
    print(f"{'=' * 60}\n")

    results = []
    success_count = 0
    fail_count = 0

    for i, task in enumerate(tasks, 1):
        print(f"\n[{i}/{len(tasks)}] {task}")
        print("-" * 60)

        cmd = [
            sys.executable,
            str(SMARTBI_CLI_DIR / "scripts" / "smartbi_cli.py"),
            "run",
            "--config", str(config_path),
            "--task", task,
            "--json"
        ]

        if dry_run:
            cmd.append("--dry-run")
        else:
            cmd.append("--overwrite")

        try:
            result = subprocess.run(
                cmd,
                cwd=str(SMARTBI_CLI_DIR),
                capture_output=True,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=300
            )

            if result.returncode == 0:
                try:
                    data = json.loads(result.stdout)
                    if data.get("status") in ("success", "dry_run"):
                        print(f"  ✓ 成功")
                        success_count += 1
                        results.append({"task": task, "status": "success", "data": data})
                    else:
                        print(f"  ⚠ 状态: {data.get('status')}")
                        results.append({"task": task, "status": "warning", "data": data})
                except json.JSONDecodeError:
                    print(f"  ⚠ 输出非 JSON: {result.stdout[:200]}")
                    results.append({"task": task, "status": "warning", "stdout": result.stdout[:500]})
            else:
                print(f"  ✗ 失败 (exit={result.returncode})")
                if result.stderr:
                    print(f"    {result.stderr[:300]}")
                fail_count += 1
                results.append({"task": task, "status": "failed", "stderr": result.stderr[:500]})

        except subprocess.TimeoutExpired:
            print(f"  ✗ 超时（>5分钟）")
            fail_count += 1
            results.append({"task": task, "status": "timeout"})
        except Exception as e:
            print(f"  ✗ 异常: {e}")
            fail_count += 1
            results.append({"task": task, "status": "error", "error": str(e)})

    # 汇总
    print(f"\n{'=' * 60}")
    print(f"下载完成")
    print(f"{'=' * 60}")
    print(f"  成功: {success_count}/{len(tasks)}")
    print(f"  失败: {fail_count}/{len(tasks)}")

    return {"success_count": success_count, "fail_count": fail_count, "results": results}


def rename_reports(output_base_dir: Path, date_windows: dict):
    """重命名报表文件（首课/首专需要区分）

    Args:
        output_base_dir: 输出根目录
        date_windows: 时间窗口字典
    """
    last_sunday = date_windows["last_sunday"]
    run_date = datetime.now().strftime("%Y-%m-%d")

    export_dir = output_base_dir / "service_weekly" / run_date

    # 首课报表重命名
    shouke_dir = export_dir / "4_1_shouke"
    if shouke_dir.exists():
        for file in shouke_dir.glob("*.xlsx"):
            new_name = file.parent / f"海外思维学管服务指标统计表_首课_{last_sunday.strftime('%Y%m%d')}.xlsx"
            file.rename(new_name)
            print(f"✓ 重命名: {file.name} → {new_name.name}")

    # 首专报表重命名
    shouzhuan_dir = export_dir / "4_1_shouzhuan"
    if shouzhuan_dir.exists():
        for file in shouzhuan_dir.glob("*.xlsx"):
            new_name = file.parent / f"海外思维学管服务指标统计表_首专_{last_sunday.strftime('%Y%m%d')}.xlsx"
            file.rename(new_name)
            print(f"✓ 重命名: {file.name} → {new_name.name}")


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="服务周报 SmartBI 报表下载器")
    parser.add_argument("--dry-run", action="store_true", help="仅测试，不实际下载")
    parser.add_argument("--max-workers", type=int, default=3, help="并发下载数（默认3）")
    parser.add_argument("--output-dir", type=str, help="自定义输出目录（默认为smartbi-data-cli的outputs目录）")
    args = parser.parse_args()

    # 1. 计算时间窗口
    print("\n=== 计算时间窗口 ===")
    date_windows = calculate_date_windows()
    prev_start, prev_end = date_windows["previous_week"]
    month_start, last_sunday = date_windows["month_to_last_sunday"]

    print(f"上周: {prev_start.strftime('%Y-%m-%d')} 到 {prev_end.strftime('%Y-%m-%d')}")
    print(f"当月到上周日: {month_start.strftime('%Y-%m-%d')} 到 {last_sunday.strftime('%Y-%m-%d')}")

    # 2. 生成动态配置
    print("\n=== 生成动态配置 ===")
    output_dir = Path(__file__).parent.parent / "configs" / "_dynamic"

    # 自定义输出目录
    custom_output = Path(args.output_dir) if args.output_dir else None
    if custom_output:
        print(f"自定义输出目录: {custom_output}")

    dynamic_config = generate_dynamic_config(CONFIG_TEMPLATE, output_dir, date_windows, custom_output)

    # 3. 下载报表
    result = download_reports(dynamic_config, max_workers=args.max_workers, dry_run=args.dry_run)

    if result and not args.dry_run:
        # 4. 重命名报表
        print("\n=== 重命名报表 ===")
        if custom_output:
            rename_reports(custom_output, date_windows)
        else:
            rename_reports(SMARTBI_CLI_DIR / "outputs" / "bi_exports", date_windows)

        print("\n✅ 全部完成！")
        final_output = custom_output / datetime.now().strftime('%Y-%m-%d') if custom_output else SMARTBI_CLI_DIR / 'outputs' / 'bi_exports' / 'service_weekly' / datetime.now().strftime('%Y-%m-%d')
        print(f"输出目录: {final_output}")
