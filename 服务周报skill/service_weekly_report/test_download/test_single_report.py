"""简化版测试：下载单个报表验证流程"""
import sys
import json
import subprocess
from pathlib import Path
from datetime import datetime, timedelta

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _paths import resolve_smartbi_cli_dir  # noqa: E402

SMARTBI_CLI_DIR = resolve_smartbi_cli_dir()
BASE_CONFIG = Path(__file__).parent.parent / "configs" / "service_weekly_smartbi_tasks.json"


def calculate_dates():
    """计算时间窗口"""
    today = datetime.today()
    days_since_monday = (today.weekday() + 7) % 7 or 7
    last_monday = today - timedelta(days=days_since_monday + 7)
    last_sunday = last_monday + timedelta(days=6)
    month_start = datetime(today.year, today.month, 1)

    return {
        "last_monday": last_monday,
        "last_sunday": last_sunday,
        "month_start": month_start,
    }


def generate_test_config(task_name: str) -> Path:
    """为单个任务生成测试配置（处理日期参数）"""
    with open(BASE_CONFIG, 'r', encoding='utf-8') as f:
        config = json.load(f)

    dates = calculate_dates()

    # 只保留测试的任务
    if task_name not in config["tasks"]:
        raise ValueError(f"任务 {task_name} 不存在")

    task_config = config["tasks"][task_name]
    config["tasks"] = {task_name: task_config}

    # 处理日期参数
    filters = task_config.get("filters", {})
    date_window = filters.get("date_window")

    if "extra_params" not in filters:
        filters["extra_params"] = []

    def fmt(dt):
        return dt.strftime("%Y-%m-%d")

    # 根据 date_window 添加日期
    if date_window == "previous_week":
        filters["extra_params"].extend([
            {"key": "开始时间", "value": fmt(dates["last_monday"])},
            {"key": "结束时间", "value": fmt(dates["last_sunday"])}
        ])
    elif date_window == "month_to_last_sunday":
        filters["extra_params"].extend([
            {"key": "开始时间", "value": fmt(dates["month_start"])},
            {"key": "结束时间", "value": fmt(dates["last_sunday"])}
        ])

    # 删除不支持的 date_window
    if "date_window" in filters:
        del filters["date_window"]

    # 保存测试配置
    test_config = Path(__file__).parent / f"test_config_{task_name}.json"
    with open(test_config, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

    return test_config


def test_single_report(task_name: str, dry_run: bool = True):
    """测试下载单个报表"""
    dates = calculate_dates()

    print(f"\n{'='*60}")
    print(f"测试任务: {task_name}")
    print(f"时间窗口: 上周 {dates['last_monday'].strftime('%Y-%m-%d')} ~ {dates['last_sunday'].strftime('%Y-%m-%d')}")
    print(f"         当月 {dates['month_start'].strftime('%Y-%m-%d')} ~ {dates['last_sunday'].strftime('%Y-%m-%d')}")
    print(f"模式: {'DRY RUN' if dry_run else '实际下载'}")
    print(f"{'='*60}\n")

    # 生成测试配置
    print("生成测试配置...")
    test_config = generate_test_config(task_name)
    print(f"✓ 配置文件: {test_config}\n")

    cmd = [
        sys.executable,
        str(SMARTBI_CLI_DIR / "scripts" / "smartbi_cli.py"),
        "run",
        "--config", str(test_config),
        "--task", task_name,
        "--json"
    ]

    if dry_run:
        cmd.append("--dry-run")
    else:
        cmd.append("--overwrite")

    print(f"执行命令: {' '.join(str(c) for c in cmd)}\n")

    result = subprocess.run(
        cmd,
        cwd=str(SMARTBI_CLI_DIR),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=180
    )

    print(f"返回码: {result.returncode}")

    if result.stdout:
        try:
            data = json.loads(result.stdout)
            print(json.dumps(data, indent=2, ensure_ascii=False))
        except json.JSONDecodeError:
            print(result.stdout)

    if result.stderr:
        print(f"\nSTDERR:\n{result.stderr}")

    return result.returncode == 0


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--task", default="service_weekly_4_3", help="任务名称")
    parser.add_argument("--execute", action="store_true", help="实际下载（默认dry-run）")
    args = parser.parse_args()

    success = test_single_report(args.task, dry_run=not args.execute)

    if success:
        print("\n✅ 测试成功")
    else:
        print("\n❌ 测试失败")
        sys.exit(1)
