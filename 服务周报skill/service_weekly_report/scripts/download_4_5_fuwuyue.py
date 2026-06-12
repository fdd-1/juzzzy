"""单独下载 4.5 服务月跟进报表（思维转介绍过程跟进报表_末次渠道）
时间窗口：当月1号到上周日
"""
import sys
import json
import subprocess
from pathlib import Path
from datetime import datetime, timedelta

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
from _paths import PROJECT_ROOT, resolve_smartbi_cli_dir  # noqa: E402

SMARTBI_CLI_DIR = resolve_smartbi_cli_dir()


def calculate_dates():
    today = datetime.today()
    days_since_monday = (today.weekday() + 7) % 7 or 7
    last_monday = today - timedelta(days=days_since_monday + 7)
    last_sunday = last_monday + timedelta(days=6)
    month_start = datetime(today.year, today.month, 1)
    return month_start, last_sunday


def download_4_5_fuwuyue(output_dir: Path):
    """下载 4.5 服务月跟进报表"""
    month_start, last_sunday = calculate_dates()

    config = {
        "version": 1,
        "base_url": "https://bi.61info.cn/smartbi/vision",
        "tasks": {
            "service_weekly_4_5_fuwuyue": {
                "enabled": True,
                "description": "4.5 服务月跟进（当月1号到上周日）",
                "report": {
                    "id": "I2c9280870191447f447f3a940191b1cf2a202710",
                    "path": "分析报表/海外直播业务线/海外后端/思维-后端/转介绍/转介绍监控/思维转介绍过程跟进报表_末次渠道",
                    "type": "SPREADSHEET_REPORT"
                },
                "filters": {
                    "overrides": [
                        {"key": "做工开始时间_周维度", "value": month_start.strftime("%Y-%m-%d"), "displayValue": month_start.strftime("%Y-%m-%d")},
                        {"key": "结束日期", "value": last_sunday.strftime("%Y-%m-%d"), "displayValue": last_sunday.strftime("%Y-%m-%d")},
                        {"key": "日期", "value": last_sunday.strftime("%Y-%m-%d"), "displayValue": last_sunday.strftime("%Y-%m-%d")}
                    ]
                },
                "output": {
                    "type": "file",
                    "dir": str(output_dir / datetime.now().strftime("%Y-%m-%d") / "4_5_fuwuyue")
                }
            }
        }
    }

    config_path = Path(__file__).parent / "fuwuyue_4_5_config.json"
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

    print(f"配置已生成: {config_path}")
    print(f"时间窗口: {month_start.strftime('%Y-%m-%d')} ~ {last_sunday.strftime('%Y-%m-%d')}\n")

    cmd = [
        sys.executable,
        str(SMARTBI_CLI_DIR / "scripts" / "smartbi_cli.py"),
        "run",
        "--config", str(config_path),
        "--task", "service_weekly_4_5_fuwuyue",
        "--overwrite",
        "--json"
    ]

    result = subprocess.run(
        cmd,
        cwd=str(SMARTBI_CLI_DIR),
        capture_output=True,
        text=True,
        encoding="utf-8",
        errors="replace",
        timeout=180
    )

    if result.returncode == 0:
        try:
            data = json.loads(result.stdout)
            print("✓ 下载成功")
            print(f"  输出: {data.get('output')}")
            print(f"  字节数: {data.get('bytes')}")
        except json.JSONDecodeError:
            print(result.stdout)
    else:
        print(f"✗ 下载失败 (exit={result.returncode})")
        print(result.stderr)

    return result.returncode == 0


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", default="downloads/smartbi_reports", help="输出目录")
    args = parser.parse_args()

    output_dir = PROJECT_ROOT / args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    success = download_4_5_fuwuyue(output_dir)
    sys.exit(0 if success else 1)
