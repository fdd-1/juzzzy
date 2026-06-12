"""单独下载 SOP 报表（修复筛选项）"""
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
    return last_monday, last_sunday

def download_sop(output_dir: Path):
    """下载 SOP 报表"""
    last_monday, last_sunday = calculate_dates()

    config = {
        "version": 1,
        "base_url": "https://bi.61info.cn/smartbi/vision",
        "tasks": {
            "service_weekly_4_1_sop": {
                "enabled": True,
                "description": "4.1 海外思维服务SOP执行情况",
                "report": {
                    "id": "I2c928087019364b764b704c8019375e98bea20d9",
                    "path": "分析报表/海外直播业务线/海外后端/思维-后端/服务/语义分析_服务/海外思维服务SOP执行情况",
                    "type": "SPREADSHEET_REPORT"
                },
                "filters": {
                    "overrides": [
                        {"key": "做工开始时间", "value": last_monday.strftime("%Y-%m-%d"), "displayValue": last_monday.strftime("%Y-%m-%d")},
                        {"key": "做工结束时间", "value": last_sunday.strftime("%Y-%m-%d"), "displayValue": last_sunday.strftime("%Y-%m-%d")},
                        {"key": "日期", "value": last_sunday.strftime("%Y-%m-%d"), "displayValue": last_sunday.strftime("%Y-%m-%d")},
                        {"key": "海外思维团队", "value": "", "displayValue": ""}
                    ]
                },
                "output": {
                    "type": "file",
                    "dir": str(output_dir / datetime.now().strftime("%Y-%m-%d") / "4_1_sop")
                }
            }
        }
    }

    # 保存配置
    config_path = Path(__file__).parent / "sop_config.json"
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config, f, indent=2, ensure_ascii=False)

    print(f"配置已生成: {config_path}")
    print(f"时间窗口: {last_monday.strftime('%Y-%m-%d')} ~ {last_sunday.strftime('%Y-%m-%d')}")
    print(f"筛选项: 美澳+港澳+台湾 三个教学服务区\n")

    # 执行下载
    cmd = [
        sys.executable,
        str(SMARTBI_CLI_DIR / "scripts" / "smartbi_cli.py"),
        "run",
        "--config", str(config_path),
        "--task", "service_weekly_4_1_sop",
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
            print(json.dumps(data, indent=2, ensure_ascii=False))
        except json.JSONDecodeError:
            print(result.stdout)
    else:
        print(f"✗ 下载失败 (exit={result.returncode})")
        print(result.stderr)

    return result.returncode == 0

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser()
    parser.add_argument("--output-dir", default="downloads/smartbi_reports_new", help="输出目录")
    args = parser.parse_args()

    output_dir = PROJECT_ROOT / args.output_dir
    output_dir.mkdir(parents=True, exist_ok=True)

    success = download_sop(output_dir)
    sys.exit(0 if success else 1)
