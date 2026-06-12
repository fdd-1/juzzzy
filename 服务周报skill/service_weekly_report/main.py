"""
服务周报自动化主流程

执行步骤：
1. 计算时间窗口
2. 下载 SmartBI 报表（11个，跳过LP架构表）
3. 数据整合与格式化
4. 创建飞书电子表格（8个）
5. 生成结论（参考文档格式）
6. 创建最终统一文档
7. 移动到目标文件夹
"""
import sys
import os
import json
import subprocess
from pathlib import Path
from datetime import datetime, timedelta

# 项目路径（基于本文件位置推算，不写绝对路径）
PROJECT_ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(PROJECT_ROOT))
sys.path.insert(0, str(PROJECT_ROOT / "modules"))

from _paths import resolve_smartbi_cli_dir, ensure_credentials  # noqa: E402

# 外部工具目录从环境变量解析
SMARTBI_CLI_DIR = resolve_smartbi_cli_dir()

# 凭据检查（绝不在代码里写明文凭据）
ensure_credentials()


def print_step(step_num: int, title: str):
    """打印步骤标题"""
    print(f"\n{'='*70}")
    print(f"步骤 {step_num}: {title}")
    print(f"{'='*70}\n")


def calculate_time_windows():
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
        "run_date": datetime.now().strftime("%Y-%m-%d"),
    }


def step1_download_reports(time_windows: dict, dry_run: bool = False) -> Path:
    """步骤1：下载 SmartBI 报表"""
    print_step(1, "下载 SmartBI 报表")

    prev_start = time_windows["last_monday"].strftime("%Y-%m-%d")
    prev_end = time_windows["last_sunday"].strftime("%Y-%m-%d")
    month_start = time_windows["month_start"].strftime("%Y-%m-%d")

    print(f"上周: {prev_start} ~ {prev_end}")
    print(f"当月到上周日: {month_start} ~ {prev_end}")
    print(f"模式: {'DRY RUN' if dry_run else '实际下载'}\n")

    # 输出目录：放到 service_weekly_report/downloads/ 下
    output_base = PROJECT_ROOT / "downloads" / "smartbi_reports"
    output_base.mkdir(parents=True, exist_ok=True)

    # 调用下载脚本
    cmd = [
        sys.executable,
        str(PROJECT_ROOT / "scripts" / "download_smartbi_reports.py"),
        "--max-workers", "3",
        "--output-dir", str(output_base)
    ]

    if dry_run:
        cmd.append("--dry-run")

    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", errors="replace")

    if result.returncode != 0:
        print(f"❌ 下载失败:\n{result.stderr}")
        return None

    print(result.stdout)

    # 返回输出目录
    output_dir = output_base / time_windows["run_date"]
    return output_dir


def step2_process_data(smartbi_output_dir: Path, time_windows: dict) -> Path:
    """步骤2：数据整合与格式化"""
    print_step(2, "数据整合与格式化")

    # TODO: 调用数据处理模块
    # 从 smartbi_output_dir 读取原始报表
    # 处理后输出到 exports/weekly_{date}/

    export_dir = PROJECT_ROOT / "exports" / f"weekly_{time_windows['last_monday'].strftime('%Y%m%d')}_{time_windows['last_sunday'].strftime('%Y%m%d')}"
    export_dir.mkdir(parents=True, exist_ok=True)

    print(f"✓ 输出目录: {export_dir}")
    print("⚠️ 数据处理模块待集成（使用现有的 process_all_sections.py）")

    return export_dir


def step3_create_feishu_sheets(export_dir: Path) -> dict:
    """步骤3：创建飞书电子表格"""
    print_step(3, "创建飞书电子表格")

    # TODO: 调用飞书表格生成模块

    print("⚠️ 飞书表格生成模块待集成（使用现有的 feishu_simple_builder.py）")

    return {}


def step4_generate_conclusions(export_dir: Path) -> dict:
    """步骤4：生成结论"""
    print_step(4, "生成结论（参考文档格式）")

    # TODO: 调用结论生成模块

    print("⚠️ 结论生成模块待集成（使用现有的 conclusion_generator_v2.py）")

    return {}


def step5_create_final_doc(sheet_ids: dict, callouts: dict, time_windows: dict) -> str:
    """步骤5：创建最终统一文档"""
    print_step(5, "创建最终统一文档")

    # TODO: 调用文档生成模块

    print("⚠️ 文档生成模块待集成（使用现有的 final_doc_builder_v3.py）")

    return ""


def main():
    """主流程"""
    import argparse

    parser = argparse.ArgumentParser(description="服务周报自动化主流程")
    parser.add_argument("--dry-run", action="store_true", help="仅测试下载，不实际执行")
    parser.add_argument("--skip-download", action="store_true", help="跳过下载步骤（使用已有数据）")
    args = parser.parse_args()

    print("\n" + "="*70)
    print("服务周报自动化 v1.0")
    print("="*70)

    # 计算时间窗口
    time_windows = calculate_time_windows()

    print(f"\n时间窗口:")
    print(f"  上周: {time_windows['last_monday'].strftime('%Y-%m-%d')} ~ {time_windows['last_sunday'].strftime('%Y-%m-%d')}")
    print(f"  当月到上周日: {time_windows['month_start'].strftime('%Y-%m-%d')} ~ {time_windows['last_sunday'].strftime('%Y-%m-%d')}")
    print(f"  运行日期: {time_windows['run_date']}")

    # 步骤1: 下载报表
    if not args.skip_download:
        smartbi_output = step1_download_reports(time_windows, dry_run=args.dry_run)
        if not smartbi_output:
            print("\n❌ 下载失败，流程终止")
            return 1
    else:
        print_step(1, "下载 SmartBI 报表 [跳过]")
        smartbi_output = PROJECT_ROOT / "downloads" / "smartbi_reports" / time_windows["run_date"]

    if args.dry_run:
        print("\n✅ Dry-run 完成！")
        print("\n后续步骤（实际执行时）：")
        print("  2. 数据整合与格式化")
        print("  3. 创建飞书电子表格")
        print("  4. 生成结论")
        print("  5. 创建最终统一文档")
        return 0

    # 步骤2: 数据处理
    export_dir = step2_process_data(smartbi_output, time_windows)

    # 步骤3: 飞书表格
    sheet_ids = step3_create_feishu_sheets(export_dir)

    # 步骤4: 生成结论
    callouts = step4_generate_conclusions(export_dir)

    # 步骤5: 最终文档
    doc_url = step5_create_final_doc(sheet_ids, callouts, time_windows)

    # 完成
    print("\n" + "="*70)
    print("✅ 服务周报自动化完成！")
    print("="*70)

    if doc_url:
        print(f"\n📄 最终文档: {doc_url}")

    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except KeyboardInterrupt:
        print("\n\n⚠️ 用户中断")
        sys.exit(130)
    except Exception as e:
        print(f"\n❌ 执行失败: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
