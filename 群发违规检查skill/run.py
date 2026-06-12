#!/usr/bin/env python3
"""群发违规检查 - 一键运行: 下载 + 检查 + 播报."""
import subprocess
import sys
import os
from datetime import date
from pathlib import Path

try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


def main():
    skill_dir = Path(__file__).parent
    bi_skill_path = Path.home() / ".workbuddy" / "skills" / "bi_skill" / "bi_skill.py"

    today = date.today()
    month_start = today.replace(day=1)

    print("=" * 60)
    print("海外LP群发违规检查 - 自动化流程")
    print("=" * 60)
    print(f"日期: {today}")
    print(f"数据范围: {month_start} 至 {today}")
    print()

    detail_file = skill_dir / "海外LP群发违规明细.xlsx"
    baseline_file = skill_dir / "报备文本.xlsx"
    if not baseline_file.exists():
        # 兼容用户实际命名
        alt = skill_dir / "海外益智群发日历-LP（主管填写）.xlsx"
        if alt.exists():
            baseline_file = alt

    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"

    # [1] 下载明细
    if bi_skill_path.exists():
        print("[1/3] 下载BI报表 海外LP群发违规明细...")
        download_cmd = [
            sys.executable, str(bi_skill_path), "search",
            "--profile-name", "海外LP群发违规明细",
            "--start-date", month_start.strftime("%Y-%m-%d"),
            "--end-date", today.strftime("%Y-%m-%d"),
            "--output", str(skill_dir),
        ]
        result = subprocess.run(download_cmd, env=env)
        if result.returncode != 0:
            print("    ⚠ BI下载失败，将尝试使用本地已有文件")
        else:
            print("    ✓ 下载完成\n")
    else:
        print(f"⚠ 未找到 bi_skill: {bi_skill_path}")
        print("  跳过下载，使用本地已有文件\n")

    if not detail_file.exists():
        print(f"错误: 找不到明细文件 {detail_file}")
        print("     请手动从 BI 下载「海外LP群发违规明细」放到该路径")
        sys.exit(1)

    if not baseline_file.exists():
        print(f"错误: 找不到报备文本 {baseline_file}")
        print("     请把钉钉文档导出为 Excel:")
        print("     https://alidocs.dingtalk.com/i/nodes/mExel2BLV542xe22HEGnGE60Wgk9rpMq")
        print(f"     另存为 {baseline_file}")
        sys.exit(1)

    # [2] 检查
    print("[2/3] 相似度过滤 + 续费场景统计...")
    output_file = skill_dir / f"违规统计_{today.strftime('%Y%m%d')}.xlsx"
    check_cmd = [
        sys.executable, str(skill_dir / "check.py"),
        str(detail_file),
        "--baseline", str(baseline_file),
        "--output", str(output_file),
    ]
    result = subprocess.run(check_cmd, env=env)
    if result.returncode != 0:
        print("检查失败")
        sys.exit(1)

    # [3] 播报
    print("\n[3/3] 钉钉群播报...")
    broadcast_cmd = [sys.executable, str(skill_dir / "broadcast.py"), str(output_file)]
    result = subprocess.run(broadcast_cmd, env=env)
    if result.returncode != 0:
        print("播报失败")
        sys.exit(1)

    print("\n" + "=" * 60)
    print("✓ 全部完成")
    print(f"  统计文件: {output_file}")
    print("=" * 60)


if __name__ == "__main__":
    main()
