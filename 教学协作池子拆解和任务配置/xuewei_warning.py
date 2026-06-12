#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""教学协作 P0 学位预警自动化流程 - 顶层入口

一键执行：
  python xuewei_warning.py all --month 2026-06 --teacher-complete-time "2026-06-30 23:59:59"

单步执行：
  python xuewei_warning.py export --month 2026-06
  python xuewei_warning.py filter --input exports/学位预警_20260605/...xlsx
  python xuewei_warning.py tag --month 2026-06
  python xuewei_warning.py group --month 2026-06
  python xuewei_warning.py polaris --target-month 6 --teacher-complete-time "2026-06-30 23:59:59"
  python xuewei_warning.py sync --month 2026-06
"""
import sys, io, subprocess, argparse, datetime as dt
from pathlib import Path

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent


def log(m): print(m, flush=True)


def run_script(script_path, *args):
    """调用脚本，返回 returncode"""
    cmd = [sys.executable, str(script_path)] + list(args)
    log(f"[RUN] {' '.join(cmd)}")
    result = subprocess.run(cmd, cwd=str(SCRIPT_DIR), encoding="utf-8", errors="replace")
    if result.returncode != 0:
        log(f"[ERROR] 脚本返回 {result.returncode}")
    return result.returncode


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("command", choices=["all", "export", "filter", "tag", "group", "polaris", "sync"],
                    help="执行步骤：all（全流程）或单步")
    ap.add_argument("--month", help="目标月份 YYYY-MM（例：2026-06）")
    ap.add_argument("--input", help="filter 步骤的输入文件（导出的报表 xlsx）")
    ap.add_argument("--target-month", type=int, help="polaris 步骤的目标月份（数字 1-12）")
    ap.add_argument("--teacher-complete-time", help="polaris 步骤的教师任务完成时间（YYYY-MM-DD HH:mm:ss）")
    args = ap.parse_args()

    if args.command == "all":
        if not args.month:
            log("[ERROR] all 命令需要 --month 参数")
            sys.exit(1)
        if not args.teacher_complete_time:
            log("[ERROR] all 命令需要 --teacher-complete-time 参数")
            sys.exit(1)

        # 解析月份，计算 target_month
        try:
            y, m = map(int, args.month.split("-"))
            target_month = m
        except Exception:
            log(f"[ERROR] --month 格式错误：{args.month}")
            sys.exit(1)

        log("=" * 60)
        log("[全流程] 开始执行 P0 学位预警自动化")
        log(f"  月份：{args.month}")
        log(f"  教师任务完成时间：{args.teacher_complete_time}")
        log("=" * 60)

        # Step 1: export
        log("\n[步骤 1/6] 导出报表")
        rc = run_script(SCRIPT_DIR / "export.py", "--month", args.month)
        if rc != 0:
            log("[ABORT] 导出报表失败")
            sys.exit(rc)

        # Step 2: filter（需要找到刚导出的文件）
        log("\n[步骤 2/6] 二次筛选")
        # 自动找最新的导出文件
        today_tag = dt.date.today().strftime("%Y%m%d")
        export_dir = SCRIPT_DIR / "exports" / f"学位预警_{today_tag}"
        if not export_dir.exists():
            log(f"[ERROR] 找不到导出目录：{export_dir}")
            sys.exit(2)
        xlsx_files = list(export_dir.glob("*.xlsx"))
        if not xlsx_files:
            log(f"[ERROR] 导出目录中没有 xlsx 文件：{export_dir}")
            sys.exit(2)
        # 取修改时间最新的
        latest_xlsx = max(xlsx_files, key=lambda p: p.stat().st_mtime)
        log(f"  -> 使用导出文件：{latest_xlsx}")
        rc = run_script(SCRIPT_DIR / "filter.py", "--input", str(latest_xlsx), "--month", args.month)
        if rc != 0:
            log("[ABORT] 筛选失败")
            sys.exit(rc)

        # Step 3: tag
        log("\n[步骤 3/6] 创建标签")
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "create_tag.py", "--month", args.month)
        if rc != 0:
            log("[ABORT] 创建标签失败")
            sys.exit(rc)

        # Step 4: group
        log("\n[步骤 4/6] 创建用户群")
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "create_group.py", "--month", args.month)
        if rc != 0:
            log("[ABORT] 创建用户群失败")
            sys.exit(rc)

        # Step 5: polaris
        log("\n[步骤 5/6] 克隆北极星任务")
        rc = run_script(SCRIPT_DIR / "polaris_task" / "update_task.py",
                        "--target-month", str(target_month),
                        "--teacher-complete-time", args.teacher_complete_time)
        if rc != 0:
            log("[ABORT] 克隆北极星任务失败")
            sys.exit(rc)

        # Step 6: sync
        log("\n[步骤 6/6] 标签数据同步")
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "sync_tag_data.py", "--month", args.month)
        if rc != 0:
            log("[ABORT] 标签数据同步失败")
            sys.exit(rc)

        log("=" * 60)
        log("[全流程] ✅ P0 学位预警自动化执行完成")
        log("=" * 60)

    elif args.command == "export":
        if not args.month:
            log("[ERROR] export 需要 --month 参数")
            sys.exit(1)
        rc = run_script(SCRIPT_DIR / "export.py", "--month", args.month)
        sys.exit(rc)

    elif args.command == "filter":
        if not args.input:
            log("[ERROR] filter 需要 --input 参数")
            sys.exit(1)
        rc = run_script(SCRIPT_DIR / "filter.py", "--input", args.input)
        sys.exit(rc)

    elif args.command == "tag":
        if not args.month:
            log("[ERROR] tag 需要 --month 参数")
            sys.exit(1)
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "create_tag.py", "--month", args.month)
        sys.exit(rc)

    elif args.command == "group":
        if not args.month:
            log("[ERROR] group 需要 --month 参数")
            sys.exit(1)
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "create_group.py", "--month", args.month)
        sys.exit(rc)

    elif args.command == "polaris":
        if not args.target_month:
            log("[ERROR] polaris 需要 --target-month 参数")
            sys.exit(1)
        if not args.teacher_complete_time:
            log("[ERROR] polaris 需要 --teacher-complete-time 参数")
            sys.exit(1)
        rc = run_script(SCRIPT_DIR / "polaris_task" / "update_task.py",
                        "--target-month", str(args.target_month),
                        "--teacher-complete-time", args.teacher_complete_time)
        sys.exit(rc)

    elif args.command == "sync":
        if not args.month:
            log("[ERROR] sync 需要 --month 参数")
            sys.exit(1)
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "sync_tag_data.py", "--month", args.month)
        sys.exit(rc)


if __name__ == "__main__":
    main()
