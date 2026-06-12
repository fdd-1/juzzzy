#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""教学协作 P1 服务池协作自动化流程 - 顶层入口

P1 月份语义说明：
  --month 传的是"做几月的 P1 任务"。P1 内部会把这个月份 +1 后再传给所有子脚本，
  这样：
    - 报表筛选：开课M计算时间 = (x+1)月1号；退费结束时间 = x月最后一天
    - 标签 / 用户群命名：用 (x+1) 月
    - 北极星任务模板月份：(x+1) 月

一键执行（做 6 月 P1 任务）：
  python fuwuchi_auto.py all --month 2026-06 --teacher-complete-time "2026-06-30 23:59:59"

单步执行：
  python fuwuchi_auto.py export --month 2026-06
  python fuwuchi_auto.py filter --input exports/服务池_20260607/...xlsx
  python fuwuchi_auto.py tag --month 2026-06
  python fuwuchi_auto.py group --month 2026-06
  python fuwuchi_auto.py polaris --target-month 6 --teacher-complete-time "2026-06-30 23:59:59"
  python fuwuchi_auto.py sync --month 2026-06
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


def shift_month_plus_one(month_str):
    """把 YYYY-MM 加 1 个月。P1 任务的内部月份语义。

    例：2026-06 -> 2026-07；2026-12 -> 2027-01。
    返回 (shifted_str, shifted_year, shifted_month_int)。
    """
    try:
        y, m = map(int, month_str.split("-"))
        if m < 1 or m > 12:
            raise ValueError
    except Exception:
        log(f"[ERROR] --month 格式错误，需 YYYY-MM：{month_str}")
        sys.exit(1)
    if m == 12:
        ny, nm = y + 1, 1
    else:
        ny, nm = y, m + 1
    return f"{ny:04d}-{nm:02d}", ny, nm


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

        # P1 语义：用户传"做几月 P1 任务"，内部要 +1 个月
        user_month = args.month
        internal_month, _, target_month = shift_month_plus_one(user_month)

        log("=" * 60)
        log("[全流程] 开始执行 P1 服务池协作自动化")
        log(f"  用户月份（做 X 月 P1 任务）：{user_month}")
        log(f"  内部月份（X+1，用于报表/命名/任务）：{internal_month}")
        log(f"  教师任务完成时间：{args.teacher_complete_time}")
        log("=" * 60)

        # Step 1: export（复用P0的export.py，传入内部月份 = X+1）
        log("\n[步骤 1/6] 导出报表")
        rc = run_script(SCRIPT_DIR / "export.py", "--month", internal_month)
        if rc != 0:
            log("[ABORT] 导出报表失败")
            sys.exit(rc)

        # Step 2: filter（使用P1专用的filter_p1.py）
        log("\n[步骤 2/6] 二次筛选（P1：不限制课时）")
        today_tag = dt.date.today().strftime("%Y%m%d")
        export_dir = SCRIPT_DIR / "exports" / f"学位预警_{today_tag}"
        if not export_dir.exists():
            log(f"[ERROR] 找不到导出目录：{export_dir}")
            sys.exit(2)
        xlsx_files = list(export_dir.glob("*.xlsx"))
        if not xlsx_files:
            log(f"[ERROR] 导出目录中没有 xlsx 文件:{export_dir}")
            sys.exit(2)
        latest_xlsx = max(xlsx_files, key=lambda p: p.stat().st_mtime)
        log(f"  -> 使用导出文件：{latest_xlsx}")
        rc = run_script(SCRIPT_DIR / "filter_p1.py", "--input", str(latest_xlsx), "--month", internal_month)
        if rc != 0:
            log("[ABORT] 筛选失败")
            sys.exit(rc)

        # Step 3: tag（命名用 X+1 月）
        log("\n[步骤 3/6] 创建标签（P1服务池）")
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "create_tag_p1.py", "--month", internal_month)
        if rc != 0:
            log("[ABORT] 创建标签失败")
            sys.exit(rc)

        # Step 4: group（命名用 X+1 月）
        log("\n[步骤 4/6] 创建用户群（P1服务池）")
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "create_group_p1.py", "--month", internal_month)
        if rc != 0:
            log("[ABORT] 创建用户群失败")
            sys.exit(rc)

        # Step 5: polaris（target-month 用 X+1）
        log("\n[步骤 5/6] 克隆北极星任务（P1服务池）")
        rc = run_script(SCRIPT_DIR / "polaris_task" / "update_task_p1.py",
                        "--target-month", str(target_month),
                        "--teacher-complete-time", args.teacher_complete_time)
        if rc != 0:
            log("[ABORT] 克隆北极星任务失败")
            sys.exit(rc)

        # Step 6: sync
        log("\n[步骤 6/6] 标签数据同步（P1服务池）")
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "sync_tag_data_p1.py", "--month", internal_month)
        if rc != 0:
            log("[ABORT] 标签数据同步失败")
            sys.exit(rc)

        log("=" * 60)
        log("[全流程] ✅ P1 服务池协作自动化执行完成")
        log("=" * 60)

    elif args.command == "export":
        if not args.month:
            log("[ERROR] export 需要 --month 参数")
            sys.exit(1)
        internal_month, _, _ = shift_month_plus_one(args.month)
        log(f"[INFO] P1 内部月份：{internal_month}（用户传入 {args.month}）")
        rc = run_script(SCRIPT_DIR / "export.py", "--month", internal_month)
        sys.exit(rc)

    elif args.command == "filter":
        if not args.input:
            log("[ERROR] filter 需要 --input 参数")
            sys.exit(1)
        rc = run_script(SCRIPT_DIR / "filter_p1.py", "--input", args.input)
        sys.exit(rc)

    elif args.command == "tag":
        if not args.month:
            log("[ERROR] tag 需要 --month 参数")
            sys.exit(1)
        internal_month, _, _ = shift_month_plus_one(args.month)
        log(f"[INFO] P1 内部月份：{internal_month}（用户传入 {args.month}）")
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "create_tag_p1.py", "--month", internal_month)
        sys.exit(rc)

    elif args.command == "group":
        if not args.month:
            log("[ERROR] group 需要 --month 参数")
            sys.exit(1)
        internal_month, _, _ = shift_month_plus_one(args.month)
        log(f"[INFO] P1 内部月份：{internal_month}（用户传入 {args.month}）")
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "create_group_p1.py", "--month", internal_month)
        sys.exit(rc)

    elif args.command == "polaris":
        if not args.target_month:
            log("[ERROR] polaris 需要 --target-month 参数")
            sys.exit(1)
        if not args.teacher_complete_time:
            log("[ERROR] polaris 需要 --teacher-complete-time 参数")
            sys.exit(1)
        rc = run_script(SCRIPT_DIR / "polaris_task" / "update_task_p1.py",
                        "--target-month", str(args.target_month),
                        "--teacher-complete-time", args.teacher_complete_time)
        sys.exit(rc)

    elif args.command == "sync":
        if not args.month:
            log("[ERROR] sync 需要 --month 参数")
            sys.exit(1)
        internal_month, _, _ = shift_month_plus_one(args.month)
        log(f"[INFO] P1 内部月份：{internal_month}（用户传入 {args.month}）")
        rc = run_script(SCRIPT_DIR / "liuyi_tag" / "sync_tag_data_p1.py", "--month", internal_month)
        sys.exit(rc)


if __name__ == "__main__":
    main()
