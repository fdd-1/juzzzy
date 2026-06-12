#!/usr/bin/env python3
"""
学情积分核算 Skill

自动从BI提取海外思维学情课包数据，核算学员积分发放。
支持手动运行和定时任务（每月1号和16号自动触发）。

核算规则：
- 4条件发放：在课包池 + 是否预习=1 + 线上作业已提交 + 消耗课时>=1
- 积分公式：(基础课时消耗 + 赠送课时消耗) × 500

用法:
  python xueqing_credit_skill.py run --auto
  python xueqing_credit_skill.py run --start 2026-05-01 --end 2026-05-15
  python xueqing_credit_skill.py setup-schedule
"""

import argparse
import subprocess
import sys
from pathlib import Path
from datetime import date, timedelta

# 配置路径
SCRIPT_DIR = Path(__file__).parent.resolve()
BASE_DIR = SCRIPT_DIR
SCRIPTS_DIR = BASE_DIR / "scripts"
EXPORTS_DIR = BASE_DIR / "01_bi_exports"
OUTPUT_DIR = BASE_DIR / "03_output"

sys.stdout.reconfigure(encoding="utf-8")


def calc_date_range_auto():
    """根据当前日期自动计算统计区间"""
    today = date.today()
    if today.day <= 15:
        # 1号触发: 上月16号 ~ 上月最后一天
        last_month_last = today.replace(day=1) - timedelta(days=1)
        start = last_month_last.replace(day=16)
        end = last_month_last
    else:
        # 16号触发: 本月1号 ~ 本月15号
        start = today.replace(day=1)
        end = today.replace(day=15)
    return start.strftime("%Y-%m-%d"), end.strftime("%Y-%m-%d")


def cmd_run(args):
    """运行学情积分核算（调用现有的run.py）"""
    # 构建参数列表
    run_args = [sys.executable, str(SCRIPTS_DIR / "run.py")]

    if args.auto:
        run_args.append("--auto")
    else:
        if args.start:
            run_args.extend(["--start", args.start])
        if args.end:
            run_args.extend(["--end", args.end])

    if args.skip_fetch:
        run_args.append("--skip-fetch")

    if args.pool:
        run_args.extend(["--pool", args.pool])
    elif args.pool_sheet:
        run_args.extend(["--pool-sheet", args.pool_sheet])
        if args.pool_source:
            run_args.extend(["--pool-source", args.pool_source])
    else:
        # 默认使用上期输出的课包池
        latest_output = find_latest_output_folder()
        if latest_output:
            pool_file = latest_output / "积分汇总.xlsx"
            if pool_file.exists():
                run_args.extend(["--pool-source", str(pool_file)])
                run_args.extend(["--pool-sheet", "学情课包ID池"])
                print(f"[自动] 使用历史课包池: {pool_file}")

    # 执行 run.py（含取数 / 处理 / 填模板）
    rc = subprocess.run(run_args, encoding="utf-8", errors="replace").returncode
    if rc != 0:
        return rc

    # 可选：连同 OA 提交一起跑
    if getattr(args, "submit_oa", False):
        print("\n>>> 步骤5: 自动提交 OA 豌豆币添加申请...")
        return cmd_submit_oa(args)
    return rc


def find_latest_output_folder():
    """找到最新的输出文件夹"""
    if not OUTPUT_DIR.exists():
        return None
    folders = [f for f in OUTPUT_DIR.iterdir() if f.is_dir() and "学情积分发放明细" in f.name]
    if not folders:
        return None
    return max(folders, key=lambda p: p.stat().st_mtime)


def cmd_setup_schedule(args):
    """配置Windows定时任务"""
    ps_script = SCRIPTS_DIR / "setup_scheduled_task.ps1"

    if not ps_script.exists():
        print(f"[ERROR] 定时任务脚本不存在: {ps_script}")
        return 1

    # 更新PowerShell脚本中的路径，指向这个skill文件
    print("正在配置Windows定时任务...")
    print(f"  任务1: 每月1号 09:30 (计算上月16号~月底)")
    print(f"  任务2: 每月16号 09:30 (计算本月1号~15号)")
    print(f"  执行命令: python {SCRIPT_DIR / 'xueqing_credit_skill.py'} run --auto")

    # 执行PowerShell脚本
    result = subprocess.run(
        ["powershell.exe", "-ExecutionPolicy", "Bypass", "-File", str(ps_script)],
        encoding="utf-8", errors="replace"
    )

    if result.returncode == 0:
        print("\n✓ 定时任务配置成功")
        print("\n查看任务: schtasks /Query /TN \"学情积分核算_月初\"")
        print("手动触发: schtasks /Run /TN \"学情积分核算_月初\"")
        print("删除任务: schtasks /Delete /TN \"学情积分核算_月初\" /F")

    return result.returncode


def cmd_submit_oa(args):
    """提交 OA 豌豆币添加申请。

    前置条件：
      - scripts/oa_login/auth_state.json 已存在（QR 登录后产物）
      - 03_output/<本期>/发放豌豆币文档填写模板_YYYYMMDD.xlsx 已生成
        （submit_oa.py 会自动找最新期次的最新模板）
    """
    submit_script = SCRIPTS_DIR / "oa_login" / "submit_oa.py"
    if not submit_script.exists():
        print(f"[ERROR] 找不到 OA 提交脚本: {submit_script}")
        return 1

    auth_path = SCRIPTS_DIR / "oa_login" / "auth_state.json"
    if not auth_path.exists():
        print(f"[ERROR] 缺少登录态: {auth_path}")
        print("       请先运行: python scripts/oa_login/login_oa_qr.py 完成扫码登录")
        return 1

    # 找出本次会用到的期次目录和模板，给用户看一眼
    if getattr(args, "output_dir", None):
        period_dir = Path(args.output_dir)
    else:
        period_dir = find_latest_output_folder()
    template = None
    if period_dir and period_dir.exists():
        candidates = sorted(
            period_dir.glob("发放豌豆币文档填写模板*.xlsx"),
            key=lambda p: p.stat().st_mtime,
            reverse=True,
        )
        template = candidates[0] if candidates else None
    if getattr(args, "attachment", None):
        template = Path(args.attachment)

    print()
    print("=" * 60)
    print("即将自动提交 OA 豌豆币添加申请")
    print("=" * 60)
    print(f"  期次目录: {period_dir}")
    print(f"  附件:    {template.name if template else '(未找到，submit_oa 会再次探测)'}")
    print(f"  汇总:    {(period_dir / '积分汇总.xlsx') if period_dir else ''}")
    print()
    print("请打开附件核对学员 ID 与积分数无误，确认后输入 y 提交，其它键取消。")

    # --yes / -y 跳过确认（用于定时任务无人值守）
    if getattr(args, "yes", False):
        print("[--yes] 跳过确认，直接提交")
    else:
        try:
            ans = input("数据是否无误？(y/N): ").strip().lower()
        except EOFError:
            ans = ""
        if ans not in ("y", "yes"):
            print("[CANCEL] 已取消，未提交 OA。可手动核对后再跑：")
            print("         python xueqing_credit_skill.py submit-oa")
            return 0

    submit_args = [sys.executable, str(submit_script)]
    if getattr(args, "output_dir", None):
        submit_args.extend(["--output-dir", args.output_dir])
    if getattr(args, "attachment", None):
        submit_args.extend(["--attachment", args.attachment])

    print(f"正在提交 OA 豌豆币添加申请: {submit_script}")
    result = subprocess.run(submit_args, encoding="utf-8", errors="replace")
    return result.returncode


def main():
    parser = argparse.ArgumentParser(
        description="学情积分核算 Skill",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
示例:
  # 自动模式（根据当前日期计算区间）
  python xueqing_credit_skill.py run --auto

  # 一条龙：取数 → 核算 → 生成模板 → 自动提交 OA
  python xueqing_credit_skill.py run --auto --submit-oa

  # 手动指定日期
  python xueqing_credit_skill.py run --start 2026-05-01 --end 2026-05-15

  # 配置Windows定时任务
  python xueqing_credit_skill.py setup-schedule

  # 自动提交 OA 豌豆币添加申请（需先 run 生成模板，且已扫码登录）
  python xueqing_credit_skill.py submit-oa
        """
    )

    subparsers = parser.add_subparsers(dest="command", help="子命令")

    # run 命令
    run_parser = subparsers.add_parser("run", help="运行学情积分核算")
    run_parser.add_argument("--start", help="开始日期 YYYY-MM-DD")
    run_parser.add_argument("--end", help="结束日期 YYYY-MM-DD")
    run_parser.add_argument("--auto", action="store_true", help="根据当前日期自动计算区间")
    run_parser.add_argument("--pool", help="历史课包池文件路径")
    run_parser.add_argument("--pool-sheet", help="从指定Excel的sheet读取课包池")
    run_parser.add_argument("--pool-source", help="包含pool-sheet的Excel文件")
    run_parser.add_argument("--skip-fetch", action="store_true", help="跳过BI取数步骤")
    run_parser.add_argument("--submit-oa", action="store_true",
                            help="跑完核算后立即自动提交 OA 豌豆币添加申请（一条龙）")
    run_parser.add_argument("--output-dir", help="（仅 --submit-oa 时）指定附件所在期次目录")
    run_parser.add_argument("--attachment", help="（仅 --submit-oa 时）直接指定附件路径")
    run_parser.add_argument("--yes", "-y", action="store_true",
                            help="（仅 --submit-oa 时）跳过提交前确认（无人值守模式用）")

    # setup-schedule 命令
    schedule_parser = subparsers.add_parser("setup-schedule", help="配置Windows定时任务")

    # submit-oa 命令
    submit_parser = subparsers.add_parser(
        "submit-oa",
        help="自动提交 OA 豌豆币添加申请（需先生成模板）",
    )
    submit_parser.add_argument("--output-dir", help="期次目录，会在其中找最新发放模板")
    submit_parser.add_argument("--attachment", help="直接指定附件路径，覆盖自动探测")
    submit_parser.add_argument("--yes", "-y", action="store_true",
                               help="跳过提交前确认（无人值守模式用）")

    args = parser.parse_args()

    if not args.command:
        parser.print_help()
        return 1

    if args.command == "run":
        return cmd_run(args)
    elif args.command == "setup-schedule":
        return cmd_setup_schedule(args)
    elif args.command == "submit-oa":
        return cmd_submit_oa(args)
    else:
        parser.print_help()
        return 1


if __name__ == "__main__":
    sys.exit(main())

