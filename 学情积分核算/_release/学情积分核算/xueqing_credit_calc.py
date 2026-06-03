#!/usr/bin/env python3
"""
学情积分核算 Skill

自动从BI提取海外思维学情课包数据，核算学员积分发放。
支持手动运行和定时任务（每月1号和16号自动触发）。

核算规则：
- 4条件发放：在课包池 + 是否预习=1 + 线上作业已提交 + 消耗课时>=1
- 积分公式：(基础课时消耗 + 赠送课时消耗) × 500

用法:
  python xueqing_credit_calc.py run --auto
  python xueqing_credit_calc.py run --start 2026-05-01 --end 2026-05-15
  python xueqing_credit_calc.py setup-schedule
"""

import argparse
import subprocess
import sys
from pathlib import Path
from datetime import datetime, date, timedelta
import openpyxl
from openpyxl.styles import Font

# 配置路径
SCRIPT_DIR = Path(__file__).parent.resolve()
BASE_DIR = SCRIPT_DIR
EXPORTS_DIR = BASE_DIR / "01_bi_exports"
OUTPUT_DIR = BASE_DIR / "03_output"

sys.stdout.reconfigure(encoding="utf-8")


# ============ 日期计算 ============

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


# ============ BI取数 ============

def run_bi_skill(args):
    """调用bi_skill下载报表（用户需自行配置 BI_SKILL_EXE 环境变量或修改路径）"""
    import os
    bi_skill_exe = Path(os.environ.get("BI_SKILL_EXE", r"C:\path\to\bi_skill.exe"))
    if not bi_skill_exe.exists():
        print(f"[ERROR] bi_skill.exe 不存在: {bi_skill_exe}")
        print("       请设置环境变量 BI_SKILL_EXE 指向你本机的 bi_skill 可执行文件")
        return 1
    result = subprocess.run([str(bi_skill_exe)] + args, encoding="utf-8", errors="replace")
    return result.returncode


def fetch_reports(start_date, end_date):
    """从BI下载两张报表"""
    print("\n>>> 步骤1: 从BI下载报表...")
    EXPORTS_DIR.mkdir(parents=True, exist_ok=True)

    print("  [1/2] 下载续费规划表...")
    rc1 = run_bi_skill([
        "search", "--profile-name", "海外思维续费规划表_新版_26年启用",
        "--extra-dates", f"当前课包签单年月开始={start_date},当前课包签单年月结束={end_date},当前课包签单时间开始={start_date},当前课包签单时间结束={end_date}",
        "--output", str(EXPORTS_DIR),
    ])
    if rc1 != 0:
        return 1

    print("  [2/2] 下载上课明细...")
    rc2 = run_bi_skill([
        "search", "--profile-name", "海外思维学员上课明细",
        "--extra-dates", f"开始日期={start_date},结束日期={end_date}",
        "--output", str(EXPORTS_DIR),
    ])
    if rc2 != 0:
        return 1

    print("  ✓ 报表下载完成")
    return 0


# ============ 数据处理 ============

def find_header_row(ws, keywords=("学员ID", "学生ID", "课包ID", "课时包ID"), max_scan=20, min_cols=15):
    """在前max_scan行中找到包含关键字的表头行"""
    for r in range(1, max_scan + 1):
        row_vals = [ws.cell(r, c).value for c in range(1, min(ws.max_column + 1, 100))]
        non_null_count = sum(1 for v in row_vals if v is not None)
        if non_null_count < min_cols:
            continue
        row_text = "|".join(str(v or "") for v in row_vals)
        if any(kw in row_text for kw in keywords):
            return r
    return 1


def read_sheet_as_dicts(ws, header_row=None):
    """读取工作表为字典列表"""
    if header_row is None:
        header_row = find_header_row(ws)
