# -*- coding: utf-8 -*-
"""
服务绩效核算 - 默认配置
定义BI报表→指标→激励项的映射规则
"""

from pathlib import Path

# ─── 默认路径配置 ────────────────────────────────────────────────

SKILL_DIR = Path(__file__).parent

# BI 报表导出目录（按月存放在 data/ 下）
BI_REPORT_DIR = SKILL_DIR / "data"

# 参考激励文件目录
REFERENCE_DIR = SKILL_DIR / "reference"
REFERENCE_INCENTIVE_FILE = REFERENCE_DIR / "5月服务激励-模板.xlsx"

# 输出目录
OUTPUT_DIR = SKILL_DIR / "output"


# ─── BI报表→指标映射规则 ────────────────────────────────────────
# 定义如何从每张BI报表中提取关键指标

METRIC_SOURCES = {
    "首通及时跟进率": {
        "file_pattern": "*首通*",
        "sheet": "Sheet1",
        "locate_by": {
            "column_header": "及时跟进率",
            "row_match": {"B": "海外教学服务部", "D": "总计"},
            "header_row": 3,
        },
        "description": "益智海外新生首通监控 - 及时跟进率",
    },
    "首课及时跟进率": {
        "file_pattern": "*学管服务指标*",
        "sheet": "新建报表",
        "locate_by": {
            "column_header": "首课及时跟进率",
            "row_match": {"B": "海外团队", "C": "总计"},
            "header_row": 2,
        },
        "description": "海外思维学管服务指标 - 首课及时跟进率（海外团队/总计行）",
    },
    "首专及时跟进率": {
        "file_pattern": "*学管服务指标*",
        "sheet": "新建报表",
        "locate_by": {
            "column_header": "首专及时跟进率",
            "row_match": {"B": "海外团队", "C": "总计"},
            "header_row": 2,
        },
        "description": "海外思维学管服务指标 - 首专及时跟进率（海外团队/总计行）",
    },
    "语义点执行率加和": {
        "file_pattern": "*SOP执行*",
        "sheet": "汇总",
        "locate_by": {
            "column_header": "语义点执行率加和",  # AN列
            "row_match": {"B": "海外团队", "D": "总计"},
            "header_row": 3,
        },
        "description": "海外思维服务SOP执行 - 语义点执行率加和",
    },
    "外呼跟进率": {
        "file_pattern": "*停课学员*",
        "sheet": "Sheet1",
        "locate_by": {
            "column_header": "外呼跟进率",
            "row_match": {"B": "海外团队", "C": "总计"},
            "header_row": 6,                   # 列头在第6行
        },
        "description": "思维停课学员执行监控 - 外呼跟进率",
    },
}


# ─── 激励项→指标映射（用于自动填充） ─────────────────────────────
# 定义激励完成情况中的文本 → 对应指标名称
# 支持模糊匹配（如"首通及时跟进率"匹配"首通"）

INCENTIVE_ITEM_METRIC_MAP = {
    "首通及时跟进率": "首通及时跟进率",
    "首课及时跟进率": "首课及时跟进率",
    "首专及时跟进率": "首专及时跟进率",
    "语义点执行率加和": "语义点执行率加和",
    "语义点执行率加和得分": "语义点执行率加和",
    "停课唤醒目标学员外呼跟进率": "外呼跟进率",
    "外呼跟进率": "外呼跟进率",
}


# ─── 激励阈值配置（用于计算实际激励金额） ──────────────────────────
# 定义各激励项的完成率阈值
# 计算公式: 实际激励金额 = 激励总额 × min(完成情况 / 阈值, 1)

INCENTIVE_THRESHOLDS = {
    "首通": 0.9,      # 首通及时跟进率 - 目标90%
    "首课": 0.8,      # 首课及时跟进率 - 目标80%
    "首专": 0.75,     # 首专及时跟进率 - 目标75%
    "语义点": 2.0,    # 语义点执行率加和 - 目标2.0
    "外呼": 0.7,      # 外呼跟进率 - 目标70%
}


# ─── 激励项 Sheet 结构配置 ───────────────────────────────────────
# 定义激励项 sheet 中各字段的列位置（基于参考文件分析）

INCENTIVE_ITEM_SHEET = {
    "name": "激励项",
    "header_row": 2,          # 列头所在行
    "columns": {
        "激励方向": "A",
        "激励人群": "B",
        "小组": "C",
        "TL": "D",
        "工号": "E",
        "激励完成情况": "F",   # 需要填入指标值的列
        "激励门槛": "G",
        "激励内容": "H",
        "激励总额": "I",
        "实际激励金额": "J",   # 含公式，保留
    },
}


# ─── 激励汇总 Sheet 结构配置 ─────────────────────────────────────

INCENTIVE_SUMMARY_SHEET = {
    "name": "激励汇总",
    "header_row": 2,
}


# ─── 激励方案 Sheet 结构配置 ─────────────────────────────────────

INCENTIVE_PLAN_SHEET = {
    "name": "激励方案",
    "header_row": 1,
}
