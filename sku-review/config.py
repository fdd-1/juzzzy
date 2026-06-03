# -*- coding: utf-8 -*-
"""SKU复盘自动化 - 配置文件"""

from pathlib import Path

SKILL_DIR = Path(__file__).parent
BI_SKILL_PATH = Path(r"C:\Users\fengjianyi\.workbuddy\skills\bi_skill\bi_skill.py")
DATA_DIR = SKILL_DIR / "data"
OUTPUT_DIR = SKILL_DIR / "output"
REFERENCE_DIR = SKILL_DIR / "reference"

# BI报表配置
BI_REPORT = {
    "profile_name": "海外益智主订单宽表",
    "start_date_field": "开始日期",
    "end_date_field": "结束日期",
}

# 筛选条件（下载后在Python中筛选）
FILTERS = {
    "订单支付时业绩归属人五级部门": "港澳益智教学服务区",
    "区域等级": "港澳",
}

# 套餐分类规则（按优先级匹配）
PACKAGE_CATEGORIES = [
    ("升舱", "升舱"),
    ("早鸟", "早鸟"),
    ("全量限定", "全量限定"),
    ("学情限定", "学情限定"),
    ("SVIP", "SVIP"),
    ("全量", "全量"),
    ("其余", "其余"),
]

# 正式池配置
POOL_SHEET_NAME = "池内（剔已续不可续）"

# 池子节点2 → SKU节点 映射规则（升舱=升舱, 早鸟池=早鸟, 其他=其余, 续池0=池外）
POOL_NODE2_TO_SKU = {
    "升舱": "升舱",
    "早鸟池": "早鸟",
    # 其他所有值（当月结课/次月结课/次次月结课/活跃低课时等）→ 其余
}

# 人群分类规则
COHORT_RULES = {
    1: "一续",      # 当前课包顺序=1
    "gt1": "多续",  # 当前课包顺序>1
    0: None,        # 池外，排除
}
