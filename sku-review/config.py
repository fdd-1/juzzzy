# -*- coding: utf-8 -*-
"""SKU复盘自动化 - 配置文件"""

import os
from pathlib import Path

SKILL_DIR = Path(__file__).parent
DATA_DIR = SKILL_DIR / "data"
OUTPUT_DIR = SKILL_DIR / "output"
REFERENCES_DIR = SKILL_DIR / "references"


def _resolve_bi_skill_path() -> Path:
    """定位 bi_skill.py。优先级：环境变量 BI_SKILL_PATH → 用户主目录 .workbuddy → 桌面同级目录 → 相对路径。"""
    env_path = os.environ.get("BI_SKILL_PATH")
    if env_path:
        return Path(env_path).expanduser()

    candidates = [
        Path.home() / ".workbuddy" / "skills" / "bi_skill" / "bi_skill.py",
        SKILL_DIR.parent / "bi_skill" / "bi_skill.py",
        SKILL_DIR.parent / ".workbuddy" / "skills" / "bi_skill" / "bi_skill.py",
    ]
    for p in candidates:
        if p.exists():
            return p

    return candidates[0]


BI_SKILL_PATH = _resolve_bi_skill_path()

# BI报表配置
BI_REPORT = {
    "profile_name": "海外益智主订单宽表",
    "start_date_field": "开始日期",
    "end_date_field": "结束日期",
}

# 区域筛选配置
# 单值：字段=值（精确匹配）
# 排除值：字段_排除=[值1, 值2]（NOT IN）
REGION_FILTERS = {
    "gangao": {
        "label": "港澳",
        "sku_keyword": "港澳",
        "filters": {
            "订单支付时业绩归属人五级部门": "港澳益智教学服务区",
            "区域等级": "港澳",
        },
    },
    "oumeiao": {
        "label": "欧美澳",
        "sku_keyword": "欧美澳",
        "filters": {
            "订单支付时业绩归属人五级部门": "欧美澳益智教学服务区",
            "区域等级_排除": ["港澳", "台湾"],
        },
    },
    "taiwan": {
        "label": "台湾",
        "sku_keyword": "台湾",
        "filters": {
            "订单支付时业绩归属人五级部门": "台湾益智教学服务区",
            "区域等级": "台湾",
        },
    },
}

ALL_REGIONS = ["gangao", "oumeiao", "taiwan"]

# 默认筛选（向后兼容；显式传 --region 时被 REGION_FILTERS 覆盖）
FILTERS = REGION_FILTERS["gangao"]["filters"]

# 套餐分类规则（按优先级匹配，"全量限定"必须在"全量"前）
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

# 池子节点2 → SKU节点 映射规则
POOL_NODE2_TO_SKU = {
    "升舱": "升舱",
    "早鸟池": "早鸟",
}

# 人群分类规则
COHORT_RULES = {
    1: "一续",
    "gt1": "多续",
    0: None,
}

# 对标偏差阈值
DEVIATION_THRESHOLDS = {
    "yellow": 0.05,  # ±5% 黄
    "red": 0.15,     # ±15% 红
}
