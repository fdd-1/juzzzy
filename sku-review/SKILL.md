---
name: sku_review
description: 港澳SKU复盘自动化。从BI下载海外益智主订单宽表，结合正式池和SKU测算文档，计算实际ASP/单课时价/套餐占比，与预算对标，生成HTML分析报告和Excel对标表。
user-invocable: true
allowed-tools:
  - Bash(python *)
  - Bash(pip install *)
  - Bash(ls *)
  - Bash(mkdir *)
  - Read
  - Write
---

# SKU复盘自动化 Skill

每月自动化执行港澳SKU复盘：下载BI报表 → 提取数据 → 匹配人群 → 计算指标 → 对标分析 → 生成报告。

## 使用方式

用户通过 `/sku_review` 调用。

## 前置准备

将以下文件放入 `data/{月份}/` 目录：
- **正式池文件**：文件名包含"正式池"（如 `4月正式池-04.01-v1.xlsx`）
- **SKU测算文件**：文件名包含"SKU"（如 `港澳4月SKU复盘-0421.xlsx`）

BI报表会自动下载，也可手动放入（文件名包含"主订单宽表"）。

## 执行命令

```bash
# 自动下载BI + 分析（指定月份）
python "C:\Users\fengjianyi\Desktop\SKU复盘\run_sku_review.py" --month 4 --year 2026

# 自动下载BI + 分析（指定日期范围）
python "C:\Users\fengjianyi\Desktop\SKU复盘\run_sku_review.py" --start 2026-04-01 --end 2026-04-30

# 跳过BI下载，使用已有文件
python "C:\Users\fengjianyi\Desktop\SKU复盘\run_sku_review.py" --month 4 --year 2026 --skip-download
```

## 目录结构

```
SKU复盘/
├── SKILL.md                # 本文档
├── run_sku_review.py       # 主入口
├── config.py               # 配置（筛选条件、字段映射、分类规则）
├── extract_data.py         # 数据提取（读Excel、匹配人群）
├── analyze.py              # 分析（聚合计算、预算提取）
├── generate_report.py      # 报告生成（HTML + Excel + CSV）
├── data/                   # 输入数据（按月份）
│   └── 4月/
│       ├── 海外益智主订单宽表-4月.xlsx   ← BI报表
│       ├── 4月正式池-04.01-v1.xlsx      ← 正式池
│       └── 港澳4月SKU复盘-0421.xlsx    ← SKU测算
└── output/                 # 输出报告（按月份）
    └── 4月/
        ├── SKU复盘分析_4月.html         ← HTML报告（可分享）
        ├── SKU精细对标_4月.xlsx         ← Excel对标表
        └── 套餐明细_4月.csv            ← CSV详细数据
```

## 筛选条件

下载后在Python中按以下字段筛选（可在config.py中修改）：
- 订单支付时业绩归属人五级部门 = 港澳益智教学服务区
- 区域等级 = 港澳

## 分析逻辑

1. **人群分类**：正式池"当前课包顺序"=1为一续，>1为多续，=0排除
2. **套餐分类**：从套餐名称匹配关键词（升舱/早鸟/其余/全量限定/学情限定/SVIP/全量）
3. **指标计算**：
   - ASP = 总金额 / 订单数
   - 含积分单课时价 = 总金额 / 课时数（含积分）
   - 不含积分单课时价 = 总金额 / 课时数（不含积分）
   - 占比 = 套餐订单数 / 同人群总订单数
4. **对标**：实际指标 vs SKU测算文档中的预算指标

## 依赖

- Python 3
- openpyxl（已安装）
- bi_skill（BI报表下载，路径：`C:\Users\fengjianyi\.workbuddy\skills\bi_skill\bi_skill.py`）
