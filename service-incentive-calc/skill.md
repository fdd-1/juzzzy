---
name: service-incentive-calc
version: 1.0.0
description: 服务绩效核算自动化 - 从BI报表提取指标，自动填入激励模板，输出服务激励核算Excel
category: data-processing
tags: [excel, bi-report, incentive, automation]
author: AI
---

# Service Incentive Calculation Skill

## 功能
自动从4张BI服务绩效报表提取关键指标，填入激励核算模板，输出完整的服务激励核算Excel文件。

## 使用前提
1. 已导出当月4张BI报表到 `桌面/服务绩效核算/` 文件夹
2. 已有参考激励文件（默认 `桌面/服务激励/4月/4月服务激励.xlsx`）

## 使用方式

### 命令行运行
```bash
python build_incentive.py [--month 5月]
```

### Python 调用
```python
from build_incentive import build_incentive_excel

output = build_incentive_excel(
    bi_dir=Path(".../服务绩效核算"),
    reference_file=Path(".../4月服务激励.xlsx"),
    output_dir=Path(".../服务激励"),
    month="5月"
)
print(f"输出: {output}")
```

## 数据流
```
BI报表 → 提取指标 → 填入激励模板 → 输出核算Excel
```

## 提取指标映射
| 指标 | 来源文件 | Sheet | 目标行 | 列 |
|------|---------|-------|--------|-----|
| 首通及时跟进率 | *首通* | Sheet1 | B=海外教学服务部, D=总计 | U |
| 首课及时跟进率 | *学管服务指标* | 新建报表 | B=海外团队 (数据行+1) | Y |
| 首专及时跟进率 | *学管服务指标* | 新建报表 | B=海外团队 (数据行+1) | AG |
| 语义点执行率加和 | *SOP执行* | 汇总 | B=海外团队, D=总计 | AN |
| 外呼跟进率 | *停课学员* | Sheet1 | B=海外团队, C=总计 | Y |

## 激励计算逻辑
公式保留在Excel中，封顶逻辑：实际/目标 ≤1 按比率，>1 封顶100%：
```
=ROUND(总额*IF(实际/目标<=1, 实际/目标, 1), 2)
```

## 月度适配
激励方案变化时：
1. 提供新月份的参考激励文件
2. 如BI报表结构变化，更新 `config.py` 中的 `METRIC_SOURCES`
3. 重新运行即可

## 文件结构
```
service-incentive-calc/
├── SKILL.md           # 本说明
├── config.py          # 映射配置（可自定义）
├── extract_metrics.py # 指标提取器
└── build_incentive.py # 主构建流程
```

## 注意事项
- openpyxl 保存后公式不会自动计算，需用Excel打开后自动重算
- 若某指标提取失败，程序会继续执行并提示未匹配项
- 临时文件 (~$开头) 会被自动忽略
- 指标定位支持：列头匹配 + 行列交叉定位 + 行偏移
