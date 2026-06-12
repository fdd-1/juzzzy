# SKU 复盘 references 索引

本目录存放 Skill 运行时**只读**参考资料：字段映射、筛选规则、分类口径、对标阈值。脚本不直接 import 这里的内容；用于人工核对、新人交接、规则变更追溯。

| 文件 | 用途 |
|---|---|
| `bi_report_fields.md` | 海外益智主订单宽表必备字段 + 表头变更应对 |
| `region_filters.md` | 港澳/欧美澳/台湾三区域的筛选条件、易错点 |
| `package_categories.md` | 套餐分类优先级（"全量限定" 必须早于 "全量"） |
| `cohort_rules.md` | 人群分类（一续/多续/排除）与正式池字段口径 |
| `benchmark_thresholds.md` | 红黄绿对标阈值与"无预算"处理 |
| `bi_skill_path.md` | bi_skill.py 跨电脑路径解析顺序与切换方式 |
| `sku_budget_layout.md` | SKU 测算文档表头结构（结构A/B 自适应解析规则、口径区分） |

修改这里的内容**不会**改变脚本行为。真正的规则在 `config.py` 的 `REGION_FILTERS` / `PACKAGE_CATEGORIES` / `COHORT_RULES` / `DEVIATION_THRESHOLDS`。两边要一起改。
