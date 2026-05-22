# weekly-report-service

周报自动化 — 服务模块（4.x）Skill

把 BI 报表批量导出 → 数据整合 → 结论生成 → 格式化 Excel → 飞书电子表格嵌入文档。

## 触发场景

- 用户说"生成本周服务周报"、"跑一下 4.1 服务指标"、"周报自动化"
- 用户提供新一周的 BI 导出目录并要求出周报
- 已有目标飞书文档要求把表格嵌进去

## 设计要点（与上一版的关键差异）

- **不在飞书文档里写原生表格**：>200 行的原生表加载卡顿，超 2000 cell 直接报错。
- **统一走"飞书电子表格 + 文档嵌入"**：每张展示表是独立电子表格，文档里只嵌 `<sheet token>` 引用块。
- **本地必须先有格式化 Excel**：作为留档与人工核对依据，再上传飞书。
- **辅助列保留在整合宽表**：跟进率 / 接通率 / AI 占比仅给结论用，不进格式化 Excel。
- **不尝试把 Excel 复制粘贴到飞书文档**：飞书自研编辑器拦截所有程序化粘贴路径（execCommand / InputEvent / clipboard.write / OS Ctrl+V），不要在这条路上花时间。

## 标准流程（以 4.1 为例）

```powershell
# 1) BI 导出（沿用 bi_skill，按 _profiles_4_1.json）
#    产出: exports/4_1/<编号>_<别名>/*.xlsx

# 2) 整合宽表 + AI 单独成表
python weekly_report\consolidate_4_1_v2.py     # → _merged_4_1.xlsx（含辅助列）
python weekly_report\consolidate_4_1_ai.py     # → _merged_4_1_ai.xlsx + 4_1_AI学情_格式化.xlsx

# 3) 主表格式化 Excel
python weekly_report\export_4_1_excel.py       # → 4_1_格式化.xlsx

# 4) 结论文本（人工核对后贴 callout）
python weekly_report\conclusions_4_1.py

# 5) 创建飞书电子表格 + 嵌入文档
python weekly_report\create_feishu_sheets.py
```

## 文件结构

```
weekly_report/
├── consolidate_4_1_v2.py     # 6 份原始 → 主表整合（含辅助列）
├── consolidate_4_1_ai.py     # AI 学情单独成表 + 格式化 Excel
├── conclusions_4_1.py        # 结论文本（含 EXCLUDE_TEAMS = {"台湾组"}）
├── export_4_1_excel.py       # 主表格式化（多级表头 / 汇总行加色）
├── create_feishu_sheets.py   # 建表 + 分批写入 + 文档嵌入
├── _profiles_4_1.json        # BI 导出配置
└── _archive/                 # 已弃用的 XML 推送 / 粘贴脚本（仅留档）
```

## 关键工程细节（写新模块前必看）

- `lark-cli sheets +create` 返回 token 在 `data.spreadsheet_token`（不是 `data.spreadsheet.spreadsheet_token`）。
- `lark-cli sheets +info` 的 sheet 列表在 `data.sheets.sheets[0].sheet_id`（嵌套两层）。
- `lark-cli sheets +write --values` **不支持** `@file`，必须 inline JSON。
- Windows 命令行 ≤ 8 KB，`sheets +write` 时 `batch_size=3~5` 行最稳。
- `lark-cli docs +update --content @file` 路径**必须相对**，配合 `cwd=xml_file.parent`。
- 嵌入语法：`<h4>{标题}</h4><sheet token="{spreadsheet_token}"></sheet>`。
- `_merged_*.xlsx` 被 Excel 打开时 `pd.read_excel` 会抢锁失败，跑脚本前先关掉 Excel。
- `~$_merged_*.xlsx` 是 Excel 临时锁文件，可直接清理。

## 数据约定

- 团队顺序 `TEAM_ORDER`：港澳1组/港澳2组/港澳组/美澳1组/美澳2组/美澳3组/美澳4组/美澳5组/台湾组。
- 台湾组数据完整保留，**只在 conclusions 里通过 `EXCLUDE_TEAMS` 排除**。
- 整合宽表必含辅助列：`首通_跟进率`、`首通_接通率`、`首通_48小时企微绑定率`、`AI_首课/首专干预中占比`。
- AI 干预中占比统一按 `sum(干预中) / sum(任务总数)` 重算，不取均值。
- 月度目标在 `conclusions_*.py` 顶部 `TARGETS`：首通及时 95%、首课及时 85%、首专及时 85%。

## 扩展到 4.2 ~ 4.7

每个子模块四件套：

| 文件 | 职责 |
|------|------|
| `consolidate_4_X.py` | 整合宽表，含辅助列 |
| `conclusions_4_X.py` | 结论文本，支持 `EXCLUDE_TEAMS` |
| `export_4_X_excel.py` | 多级表头 / 汇总加色 / 数值格式化 |
| 复用 `create_feishu_sheets.py` | 公共函数：`create_sheet` / `write_data_in_batches` / `embed_sheet_in_doc`，子模块只配 Excel 路径 + 标题 |

业务字段变化只触达 `consolidate_*.py`；渲染层和上传层不动。

## 完整流程文档

详见仓库根目录 [`周报自动化-实施流程.md`](../周报自动化-实施流程.md)。
