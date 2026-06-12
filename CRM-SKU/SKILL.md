---
name: CRM-SKU
description: 在豌豆思维 CRM「财务 → 课时包管理 / 套餐管理」批量创建课时包和对应商品套餐。基于 Playwright 自动化 Element UI 表单，输入是 Excel SKU 配置表，输出是逐条 CSV 日志 + 失败截图。当用户要求"批量创建课时包"、"批量创建套餐"、"按 Excel 跑 SKU"、"配置 CRM 套餐"时触发。
---

# CRM-SKU 批量创建工具

豌豆思维 CRM 的 SKU（课时包 + 商品套餐）批量创建工具。一次配置 Excel，分两步自动化：先建课时包，再建对应套餐。

## 1. 何时使用

触发条件（任一）：

- 用户给出 Excel SKU 配置表（如 `美澳-6月.xlsx`），要求把课时包和套餐配置到 CRM
- 用户说"批量创建课时包"、"批量创建套餐"、"按 Excel 跑 SKU"
- 用户给出 `https://crm.vipthink.cn/.../ClassPackageManage` 或 `.../packageMge` 链接
- 跨月 / 不同区域（美澳、港澳、台湾）的 SKU 上线

不要用本 skill 处理：

- 编辑 / 删除已有 SKU（脚本只覆盖"添加"流程，删改是高风险操作）
- 课时包 / 套餐以外的 CRM 表单（学员、订单、退费等 → 用 `crm_operator` skill）

## 2. 项目位置

```
.claude/skills/CRM-SKU/
├── SKILL.md                                # 本文档
├── README.md                               # 快速上手
├── scripts/
│   ├── crm_batch_create_lesson_packages.py # 课时包批量创建
│   ├── crm_batch_create_packages.py        # 套餐批量创建
│   ├── record_user_actions_v2.py           # 录制脚本（适配新 UI 时用）
│   ├── config.example.json                 # 配置示例（不含敏感）
│   ├── config.local.json                   # 本地配置（gitignore）
│   ├── auth_state.json                     # 登录态（首跑后生成，gitignore）
│   ├── utils/
│   │   └── element_ui.py                   # Element UI 表单工具
│   └── logs/                               # CSV 日志 + 失败截图（gitignore）
├── references/
│   ├── excel-schema.md                     # Excel 配置表约定
│   ├── workflow.md                         # 完整工作流（先课包再套餐）
│   ├── element-ui-patterns.md              # Element UI 自动化模式（el-tree/el-cascader/el-switch）
│   └── known-pitfalls.md                   # 已知坑与解决方案
└── templates/
    └── 课包配置模板-美澳.xlsx               # Excel 输入模板
```

## 3. 完整流程（先课包再套餐）

### 第一步：批量创建课时包

```bash
cd .claude/skills/CRM-SKU/scripts
PYTHONIOENCODING=utf-8 python crm_batch_create_lesson_packages.py \
  --xlsx "<Excel 路径>" --use-password
```

参数：
- `--xlsx`：Excel 路径（必填）
- `--start N`：从第 N 条开始（断点续跑）
- `--limit N`：最多 N 条（小批验证）
- `--skip-existing`：跳过已存在（重跑必加）
- `--use-password`：用账号密码登录（首跑用；密码取自 `CRM_PASSWORD` 环境变量或交互式输入）
- `--no-precheck`：跳过执行前校验（不推荐，仅在确认无误时使用）

### 第二步：批量创建对应套餐

```bash
PYTHONIOENCODING=utf-8 python crm_batch_create_packages.py \
  --xlsx "<Excel 路径>" --use-password
```

套餐自动按 Excel 中的"课时包名称"作为套餐名，并将同名课时包加入。每个套餐固定配置：
- 套餐类型：正课套餐
- 服务协议：豌豆益智直播课合同（海外-不含教具）-20260403
- 使用入口：续费
- 是否换课：是 → 海外换课-0节
- 重复购买次数：开启 → 2
- 转介绍规则：全年课（A0B0）
- 赠送礼品：取自 Excel 的「赠送礼品」列（无则跳过）

如需修改默认配置，改 `crm_batch_create_packages.py` 中 `package_conf` 字典。

## 4. Excel 配置（横向 KV）

每行一个 SKU，行首必须是字符串 `课时包名称`：

```
课时包名称 | 【VIPTHINK】... | 课包类型 | 中课包 | 课包分类 | 小班直播课 | 有效期 | 9 | ...
```

字段映射（详见 [references/excel-schema.md](references/excel-schema.md)）：

| 列 | 必填 | 说明 |
|---|---|---|
| 课时包名称 | ✓ | 全局唯一，套餐与课包同名 |
| 课包类型 | ✓ | 中课包 / 年课包 / 季课包 / 两年课包 / 其他类型课包（"年课包pro" 自动映射为"年课包"） |
| 课包分类 | ✓ | 通常 `小班直播课` |
| 有效期 / 补课次数 / 普通课时 / 赠送课时 / 原价 / 优惠价 / 试学期 / 打卡次数 / 停课次数 | ✓ | 整数 |
| 适用课类 | ✓ | 顿号 / 中文逗号 / 英文逗号分隔 |
| 赠送礼品 | 可选 | 套餐用，如「海外特批16000豌豆币」；无填空或省略列 |

模板：[templates/课包配置模板-美澳.xlsx](templates/课包配置模板-美澳.xlsx)

## 5. 关键 UI 模式（Element UI）

详见 [references/element-ui-patterns.md](references/element-ui-patterns.md)。本 skill 已踩过的坑：

| 控件 | 关键技巧 |
|---|---|
| `el-cascader`（课包类型）| 点 `.el-cascader` 容器，等下拉，按 `.el-cascader-node` 文本匹配 |
| `el-autocomplete`（课时包搜索）| 输入名后点击 `常规正课 / 中课包 / xxx` 建议项 |
| 服务协议（搜索式）| 点 input → 输入「0403」过滤 → 点匹配项 |
| `el-tree` 复选框（赠送礼品）| 礼品名文本 → 向上找祖先 `.el-checkbox__inner` → 点击 → 验证 `.el-checkbox.is-checked` |
| `el-switch` + 隐藏 input（是否换课/重复购买次数）| JS 找文本元素，向上找开关；开后 input 可能 `aria-disabled=true` 但仍可填值（直接 `set value + dispatch input/change`）|
| 转介绍规则 | 实际选项是「全年课（A0B0）」，不是文档里的「其他 - A0B0」 |
| 礼品弹窗关闭 | 点主弹窗标题区或按 ESC（无确定按钮）|

## 6. 错误处理

| 错误 | 排查 |
|---|---|
| 课时包创建失败「找不到选项: 年课包pro」 | Excel 是 `年课包pro`，CRM 里只有「年课包」。解析时已自动映射，若新增类型按需扩展 `crm_batch_create_lesson_packages.py` 里的 `mapping` |
| 套餐创建失败「找不到礼品 checkbox」| 礼品名称与 CRM 里的不一致；先在 CRM 中确认礼品确实存在并搜得到 |
| 「请输入次数」校验未通过 | 重复购买次数的 input 没填中。脚本已绕过 `aria-disabled` 直接 set value，仍报错时看截图 |
| 弹窗未关闭 / 报错 | 看 `logs/package-fail-*.png` / `logs/package-error-*.png`，对照 [references/known-pitfalls.md](references/known-pitfalls.md) |
| UI 改版导致定位失败 | 用 `record_user_actions_v2.py` 重新录制，看 [references/workflow.md](references/workflow.md) 的录制章节 |

## 7. 安全 / 边界

- 默认 `headless=False`，肉眼监控；批量跑也保持有头模式
- 不实现"删除/编辑"流程，避免误操作
- `auth_state.json` / `config.local.json` / `logs/` 全部 gitignore
- 重跑课包前用 `--skip-existing`，重跑套餐前先在 CRM 列表确认已创建情况
- **凭据不写入配置文件**：密码读取顺序为 `环境变量 CRM_PASSWORD` → `交互式输入（getpass，不回显）`；账号可在 `config.local.json` 的 `auth.username` 设默认值或用 `CRM_USERNAME` 覆盖。首跑后 `auth_state.json` 自动维护登录态，后续不再需要密码。

## 7.1 执行前校验（自动）

每次跑批前，脚本会先做一次 Excel 体检：

- 列头覆盖：每行必须含 13 个必填 label（课时包名称、课包类型、课包分类、有效期、补课次数、普通课时、赠送课时、原价、优惠价、试学期、打卡次数、停课次数、适用课类）
- 课包类型映射：所有 Excel 出现过的「课包类型」都能映射到 `TYPE_PARENT` 中的合法叶子，未覆盖时直接报错并提示要补哪个映射
- 数值字段：9 个整数字段做 `int()` 转换检查
- 名称唯一性：同一份 Excel 内课时包名称不重复

不通过则中止，不会启动浏览器。可加 `--no-precheck` 跳过（不推荐）。

## 7.2 执行后验证（自动）

每条创建成功（toast 显示"保存成功"）之后，脚本回到列表页搜索这个名称，确认列表里能搜到。

- CSV 日志新增「搜索校验」列：`列表已找到` / `列表未找到` / 异常详情
- 若 toast 报成功但列表搜不到，结果列从 `OK` 改为 `OK_BUT_NOT_FOUND`，并在终端打 `[WARN]`

## 8. 完整使用示例（一次完整任务）

```bash
cd .claude/skills/CRM-SKU/scripts

# 1. 首跑：登录并验证第一个课包
PYTHONIOENCODING=utf-8 python crm_batch_create_lesson_packages.py \
  --xlsx "<Excel 路径>" --limit 1 --use-password

# 2. 看 logs/batch-*.csv，确认成功；失败先按 references/known-pitfalls.md 修

# 3. 全量课包（带去重）
PYTHONIOENCODING=utf-8 python crm_batch_create_lesson_packages.py \
  --xlsx "<Excel 路径>" --skip-existing --use-password

# 4. 全量套餐
PYTHONIOENCODING=utf-8 python crm_batch_create_packages.py \
  --xlsx "<Excel 路径>" --use-password

# 5. 失败重试（如果有）
python crm_batch_create_packages.py --xlsx "<Excel 路径>" --start <失败行号> --use-password
```

## 9. 参考文档

- [references/excel-schema.md](references/excel-schema.md) - Excel 配置表约定
- [references/workflow.md](references/workflow.md) - 完整工作流 + 录制方法
- [references/element-ui-patterns.md](references/element-ui-patterns.md) - Element UI 自动化模式
- [references/known-pitfalls.md](references/known-pitfalls.md) - 已知坑与解决方案
- [README.md](README.md) - 快速上手
