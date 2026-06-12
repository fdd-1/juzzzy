# CRM-SKU 快速上手

豌豆思维 CRM 批量创建课时包 + 商品套餐的工具。

## 三步上手

### 0. 配置凭据（首次）

不再把密码写进 `config.local.json`。两种方式二选一：

**方式 A：环境变量（推荐，CI/批量跑）**

```powershell
# PowerShell
$env:CRM_USERNAME = "xiongtingshi"
$env:CRM_PASSWORD = "你的密码"
```

```bash
# bash
export CRM_USERNAME=xiongtingshi
export CRM_PASSWORD=********
```

**方式 B：交互式输入（默认，本地手跑）**

不设环境变量，加 `--use-password` 跑时会提示输入；密码用 `getpass` 隐藏不回显。

### 1. 准备 Excel

按 [references/excel-schema.md](references/excel-schema.md) 准备配置表，模板：[templates/课包配置模板-美澳.xlsx](templates/课包配置模板-美澳.xlsx)

每行一个 SKU，横向 KV 格式：

```
课时包名称 | 【VIPTHINK】... | 课包类型 | 中课包 | 课包分类 | 小班直播课 | ... | 赠送礼品 | 海外特批16000豌豆币
```

### 2. 跑课时包

```bash
cd .claude/skills/CRM-SKU/scripts

# 首跑（含登录）
PYTHONIOENCODING=utf-8 python crm_batch_create_lesson_packages.py \
  --xlsx "<Excel 路径>" --limit 1 --use-password

# 全量
PYTHONIOENCODING=utf-8 python crm_batch_create_lesson_packages.py \
  --xlsx "<Excel 路径>" --skip-existing --use-password
```

### 3. 跑套餐

```bash
PYTHONIOENCODING=utf-8 python crm_batch_create_packages.py \
  --xlsx "<Excel 路径>" --use-password
```

## 输出

- 成功 / 失败日志：`scripts/logs/batch-<ts>.csv`、`scripts/logs/package-batch-<ts>.csv`
- 失败截图：`scripts/logs/submit-fail-*.png`、`scripts/logs/package-fail-*.png`
- 礼品验证截图：`scripts/logs/gift-checked-*.png`

## 出错了？

按这个顺序排查：

1. 看 CSV 日志「详情」列
2. 看对应失败截图
3. 查 [references/known-pitfalls.md](references/known-pitfalls.md)
4. UI 改版了：用 `record_user_actions_v2.py` 录制，对照 [references/workflow.md](references/workflow.md) §2 修复

## 文件结构

```
CRM-SKU/
├── SKILL.md                                # 完整说明
├── README.md                               # 本文件
├── scripts/
│   ├── crm_batch_create_lesson_packages.py # 课时包批量创建
│   ├── crm_batch_create_packages.py        # 套餐批量创建
│   ├── record_user_actions_v2.py           # 录制脚本
│   ├── config.example.json                 # 配置示例
│   └── utils/element_ui.py                 # Element UI 表单工具
├── references/
│   ├── excel-schema.md                     # Excel 配置表约定
│   ├── workflow.md                         # 完整工作流 + 录制方法
│   ├── element-ui-patterns.md              # Element UI 自动化模式
│   └── known-pitfalls.md                   # 已知坑 + 解决方案
└── templates/
    └── 课包配置模板-美澳.xlsx               # Excel 模板
```

## 命令速查

| 命令 | 用途 |
|---|---|
| `--xlsx <path>` | Excel 路径（必填）|
| `--start N` | 从第 N 条开始（断点续跑）|
| `--limit N` | 最多 N 条（先小批验证）|
| `--skip-existing` | 跳过已存在（重跑课包必加）|
| `--use-password` | 用账号密码登录（首跑用；密码取 env / 交互输入）|
| `--no-precheck` | 跳过执行前校验（不推荐）|

## 自动校验

- **执行前**：检查 Excel 必填列头、课包类型映射覆盖、整数字段、名称唯一性。不通过直接中止，不启动浏览器。
- **执行后**：每条创建成功后自动回列表页搜名称，CSV 日志新增「搜索校验」列；若 toast 报成功但搜不到，结果列标记为 `OK_BUT_NOT_FOUND`。

## 详细文档

- [SKILL.md](SKILL.md) - 完整说明、何时使用、坑与解决方案
- [references/excel-schema.md](references/excel-schema.md) - Excel 配置约定
- [references/workflow.md](references/workflow.md) - 完整流程 + UI 改版时的录制方法
- [references/element-ui-patterns.md](references/element-ui-patterns.md) - 各类 Element UI 控件的自动化模式
- [references/known-pitfalls.md](references/known-pitfalls.md) - 所有踩过的坑
