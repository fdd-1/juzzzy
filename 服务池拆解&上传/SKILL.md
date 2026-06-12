---
name: service-pool-monthly
description: 海外益智服务池学员名单自动化生成与上传。每月初触发，从 BI 下载续费规划表，筛选并匹配学员流转数据，输出双 Sheet Excel，再上传六一工作台建标签 + 用户群，最后同步到豌豆数仓。当用户提到"服务池拆解"、"服务池学员名单"、"海外益智服务池"、"月初服务池上传六一"、"标签同步豌豆数仓"等场景时调用。
version: 1.1.0
---

# 海外益智服务池学员名单 Skill

## 触发场景

每月初执行，生成海外益智服务池学员名单并同步到六一工作台 + 豌豆数仓。典型用户表述：
- "跑下这个月的服务池"
- "把 X 月服务池学员上传到六一"
- "服务池标签同步数仓"

## 路径约定

本 Skill 使用以下变量，避免硬编码桌面绝对路径：

| 变量 | 默认值 | 说明 |
|------|--------|------|
| `$SKILL_DIR` | 当前 SKILL.md 所在目录 | 主入口、`config.json`、`sync_tag_data.py` 都在这里 |
| `$DESKTOP` | `%USERPROFILE%\Desktop`（Windows）/ `$HOME/Desktop` | 默认根目录 |
| `$POOL_DIR` | `$DESKTOP\每月池子` | 学员流转文件按月归档 |
| `$LIUYI_DIR` | `$DESKTOP\六一标签` | 六一标签工具 + `auth_state.json` |
| `$BI_TOOL_DIR` | `$DESKTOP\smartbi-data-cli-internal-*\scripts` | BI 浏览器导出脚本 |

调用时优先读 `config.json`；`config.json` 里的绝对路径仅作回退。下文示例统一用变量写法。

## 前置条件

- 当月学员流转文件已就绪：`$POOL_DIR\<M>月\学员流转定稿-终版v1.xlsx`
- 六一工作台登录态有效：`$LIUYI_DIR\auth_state.json`
- BI 账号未在其他浏览器占用

## 每月需要更新的配置

编辑 `$SKILL_DIR\config.json`，仅改这两处：

```json
{
  "current_month": "2026-07",
  "liuzhuang_file": "${POOL_DIR}\\7月\\学员流转定稿-终版v1.xlsx"
}
```

> `config.json` 支持 `${VAR}` 占位符（如 `${DESKTOP}`、`${POOL_DIR}`），运行时会从环境变量或默认值替换。也可继续写绝对路径，向下兼容。

## 完整自动化命令

### 一键全流程（推荐）

```bash
cd $SKILL_DIR
python service_pool_automation.py
```

### 分步执行

```bash
# Step 1-4：只跑数据处理，跳过六一工作台
python service_pool_automation.py --skip-liuyi

# Step 5：上传标签和用户群（需要 auth_state.json 有效）
python "$LIUYI_DIR\create.py" \
  --input "$SKILL_DIR\data\processed\dadou_ids_YYYYMM.xlsx" \
  --id-type wandou \
  --tag-name "【海外益智】26年X月服务池学员名单" \
  --group-name "【海外益智】26年X月服务池学员名单"

# Step 6：标签数据同步到豌豆数仓
cd $SKILL_DIR
python sync_tag_data.py --user-group-name "【海外益智】26年X月服务池学员名单"
```

## 7 步流程说明

| 步骤 | 内容 | 脚本 |
|------|------|------|
| 1 | BI 下载：海外思维续费规划表_新版_26年启用（筛选：开课M=当月1号，退费结束=上月末，池子节点3=服务月） | smartbi_browser_export.py |
| 2 | 数据筛选：是否可续学员=1，月初是否续费=空白 | service_pool_automation.py |
| 3 | 匹配：BI 大账号ID ↔ 学员流转 学员id → 最终归属LP、最终归属小组 | service_pool_automation.py |
| 4 | 输出 Excel（Sheet1 完整数据 + Sheet2 大账号ID） | service_pool_automation.py |
| 5 | 六一工作台新建标签 + 用户群 | 六一标签/create.py |
| 6 | 标签数据同步：新增 → 填写表单 → 确认 | sync_tag_data.py |
| 7 | 完成日志统计 | — |

## 输出文件

```
$SKILL_DIR\data\processed\
├── 服务池学员_YYYYMM.xlsx   # Sheet1: 完整数据含最终归属LP/小组
│                             # Sheet2: 大账号ID（六一标签数据）
└── dadou_ids_YYYYMM.xlsx    # 单独的大账号ID文件（create.py 用）
```

## 标签命名规则

`【海外益智】{年份}年{月份}月服务池学员名单`

例：`【海外益智】26年6月服务池学员名单`

## 边界处理与失败重试

### 自动重试的环节

| 步骤 | 重试策略 | 退出条件 |
|------|----------|----------|
| BI 下载（Step 1） | 默认 3 次，每次间隔 30s | 命中"账号在其他地方登录"立即停，提示用户关浏览器 |
| 学员流转读取（Step 3） | 不重试，直接报错 | 文件不存在 → 提示检查 `liuzhuang_file` 路径 |
| 六一标签上传（Step 5） | 2 次（auth_state 失效时不重试） | `code: 101` 直接停，提示检查 dadou_ids 列格式 |
| 数仓同步（Step 6） | 2 次，间隔 10s | 表单字段缺失立即停 |

### 边界场景

- **首次运行**：`data/`、`logs/` 目录会自动创建。
- **当月已跑过**：`服务池学员_YYYYMM.xlsx` 会被覆盖；如需保留历史，先手工备份或加 `--month` 跑其他月份。
- **匹配数为 0**：`config.json` 里 `liuzhuang_file` 月份对不上是最常见原因；先用 `--skip-liuyi --dry-run` 排查。
- **BI 文件下载下来是 0 行**：检查 `current_month` 是否为未来月份，或 `池子节点3=服务月` 在 BI 端是否还有学员。
- **auth_state.json 过期**：六一标签步骤会立即报"找不到新增按钮"；跑 `python $LIUYI_DIR\login.py` 重扫。

### 中断恢复

主脚本是幂等的：1-4 步重跑覆盖输出文件即可；5-6 步如果标签已建过，再次运行会复用现有 tag_id（避免重复建群）。

## 校验自检清单

跑之前过一遍：

- [ ] `config.json` 的 `current_month` 是当月（格式 `YYYY-MM`）
- [ ] `config.json` 的 `liuzhuang_file` 指向当月学员流转文件，且该文件存在
- [ ] `$LIUYI_DIR\auth_state.json` 修改时间在 7 天内（过期会失败）
- [ ] BI 浏览器没有其他人登录
- [ ] `$SKILL_DIR\data\processed\` 没有当月旧文件占着名（或确认可覆盖）

跑完后核对：

- [ ] `logs/service_pool_YYYYMMDD_*.log` 末尾出现"完成"字样
- [ ] `data/processed/服务池学员_YYYYMM.xlsx` Sheet1 行数 ≈ BI 下载行数 × 命中筛选比例（参考上月）
- [ ] Sheet1 的 `最终归属LP` 列非空率 ≥ 80%（低于这个值说明学员流转匹配异常）
- [ ] `dadou_ids_YYYYMM.xlsx` 第一列是纯数字（否则六一会报 code:101）
- [ ] 六一工作台能搜到 `【海外益智】26年X月服务池学员名单` 标签 + 同名用户群
- [ ] 数仓同步页面看到本月配置（业务系统=豌豆数仓 / 频率=每天 / 状态=启用）

## 反模式（不要这样做）

- **不要直接编辑 `data/downloads/` 里的 BI 原始文件再跑后续步骤**：脚本每次都会重新下载覆盖，手工改的会丢。要改条件就改 BI 端筛选项，或在 `config.json` 加二次筛选。
- **不要复制粘贴上月的 `服务池学员_YYYYMM.xlsx` 改名当本月用**：Sheet2 的大账号 ID 没换，上传到六一会创建错误标签。
- **不要在跑流程中途登录其他 BI 账号**：会触发"账号在其他地方登录"，已下载的部分数据可能不完整。
- **不要把 `auth_state.json` 提交到 Git 或发给别人**：里面是登录 cookie，等同账号密码。
- **不要为了"看起来更整齐"手动改 `dadou_ids_YYYYMM.xlsx` 的列名或加表头**：`create.py` 默认读第一列原始 ID，改了会报 `code: 101`。
- **不要跳过 `--skip-liuyi` 直接重跑全流程来"补一次"**：会重复建标签和用户群（同名也会建第二个），后续同步会指向错版本。要补就单跑 Step 6。
- **不要把 `current_month` 写成下个月来"提前跑"**：BI 端开课日期还没到，结果集会是空的。
- **不要在 SKILL.md 里写绝对桌面路径**：换机器或改账号就全断；用 `$SKILL_DIR` / `$DESKTOP` 等变量。

## 登录六一工作台（登录态过期时）

```bash
python "$LIUYI_DIR\login.py"
```

## 关键文件路径

| 文件 | 路径 |
|------|------|
| 主脚本 | `$SKILL_DIR\service_pool_automation.py` |
| 同步脚本 | `$SKILL_DIR\sync_tag_data.py` |
| 配置文件 | `$SKILL_DIR\config.json` |
| 六一登录态 | `$LIUYI_DIR\auth_state.json` |
| BI 导出工具 | `$BI_TOOL_DIR\smartbi_browser_export.py` |
| 六一标签工具 | `$LIUYI_DIR\create.py` |

## 常见问题

| 问题 | 原因 | 解决 |
|------|------|------|
| BI 下载失败：账号在其他地方登录 | BI 账号被占用 | 关闭其他 BI 标签页后重新运行 |
| 六一工作台找不到「新增」按钮 | auth_state.json 过期 | 重新运行 `login.py` |
| 匹配数为 0 | 学员流转文件路径或月份不对 | 检查 config.json 的 `liuzhuang_file` |
| 标签创建报 `code: 101` | dadou_ids 文件第一列不是纯数字 | 删表头/重跑 Step 1-4 |
| 同步报"用户群不存在" | Step 5 没跑成功就跳到 Step 6 | 先在六一工作台确认用户群已建 |
