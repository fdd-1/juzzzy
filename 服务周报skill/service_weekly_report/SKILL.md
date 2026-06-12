---
name: service-weekly-report
description: "服务周报自动化 (Service Weekly Report Automation)：海外思维服务周报全流程自动化，从 SmartBI 报表下载 → 数据整合格式化 → 飞书电子表格 → 结论生成 → 统一文档发布。覆盖 4.1-4.6 六大板块（首通/首课/首专/语义、组班意向、群发消息、停课唤醒、服务月跟进、服务池SOP、外呼监控、企微回复）。自动计算时间窗口（4.1 上周一-上周日；4.2-4.6 当月1号-上周日），含落后组识别和落后LP点名。Use when user says \"生成服务周报\", \"做本周服务周报\", \"跑 4.X 数据\", \"服务周报自动化\", \"跑周报\", \"weekly service report\"."
---

# 服务周报自动化 Skill

> **v1.0 · 2026-06-10** — 服务周报全流程自动化，从 BI 报表下载到飞书文档发布。
>
> 触发：用户说"生成本周服务周报"、"做服务周报"、"跑周报"

---

## 零、安全与路径规约（必读，凡新增/改动文件前必须自检）

### 0.1 显式反模式黑名单（出现即视为缺陷，必须修复）

> 以下写法**任何**情况下都不允许出现在源码、配置、文档、SOP、README、Skill 内联示例里。
> 自检方法：在项目根执行
> `Get-ChildItem -Recurse -Include *.py,*.md,*.json | Select-String -Pattern "fengjianyi|76218|123456|smartbi-data-cli-internal-20260526"`
> 命中任何一项就是缺陷。

| # | 反模式 | 为什么不行 | 正确做法 |
|---|---|---|---|
| 1 | 路径写死 `C:/Users/fengjianyi/...` 或 `C:\Users\fengjianyi\...` | 仅在原作者机器有效，换人/换机立即报路径不存在 | 项目内：`Path(__file__).resolve().parent[...]`；项目外工具：环境变量 |
| 2 | 路径写死带版本号的工具目录 `smartbi-data-cli-internal-20260526` | 工具升级换目录后全链路路径失效 | `_paths.resolve_smartbi_cli_dir()`，按 `$SMARTBI_CLI_DIR` 或同级 glob 自动查找 |
| 3 | 任何形式的明文工号/密码/token，例如 `"76218"`、`"123456"`、`Bearer xxx` | 凭据进 Skill = 任何拿到 Skill 的人都能登 BI，属信息安全事故 | 只读 `os.environ`；文档示例只用 `<your-username>` 占位符 |
| 4 | `cd "C:\Users\<人名>\..."` 这种依赖个人桌面布局的指令 | 协作场景下其他人复制粘贴必失败 | 用 `<…>\服务周报skill\service_weekly_report` 占位符或 `Set-Location -LiteralPath` 相对写法 |
| 5 | 在 `if __name__ == "__main__":` 里写死测试用绝对路径 | 改成另一周/另一台机器就跑不起来 | 用 `argparse`，`default` 写为 `PROJECT_ROOT / "exports" / ...` 这种相对值 |
| 6 | 把凭据写进 `configs/*.json`、`run.json`、`scheduled_tasks.json` 这类落盘文件 | 落盘 = 进 git / 进备份 / 进同步盘，泄露面立刻放大 | 配置文件只放 ID/筛选项，运行时再从环境变量注入 |

### 0.2 路径与凭据"怎么做"指引（标准操作）

#### A. 项目内路径（输入/输出/配置/模块）
- **必须**用 `Path(__file__).resolve().parent` 推算，**绝不**写绝对路径。
- 一处定义、多处复用：统一在 [_paths.py](_paths.py) 暴露 `PROJECT_ROOT`。
- `scripts/`、`modules/`、`test_download/` 下的脚本要 import 它，先：
  ```python
  sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
  from _paths import PROJECT_ROOT
  ```

#### B. 项目外工具路径（smartbi-data-cli、bi_skill）
- 解析顺序：① 环境变量 → ② 项目同级目录 glob 自动发现 → ③ 报错并打印"如何设置"。
- 已封装在 [_paths.py](_paths.py)：`resolve_smartbi_cli_dir()`、`resolve_bi_skill_path()`。
- 缺失时**禁止**默认到某条具体绝对路径；必须 `sys.exit` 并提示：
  ```
  [X] 未找到 smartbi-data-cli 工具目录。请设置环境变量：
     PowerShell: $env:SMARTBI_CLI_DIR = "<smartbi-data-cli-internal-* 绝对路径>"
  ```

#### C. 凭据
- 只通过 `os.environ.get("SMARTBI_USERNAME")` / `os.environ.get("SMARTBI_PASSWORD")` 读取。
- 启动时调用 `_paths.ensure_credentials()` 做"缺失即停"检查。
- 文档/示例里只能写：
  ```powershell
  $env:SMARTBI_USERNAME = "<your-username>"
  $env:SMARTBI_PASSWORD = "<your-password>"
  ```
- 不要写 `.env`、不要写 PowerShell profile 持久化的命令；让用户自己决定持久化方式（每次新会话设一次最稳）。

#### D. 提交前检查清单
- [ ] 全文 grep `fengjianyi`、`76218`、`123456`、`smartbi-data-cli-internal-20260526`，零命中
- [ ] 全文 grep `C:\Users\` / `C:/Users/`，零命中（_paths 内部例外）
- [ ] 新增脚本顶部使用 `_paths` 而非自己写绝对路径
- [ ] 配置 JSON 不含凭据字段

---

## 一、前置条件检查（执行前必须验证）

🔴 **CHECKPOINT**：以下条件不满足则停止并提示用户

```bash
# 1. 检查 SmartBI 凭据（仅检查环境变量是否存在，绝不在脚本里写值）
if [ -z "$SMARTBI_USERNAME" ] || [ -z "$SMARTBI_PASSWORD" ]; then
  echo "❌ 请先在当前 PowerShell 会话设置环境变量：SMARTBI_USERNAME / SMARTBI_PASSWORD"
  exit 1
fi

# 2. 检查 lark-cli 已认证
lark-cli auth status | grep "ready" || exit 1

# 3. 检查 smartbi-data-cli 工具目录可定位（自动从 $SMARTBI_CLI_DIR 或项目同级目录解析）
test -n "$SMARTBI_CLI_DIR" -o -n "$(ls -d ../smartbi-data-cli-internal* 2>/dev/null | head -n1)" || exit 1
```

**如果 SmartBI 凭据缺失** → 提示用户在 PowerShell 中执行（替换为本人凭据，**不要把凭据提交到代码或文档**）：
```powershell
$env:SMARTBI_USERNAME = "<your-username>"   # 例如工号
$env:SMARTBI_PASSWORD = "<your-password>"
# 可选：若 smartbi-data-cli 未放在项目同级目录，需指定绝对路径
$env:SMARTBI_CLI_DIR = "<smartbi-data-cli-internal-* 的绝对路径>"
```

> 🔴 **凭据红线**：SKILL / SOP / README / 任意配置 / 任意脚本中**永远不出现**真实工号、密码、token。看到示例值（`76218`、`123456`）就立即清掉，改写占位符。


---

## 二、时间窗口计算规则（自动）

| 板块 | 时间窗口 |
|---|---|
| 4.1 | 上周一 ~ 上周日 |
| 4.2-4.6 | 当月1号 ~ 上周日 |

示例（今天=2026-06-10 周二）：
- 上周一 = 2026-06-01
- 上周日 = 2026-06-07
- 当月1号 = 2026-06-01

**实现**：`scripts/download_smartbi_reports.py` 中的 `calculate_date_windows()` 自动计算。

---

## 三、完整工作流（5 步）

### 步骤 1：下载 SmartBI 报表（10 份）

**输入**：环境变量 `SMARTBI_USERNAME` / `SMARTBI_PASSWORD`
**输出**：`downloads/smartbi_reports/{run_date}/4_X_*/...xlsx`

```powershell
# 进入项目目录（替换为本机的 service_weekly_report 路径，或直接在该目录下打开 PowerShell）
Set-Location -LiteralPath "<…>\服务周报skill\service_weekly_report"

# 主要 9 份报表（不含 4.5 SOP，因为复用 4.1 SOP）
python scripts\download_smartbi_reports.py --output-dir "downloads\smartbi_reports"

# 4.5 服务月跟进单独下载（参数名特殊）
python scripts\download_4_5_fuwuyue.py
```

**报表清单**（含 Report ID 和筛选项）：

| 板块 | 报表名 | Report ID | 时间字段 |
|---|---|---|---|
| 4.1 | 益智海外新生首通监控 | `I2c928087018722bf22bf7d4d0187ff868fab30da` | 首次分配开始/结束时间 |
| 4.1 | 海外思维学管服务指标统计表（首课） | `I2c9280870189f6f1f6f10d05018a64543b1960d0` | 开始日期/结束时间/LP做工开始/结束时间 |
| 4.1 | 海外思维学管服务指标统计表（首专） | `I2c9280870189f6f1f6f10d05018a64543b1960d0` | 同上（共用 ID） |
| 4.1 | 海外思维服务SOP执行情况 | `I2c928087019364b764b704c8019375e98bea20d9` | 日期/做工开始/结束时间，**海外思维团队="" 空字符串** |
| 4.2 | 思维LP组班意向提交播报 | `I2c9280870199baeabaea35bb0199d14379526743` | 开始时间/结束时间 |
| 4.3 | 思维海外群发消息汇总数据播报 | `I2c92808701989a869a8616730198bc740f510455` | 开始时间/结束时间 |
| 4.4 | 思维停课学员执行监控 | `I2c928087018b5acd5acd0fdc018c24e77b3c3525` | 开始时间/结束时间 |
| 4.5 | 思维转介绍过程跟进报表_末次渠道 | `I2c9280870191447f447f3a940191b1cf2a202710` | **做工开始时间_周维度/结束日期/日期**（特殊参数名！） |
| 4.5 | 海外思维服务SOP执行情况（服务池） | 复用 4.1 SOP | 不重复下载 |
| 4.6 | LP系统外呼监控-分池子 | `I2c928087019c8be48be48889019c97e154655e1e` | 开始时间/结束时间 |
| 4.6 | LP企微回复比监控-分池子 | `I2c928087019c8be48be48889019c994d3f4a3e2c` | 开始时间/结束时间 |

**🔴 失败处理**：
- 如果 SmartBI 登录失败 → 检查凭据，验证账号未过期
- 如果某报表下载超时（>180s）→ 单独重试 `--task service_weekly_4_X`
- 如果 SOP 数据为空（只有 24 行口径） → 检查"海外思维团队"参数值是否为空字符串 `""`（不是"全选"）
- 如果 4.5 服务月数据时间不对 → 用 `download_4_5_fuwuyue.py` 单独下载

---

### 步骤 2：数据整合与格式化

**输入**：`downloads/smartbi_reports/{run_date}/`
**输出**：`exports/weekly_{start}_{end}/_merged_4_X.xlsx`

```powershell
python scripts\process_data.py `
  --downloads-dir "downloads\smartbi_reports\{run_date}" `
  --output-dir "exports\weekly_{start}_{end}"
```

**格式化规则**：

| 列头特征 | 格式 | 示例 |
|---|---|---|
| 含"占比"/"率"且**不含**"加和" | 百分比，保留两位小数 | `0.8573` → `85.73%` |
| **含"加和"（如执行率加和、语义点执行率加和）** | **数值，保留两位小数（不带%）** | `0.535212` → `5.35`，`1.55` → `1.55` |
| 列尾是"数"/"人数"/"次数" | 整数 | `142` |
| 其他数值列 | 保留两位小数 | `1.55` |

**🔴 加和列规则特别注意**：
- 4.1 中的 `首通语义点执行_执行率加和`、`首课语义点执行_执行率加和` → 保留两位小数（不要转 `535.21%`）
- 4.5 SOP 中的 `服务池-语义点执行率加和` → 保留两位小数
- 这些列代表多个执行率的求和，不是百分比比率，所以不用 % 显示

**清洗规则**：
- forward-fill 合并单元格列（团队/小组列）
- 删除全空行、全空列
- 删除口径说明行
- 排除大区总计（台湾、欧美澳、港澳总计）

**4.1 整合特殊处理**：5 张表 LEFT JOIN，主键 (团队 + LP)

**列序**：团队 → LP → 首通语义点执行 → 首通 → 首课语义点执行 → 首课 → 首专 → LP入职时长

**4.5 服务月跟进**：从转介绍报表中提取"服务池"列（列113-125）

**4.5 服务池SOP**：复用 4.1 SOP 数据，提取服务池部分（列34-40），加和列移到 LP 右边

**🔴 失败处理**：
- 如果 4.1 缺少 LP 架构表 → 跳过 LP 架构合并（已在 `processor_4_1.py` 中容错）
- 如果某板块表头解析失败 → 检查 `excel_parser.py` 的 `key_column` 参数

---

### 步骤 3：创建飞书电子表格（8 个）

**输入**：`exports/weekly_{start}_{end}/_merged_4_X.xlsx`
**输出**：8 个独立飞书电子表格的 token

```powershell
# 4.1（特殊处理：色阶+排序+居中）
python modules\feishu_builder_4_1.py `
  --merged "exports\weekly_{start}_{end}\_merged_4_1.xlsx" `
  --start-date {start} --end-date {end}

# 4.2-4.6（通用处理，逐个执行）
python modules\feishu_simple_builder.py --input "..._merged_4_2.xlsx" --title "4.2 组班意向 0601-0607"
python modules\feishu_simple_builder.py --input "..._merged_4_3.xlsx" --title "4.3 群发消息 0601-0607"
python modules\feishu_simple_builder.py --input "..._merged_4_4.xlsx" --title "4.4 停课唤醒 0601-0607"
python modules\feishu_simple_builder.py --input "..._merged_4_5_fuwuyue.xlsx" --title "4.5 服务月跟进 0601-0607"
python modules\feishu_simple_builder.py --input "..._merged_4_5_sop.xlsx" --title "4.5 服务池SOP 0601-0607"
python modules\feishu_simple_builder.py --input "..._merged_4_6_waihu.xlsx" --title "4.6 外呼监控 0601-0607"
python modules\feishu_simple_builder.py --input "..._merged_4_6_qiwei.xlsx" --title "4.6 企微回复 0601-0607"
```

记录每个表格返回的 `token`，写入 `final_doc_builder_v3.py` 的 `SHEET_TOKENS` 字典。

---

### 步骤 4：飞书表格优化（删空行 + 上色阶）

**输入**：8 个飞书表格的 token
**输出**：清理后的飞书表格（含色阶）

```powershell
python scripts\polish_feishu_sheets.py
```

**色阶配置**：

| 板块 | 色阶列 | 反向（越低越好） |
|---|---|---|
| 4.1 | 各执行率/跟进率/企微绑定率 | 秒挂占比 |
| 4.3 | 个人群发占比 | - |
| 4.4 | 唤醒率 | 停课占比 |
| 4.5 服务月 | 外呼跟进率/综合有效跟进率 | - |
| 4.5 SOP | 语义点执行率加和 | - |
| 4.6 外呼 | 整体覆盖率/外呼接通率/有效接通率 | - |

---

### 步骤 5：生成结论 + 创建统一文档

**输入**：8 个 _merged_4_X.xlsx + 8 个飞书表格 token
**输出**：1 个统一周报文档（含 8 个 callout + 嵌入表格）

```powershell
python modules\final_doc_builder_v3.py
```

**结论格式**（参考 https://my.feishu.cn/docx/SbLFdUogiouIErx0zpXcw4krnij）：

```xml
<callout emoji="❗">
<p><b>整体：</b>[关键指标汇总]</p>

<p><b>亮点</b></p>
<p>·[小组+具体数据]</p>

<p><b>风险</b></p>
<p>·[落后小组+具体百分比]</p>
<p>·[落后小组内的落后LP+百分比]</p>

<p><b>待办</b></p>
<p>·[具体可执行行动]</p>
</callout>
```

**移动到目标文件夹**：

```bash
lark-cli drive +move \
  --file-token <DOC_ID> \
  --folder-token JpSRflVoWlwxZxdBgg7cFbBNnrc \
  --type docx
```

**最终交付**：飞书文档 URL（位置：`云盘 > 09-思维后端-LP周会 > 后端业务周会`）

---

## 四、检查点（每步完成后）

🔴 **STEP 1 完成后**：检查 `downloads/smartbi_reports/{run_date}/` 下有 10 个子文件夹，每个文件夹都有 `.xlsx`

🔴 **STEP 2 完成后**：检查 `exports/weekly_{start}_{end}/` 下有 8 个 `_merged_4_X.xlsx`，每个 > 5KB

🔴 **STEP 3 完成后**：每个飞书表格能在浏览器打开，数据完整

🔴 **STEP 4 完成后**：表格底部无空行，相关列有色阶

🔴 **STEP 5 完成后**：要求用户打开最终文档确认：
- 8 个板块的标题、callout、嵌入表格全部存在
- 4.1 数据表正确嵌入（重点：之前出过 bug）
- 落后小组/落后LP 在风险部分被点名
- 文档已移动到目标文件夹

---

## 五、关键技巧（必读，避免重复踩坑）

### 5.1 SOP 报表"海外思维团队"全选 = 空字符串

**坑**：写 "全选" 或具体服务区名都会导致数据为空（只有 24 行口径）。
**正解**：value 和 displayValue 都传 `""`。
**验证**：修复后从 24 行 → 119 行（10 个小组完整）。

### 5.2 4.5 转介绍报表的特殊参数名

**坑**：用 "开始时间"/"结束时间" 会被忽略。
**正解**：必须用 `做工开始时间_周维度`、`结束日期`、`日期`。

### 5.3 4.1 数据表 sheet_id 不能写死

**坑**：`final_doc_builder_v3.py` 之前把 4.1 sheet_id 写死为 `6e8dab`，但每次新建表格 sheet_id 不同。
**正解**：所有 sheet_id 都通过 `get_sheet_id()` API 实时获取。

### 5.4 LP 架构表是 SIMPLE_REPORT

smartbi-data-cli 不支持导出 SIMPLE_REPORT，跳过即可。
如必须需要：用 bi_skill（playwright 浏览器方式）下载。

### 5.5 数据条色阶 = 渐变色彩

飞书表格不直接支持数据条，用 `+batch-set-style` 逐单元格设置背景色（绿→黄→红）模拟。

### 5.6 文档目标文件夹

`JpSRflVoWlwxZxdBgg7cFbBNnrc` (云盘 > 09-思维后端-LP周会 > 后端业务周会)

---

## 六、文件资源（references）

```
service_weekly_report/
├── SKILL.md                              # 本文档
├── SOP.md                                # 详细 SOP（同等权威）
├── README.md                             # 项目说明
├── main.py                               # 主流程入口
├── scripts/
│   ├── download_smartbi_reports.py       # 下载 9 份主要报表
│   ├── download_4_5_fuwuyue.py           # 4.5 服务月单独下载
│   ├── download_sop_fixed.py             # 4.1 SOP 单独下载（空字符串技巧）
│   ├── process_data.py                   # 数据整合
│   └── polish_feishu_sheets.py           # 删空行 + 上色阶
├── configs/
│   ├── service_weekly_smartbi_tasks.json # 12 个报表配置
│   └── _dynamic/                         # 动态生成的配置
├── modules/
│   ├── data_formatter.py                 # 数据格式化
│   ├── processor_4_1.py                  # 4.1 整合
│   ├── processor_4_4_v3.py               # 4.4 整合
│   ├── feishu_simple_builder.py          # 飞书表格通用创建
│   ├── feishu_builder_4_1.py             # 4.1 特殊处理（含色阶）
│   ├── conclusion_generator_v2.py        # 结论生成（参考文档格式）
│   └── final_doc_builder_v3.py           # 最终文档+移动到目标文件夹
└── docs/
    ├── SMARTBI_CLI_GUIDE.md              # SmartBI CLI 使用说明
    └── REPORT_ID_MAPPING.md              # 报表 ID 映射表
```

---

## 七、一键执行（推荐）

```powershell
# 1. 设置凭据（首跑前；不要把凭据写进任何文件）
$env:SMARTBI_USERNAME = "<your-username>"
$env:SMARTBI_PASSWORD = "<your-password>"

# 可选：如果 smartbi-data-cli 没放在与本项目同级的目录，再额外指定
# $env:SMARTBI_CLI_DIR = "<smartbi-data-cli-internal-* 绝对路径>"

# 2. 进入项目目录（用 -LiteralPath 兼容中文路径）
Set-Location -LiteralPath (Split-Path -LiteralPath $PSCommandPath -Parent)
# 或直接 cd 到本 SKILL.md 所在目录

# 3. 执行主流程
python main.py
```

执行时间：约 8-12 分钟（下载占大头）。

---

## 八、错误恢复表（具体到命令）

| 错误现象 | 排查命令 |
|---|---|
| `auth_error: SMARTBI_USERNAME is required` | 在当前 PowerShell 设置 `$env:SMARTBI_USERNAME` / `$env:SMARTBI_PASSWORD`（值用本人凭据，**不要**写进文件） |
| `UnicodeDecodeError: 'gbk' codec` | 已在脚本中修复（utf-8 + errors='replace'） |
| SOP 数据只有 24 行口径 | 检查 `service_weekly_smartbi_tasks.json` 中"海外思维团队" value 是否为 `""` |
| 4.5 服务月时间错 | `python scripts\download_4_5_fuwuyue.py` 单独重下 |
| 4.1 数据表未嵌入文档 | 检查 `final_doc_builder_v3.py` 第 152-156 行是否动态获取 sheet_id |
| 飞书表格底部有空行 | `python scripts\polish_feishu_sheets.py` |
| 文档未移到目标文件夹 | `lark-cli drive +move --file-token <DOC_ID> --folder-token JpSRflVoWlwxZxdBgg7cFbBNnrc --type docx` |

---

## 九、版本历史

- **v1.0 (2026-06-10)**：完整流程跑通，包括：
  - SmartBI 自动下载（自动算时间）
  - SOP 空字符串技巧（解决"全选"问题）
  - 4.5 转介绍特殊参数名修复
  - 飞书表格删空行 + 色阶
  - 参考文档格式的结论生成（亮点/风险/待办）
  - 文档自动移动到目标文件夹
