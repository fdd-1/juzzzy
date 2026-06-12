# 服务周报自动化 - 完整流程文档

> 最后更新：2026-06-09

---

## 一、项目概述

### 1.1 目标
自动化生成海外思维服务周报（4.1-4.6 六个板块），从报表下载到飞书文档生成的全流程自动化。

### 1.2 时间窗口规则（已记住）
- **4.1 板块**：上周一到上周日
- **4.2-4.6 板块**：当月1号到上周日（自动计算）

### 1.3 凭据要求
> 仅通过环境变量传入，**不要**写进任何文件、脚本、配置或文档。
```powershell
$env:SMARTBI_USERNAME = "<your-username>"
$env:SMARTBI_PASSWORD = "<your-password>"
# 可选：smartbi-data-cli 不在项目同级目录时
$env:SMARTBI_CLI_DIR = "<smartbi-data-cli-internal-* 绝对路径>"
```

---

## 二、项目结构

```
service_weekly_report/
├── main.py                          # 主流程脚本
├── SOP.md                           # 完整 SOP 文档
├── scripts/
│   ├── download_smartbi_reports.py  # SmartBI 报表下载器（自动计算时间）
│   ├── process_data.py              # 数据整合与格式化
│   └── ...（其他脚本待集成）
├── configs/
│   ├── service_weekly_smartbi_tasks.json  # SmartBI 配置（12个报表ID）
│   └── _dynamic/                    # 动态生成的配置（含计算后的日期）
├── downloads/
│   └── smartbi_reports/
│       └── 2026-06-09/              # 按运行日期组织
│           ├── 4_1_shoutong/        # 益智海外新生首通监控
│           ├── 4_1_shouke/          # 学管服务指标（首课）
│           ├── 4_1_shouzhuan/       # 学管服务指标（首专）
│           ├── 4_1_sop/             # SOP执行情况（4.1 和 4.5 共用）
│           ├── 4_2/                 # 组班意向
│           ├── 4_3/                 # 群发消息
│           ├── 4_4/                 # 停课唤醒
│           ├── 4_5_fuwuyue/         # 转介绍_服务月
│           ├── 4_6_waihu/           # 外呼监控
│           └── 4_6_qiwei/           # 企微回复
├── exports/
│   └── weekly_YYYYMMDD_YYYYMMDD/   # 处理后的数据
│       ├── _merged_4_1.xlsx
│       ├── _merged_4_2.xlsx
│       └── ...
├── modules/                         # 数据处理模块（现有代码）
│   ├── data_formatter.py
│   ├── excel_parser.py
│   ├── processor_4_1.py
│   ├── processor_4_4_v3.py
│   └── ...
├── docs/
│   ├── SMARTBI_CLI_GUIDE.md         # SmartBI CLI 使用说明
│   └── REPORT_ID_MAPPING.md         # 报表ID映射表
└── test_download/                   # 测试脚本
    └── test_single_report.py
```

---

## 三、已完成的工作（2026-06-09）

### ✅ 3.1 SmartBI 报表下载
- [x] 找到所有 12 个报表的 ID
- [x] 创建 SmartBI 配置文件
- [x] 实现智能下载脚本（自动计算时间窗口）
- [x] 优化：4.5 SOP 与 4.1 SOP 共用同一份数据
- [x] 测试：Dry-run 通过（11个任务）
- [x] 实际下载：10 个报表全部成功

**输出路径**：
```
downloads/smartbi_reports/2026-06-09/
```

**下载结果**：
| 文件夹 | 文件名 | 大小(KB) | 状态 |
|---|---|---|---|
| 4_1_shoutong | 益智海外新生首通监控.xlsx | 38.0 | ✅ |
| 4_1_shouke | 海外思维学管服务指标统计表.xlsx | 55.8 | ✅ |
| 4_1_shouzhuan | 海外思维学管服务指标统计表.xlsx | 55.8 | ✅ |
| 4_1_sop | 海外思维服务SOP执行情况.xlsx | 13.6 | ✅ |
| 4_2 | 思维LP组班意向提交播报.xlsx | 12.1 | ✅ |
| 4_3 | 思维海外群发消息汇总数据播报.xlsx | 18.9 | ✅ |
| 4_4 | 思维停课学员执行监控.xlsx | 72.1 | ✅ |
| 4_5_fuwuyue | 思维转介绍过程跟进报表_末次渠道.xlsx | 104.4 | ✅ |
| 4_6_waihu | LP系统外呼监控-分池子.xlsx | 88.2 | ✅ |
| 4_6_qiwei | LP企微回复比监控-分池子.xlsx | 75.5 | ✅ |

### ✅ 3.2 数据处理脚本
- [x] 创建 `process_data.py`（适配新的下载路径）
- [x] 测试：7/8 板块处理成功
- [x] 输出格式化的 Excel 文件

**已集成的板块**：
- 4.2 组班意向 ✅
- 4.3 群发消息 ✅
- 4.4 停课唤醒 ✅
- 4.5 服务月跟进 ✅
- 4.5 服务池SOP ⚠️（行数为0，需要修复）
- 4.6 外呼监控 ✅
- 4.6 企微回复 ✅

**待修复**：
- 4.1 整合（缺少 LP 架构表）
- 4.5 SOP（表头解析问题）

### ✅ 3.3 主流程框架
- [x] 创建 `main.py`（5步流程框架）
- [x] 步骤1：下载报表 ✅
- [x] 步骤2：数据处理（框架完成，待完善）
- [ ] 步骤3：飞书表格生成（待集成）
- [ ] 步骤4：结论生成（待集成）
- [ ] 步骤5：最终文档（待集成）

---

## 四、使用方法

### 4.1 完整流程（一键执行）

```powershell
# 设置凭据（首跑前；不要写入任何文件）
$env:SMARTBI_USERNAME = "<your-username>"
$env:SMARTBI_PASSWORD = "<your-password>"

# 进入项目目录（在本机的 service_weekly_report 文件夹下打开 PowerShell，或替换占位符）
Set-Location -LiteralPath "<…>\服务周报skill\service_weekly_report"

# 执行主流程
python main.py

# Dry-run 测试（不实际下载）
python main.py --dry-run

# 跳过下载（使用已有数据）
python main.py --skip-download
```

### 4.2 单独下载报表

```powershell
# 下载到项目目录
python scripts\download_smartbi_reports.py --output-dir "downloads\smartbi_reports"

# Dry-run 测试
python scripts\download_smartbi_reports.py --dry-run

# 指定并发数
python scripts\download_smartbi_reports.py --max-workers 5
```

### 4.3 单独处理数据

```powershell
python scripts\process_data.py `
  --downloads-dir "downloads\smartbi_reports\2026-06-09" `
  --output-dir "exports\weekly_20260601_20260607"
```

### 4.4 测试单个报表

```powershell
cd test_download
python test_single_report.py --task service_weekly_4_3         # Dry-run
python test_single_report.py --task service_weekly_4_3 --execute  # 实际下载
```

---

## 五、关键设计决策

### 5.1 时间窗口自动计算
- 脚本自动识别"今天"，计算出上周一、上周日、当月1号
- 用户无需手动输入日期，只需说"生成本周服务周报"

### 5.2 报表 ID 映射
- 所有 12 个报表的 ID 已固化在配置文件中
- 使用动态配置生成器将时间窗口转换为 `extra_params`

### 5.3 数据复用优化
- 4.1 SOP 和 4.5 SOP 使用同一份报表数据
- 首课和首专使用同一个 report_id，通过筛选项区分

### 5.4 输出目录组织
```
downloads/smartbi_reports/{run_date}/    # 原始下载文件
exports/weekly_{start}_{end}/            # 处理后的数据
```

### 5.5 顺序下载模式
- 放弃 smartbi-data-cli 的批量模式（编码问题）
- 改用顺序下载（一个一个下载），避免 GBK/UTF-8 冲突

---

## 六、待完成的工作

### 🔧 6.1 修复现有问题
1. **4.1 整合**：LP 架构表缺失（SIMPLE_REPORT 类型，需要特殊处理）
2. **4.5 SOP**：表头解析导致 0 行输出（需要调整 process_data.py）

### 📝 6.2 集成飞书模块
1. 飞书电子表格生成（使用现有的 `feishu_simple_builder.py`）
2. 样式应用（表头、汇总行、数据条色阶）
3. 创建 8 个独立的电子表格

### 📊 6.3 集成结论生成
1. 使用现有的 `conclusion_generator_v2.py`
2. 按参考文档格式生成 callout（整体→亮点→风险→待办）
3. 风险部分：落后小组 + 组内落后LP

### 📄 6.4 集成文档生成
1. 使用现有的 `final_doc_builder_v3.py`
2. 创建最终统一文档（包含全部 callout + 嵌入表格）
3. 移动到目标文件夹：`JpSRflVoWlwxZxdBgg7cFbBNnrc`

### 🎨 6.5 优化与测试
1. 端到端测试（下载 → 处理 → 飞书 → 文档）
2. 错误处理与重试机制
3. 日志输出优化
4. 文档重命名规则（首课/首专）

---

## 七、技术栈

| 组件 | 工具/库 | 版本 |
|---|---|---|
| 报表下载 | smartbi-data-cli | internal-20260526 |
| 数据处理 | pandas, openpyxl | - |
| 飞书 API | lark-cli | - |
| Python | Python 3.12 | - |
| 操作系统 | Windows 10 | - |

---

## 八、故障排查

### 8.1 常见问题

**问题 1：凭据错误**
```
auth_error: SMARTBI_USERNAME/SMARTBI_PASSWORD is required
```
**解决**：在当前 PowerShell 会话设置环境变量（**不要**把值写进任何文件）
```powershell
$env:SMARTBI_USERNAME = "<your-username>"
$env:SMARTBI_PASSWORD = "<your-password>"
```

**问题 2：中文路径编码错误**
```
UnicodeDecodeError: 'gbk' codec can't decode
```
**解决**：已在脚本中设置 `encoding='utf-8', errors='replace'`

**问题 3：报表无数据**
```
报表无数据，点击刷新...
```
**解决**：检查筛选项日期是否正确，或手动在 BI 系统中刷新报表

**问题 4：LP 架构表不支持**
```
SIMPLE_REPORT 类型不支持导出
```
**解决**：LP 架构表已标记为 `enabled: false`，跳过下载

---

## 九、联系与支持

- **项目位置**：本仓库根目录（`service_weekly_report/`），无固定盘符；以 `Path(__file__).resolve().parent` 推算
- **SmartBI CLI**：通过 `$env:SMARTBI_CLI_DIR` 指定，或放在与 `service_weekly_report` 同级目录（自动发现 `smartbi-data-cli-internal-*`）
- **目标飞书文件夹**：https://hcnig43mb8gp.feishu.cn/drive/folder/JpSRflVoWlwxZxdBgg7cFbBNnrc

---

## 十、版本历史

| 版本 | 日期 | 更新内容 |
|---|---|---|
| v1.0 | 2026-06-09 | 完成报表下载、数据处理框架、主流程搭建 |

---

**下一步**：修复 4.1/4.5 数据处理问题，集成飞书模块。
