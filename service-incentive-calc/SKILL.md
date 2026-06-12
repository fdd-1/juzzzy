---
name: service-incentive-calc
version: 2.1.0
description: 服务激励核算全自动化 - 下载BI报表→提取指标→计算激励→输出Excel。支持定时任务每月1号自动结算。
category: data-processing
tags: [excel, bi-report, incentive, automation, scheduled-task]
author: AI
---

# Service Incentive Calculation Skill

## 功能
全自动从 4 张 BI 服务绩效报表下载数据、提取 5 个关键指标、计算激励金额、输出核算 Excel。

## 详细参考（必读）

> 取数细节和报表 ID 已下沉到 `reference/`，本文只放流程概述。

- **指标取数规则**：见 [reference/metric_rules.md](reference/metric_rules.md)
- **BI 报表 ID / 下载规则**：见 [reference/bi_report_ids.md](reference/bi_report_ids.md)
- **凭据配置**：见根目录 `.env.example`

## 凭据配置（首次使用必做）

凭据走环境变量 / `.env` 文件加载，**禁止写入任何 .json / .py / .md 文件**。

```powershell
# 复制模板并填入实际值
copy .env.example .env
notepad .env

# 或者手动 set 当次会话
$env:SMARTBI_USERNAME="<工号>"; $env:SMARTBI_PASSWORD="<密码>"
```

## 使用方式

### 自动定时执行（已配置）
Windows 定时任务 `ServiceIncentiveCalc`，每月 1 号 10:30 自动计算上月激励。

### 手动执行
```bash
python run_monthly.py                                       # 计算上月（默认）
python run_monthly.py --month 5 --year 2026                 # 指定月份
python run_monthly.py --start 2026-05-01 --end 2026-05-27   # 指定日期范围（月中查看进度）
python run_monthly.py --skip-download --start ... --end ... # 跳过下载，直接用已有报表计算
```

## 日期范围规则
- **月中核算**：当月 1 号 ~ 今天
- **月初结算**：上月 1 号 ~ 上月最后一天
- 每次核算必须重新导出报表，不能复用旧数据

## 数据流
```
BI 系统 → 下载 4 张报表 → 提取 5 个指标 → 计算激励金额 → 输出 Excel + 校验清单
```

## 完整流程

🔴 **CHECKPOINT 1：下载前确认**
- 日期范围是否正确？（月中=当月1号到今天，月初结算=上月全月）
- `.env` 是否已配置（`SMARTBI_USERNAME` / `SMARTBI_PASSWORD`）？
- 4 个报表 ID 是否对应正确（首通别下成 M0 那张）→ 见 [reference/bi_report_ids.md](reference/bi_report_ids.md)

1. 创建本月配置 `service_incentive_<日期>.json`（复制上次的，改日期 overrides）
   - ⚠️ 如果上月配置文件不存在 → 从 skill 目录拷贝 `service_incentive_0605.json` 作为模板
2. 加载 `.env` 凭据，依次运行 4 个 task 下载报表
   - ⚠️ 下载超时 → 重试 3 次 → 失败则记录日志并终止
   - ⚠️ 环境变量未设置 → 报错 `auth missing` → 检查 `.env` 或 `$env:SMARTBI_USERNAME`
   - ⚠️ 某张报表 0 行 → 检查报表 ID（特别是首通别下错成 M0 那张）
3. 把下载结果从 `周报自动化/data/<月份>/` 拷到本 skill 的 `data/<月份>/`
4. 确认 `reference/<月份>服务激励-模板.xlsx` 存在（不存在就复制上月模板改名）

🔴 **CHECKPOINT 2：指标提取后验证**
- 5 个指标数值是否都在合理范围 → 见 [reference/metric_rules.md](reference/metric_rules.md) 「合理值范围」
- 是否都取到了「海外团队-总计」行（首通是「海外教学服务部-总计」）？
- 首课/首专数值是否疑似台湾组（曾经的 bug：77.78% / 89.47% 而非正确的 74.31% / 72.87%）？

5. 运行 `python run_monthly.py --start <开始> --end <结束> --skip-download` 计算
6. 输出文件落在 `output/<月份>/<月份>服务激励_<时间戳>.xlsx`

🔴 **CHECKPOINT 3：输出后校验清单（程序自动打印）**
程序运行结束会自动输出以下校验项，**任何一项 FAIL 都需要人工排查**：
- ✓/✗ 5 个指标全部提取成功（5/5）
- ✓/✗ 每个指标值在合理范围内
- ✓/✗ 激励总额 ≤ 2000 元
- ✓/✗ 输出 Excel 文件已生成

## 文件结构
```
service-incentive-calc/
├── run_monthly.py       # 全自动执行入口（下载+计算+输出+校验）
├── config.py            # 配置（指标映射、阈值、路径）
├── extract_metrics.py   # 指标提取模块
├── build_incentive.py   # Excel 构建模块
├── skill.md             # 本说明（精简流程）
├── .env.example         # 凭据模板（实际 .env 不入库）
├── .gitignore
├── data/                # BI 报表下载存放（按月份子目录）
│   └── 5月/
├── output/              # 核算结果输出（按月份子目录）
│   └── 5月/
└── reference/           # 详细规则与模板
    ├── metric_rules.md          # 指标取数规则（必读）
    ├── bi_report_ids.md         # 报表 ID 与下载规则（必读）
    ├── 5月服务激励方案.xlsx
    └── 5月服务激励-模板.xlsx
```

## 月度适配
激励方案变更时，将新方案文件放入 `reference/` 目录即可。如取数列/行变更，**只改 `reference/metric_rules.md` + `config.py` 的 `METRIC_SOURCES`**，不要改业务代码。

## 注意事项
- 定时任务需要电脑开机且用户已登录
- 下载用 smartbi_cli（HTTP 直连），别用 bi_skill 的 playwright 方式
- openpyxl 保存后公式不自动计算，需用 Excel 打开重算
- 依赖：openpyxl, smartbi_cli（已在周报自动化目录）

## ❌ 禁止操作黑名单

1. ❌ **不要复用旧月份的报表文件** —— BI 数据每天更新，旧报表会偏离实际
2. ❌ **不要下载「M0管理_益智海外新生首通监控」** —— 详见 [reference/bi_report_ids.md](reference/bi_report_ids.md)
3. ❌ **不要对首课/首专指标加 `data_row_offset`** —— 详见 [reference/metric_rules.md](reference/metric_rules.md)
4. ❌ **不要用 bi_skill 的 playwright 方式下载这 4 张报表** —— monitor 类会超时、GBK 编码错
5. ❌ **不要在月中用上月的日期范围**
6. ❌ **不要手动修改 `output/<月份>/` 下的输出 Excel** —— 改 `config.py` 后重跑
7. ❌ **不要把凭据写进 `.json` / `.py` / `.md`** —— 一律走 `.env` 或环境变量，`.env` 已加入 `.gitignore`
