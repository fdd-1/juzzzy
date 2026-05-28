# 学情积分核算

自动从 BI 提取海外思维学情课包数据，核算学员积分发放，自动生成发放模板，并自动提交 OA 豌豆币添加申请。

> **注意**：本仓库为**代码版本**，不含真实学员数据、BI 报表、OA 登录态。`samples/` 下提供了一份脱敏样例展示数据结构。详细技术文档见 [SKILL.md](SKILL.md)。

## 流程总览

```
BI 取数 → 4 条件筛选核算 → 生成发放模板（带日期） → 询问确认 → 自动提交 OA
```

5 步串成一条命令：

```bash
python xueqing_credit_skill.py run --auto --submit-oa
```

## 目录结构

```
学情积分核算/
├── xueqing_credit_skill.py        # 主入口
├── SKILL.md                       # 完整技术文档（OA 自动化关键路径）
├── README.md                      # 本文档
├── scripts/
│   ├── run.py                     # 一键运行（取数 → 处理 → 填模板）
│   ├── fetch_reports.py           # BI 取数
│   ├── fetch_reports_smartbi.py   # SmartBI 取数（备选）
│   ├── smartbi_browser_export.py  # SmartBI 浏览器导出
│   ├── process_xueqing.py         # 数据处理（4 条件筛选 + 积分计算）
│   ├── fill_wandou_template.py    # 填发放模板（文件名带 _YYYYMMDD）
│   ├── setup_scheduled_task.ps1   # Windows 定时任务配置
│   └── oa_login/
│       ├── login_oa_qr.py         # 扫码登录（生成 auth_state.json）
│       ├── submit_oa.py           # 自动提交 OA（Playwright）
│       ├── record_dept_flow.py    # 录制工具（K2 控件 debug 用）
│       ├── enter_oa.py
│       └── fill_oa_form.py
├── configs/
│   └── smartbi_simple_report_tasks.json
├── docs/
│   └── 流程.md
├── samples/
│   ├── 积分汇总_sample.xlsx       # 脱敏样例（结构展示）
│   └── generate_sample.py
└── 03_output/
    └── 发放豌豆币文档填写模板.xlsx  # 原始空模板（拷贝源）
```

## 准备

1. **Python 环境**

```bash
pip install openpyxl playwright
playwright install chrome
```

2. **OA 扫码登录**（首次必做，登录态保存到本地）

```bash
python scripts/oa_login/login_oa_qr.py
```

执行后会生成 `scripts/oa_login/auth_state.json`，含会话 Cookie，**已加入 .gitignore，不会上传**。

3. **配置 BI 报表**

编辑 `configs/smartbi_simple_report_tasks.json`，根据自己的 BI 环境调整 `base_url` 和 `report.id`。

## 用法

### 一条龙

```bash
# 自动模式（根据当前日期算区间：1 号 → 上月 16-月底，16 号 → 本月 1-15）
python xueqing_credit_skill.py run --auto --submit-oa

# 手动指定区间
python xueqing_credit_skill.py run --start 2026-05-01 --end 2026-05-15 --submit-oa

# 无人值守（跳过提交前确认，定时任务用）
python xueqing_credit_skill.py run --auto --submit-oa --yes
```

> 提交 OA 前会列出期次目录、附件、积分汇总，并询问 `数据是否无误？(y/N)`。输 `y` 才会调 Playwright 提交，其它键取消。

### 分步运行

```bash
# 只跑核算 + 模板生成
python xueqing_credit_skill.py run --auto

# 跳过 BI 取数，用已下载的报表重跑
python xueqing_credit_skill.py run --auto --skip-fetch

# 单独提交 OA（自动找最新期次最新模板）
python xueqing_credit_skill.py submit-oa
```

### 配置定时任务

```bash
python xueqing_credit_skill.py setup-schedule
```

逻辑：
- 每月 1 号 09:30：计算上月 16 号 ~ 上月最后一天
- 每月 16 号 09:30：计算本月 1 号 ~ 本月 15 号

## 核算规则

发放条件（4 条件全部满足）：
1. 在课包池中：`(学生 ID, 课时包 ID)` 命中累计课包池
2. 是否预习 = 1
3. 线上作业提交状态 = 已提交
4. 消耗课时 ≥ 1

积分公式：

```
发放积分 = (基础课时消耗 + 赠送课时消耗) × 500
```

## OA 字段映射

| 字段 | 值 | 控件 |
|---|---|---|
| 申请类型 | 豌豆商城豌豆币添加 | radio |
| 申请原因 | 海外益智学情积分（预习课后习题消课）| 文本 |
| 科目 | 益智 | radio |
| 是否合同赠送 | 否 | radio |
| 虚拟币类型 | 豌豆币 | radio |
| 活动批准审批单号 | 无 | 文本 |
| 用户ID | 详情见附件 | 文本 |
| 积分成本归属部门 | 欢乐童年_海外直播业务线_海外业务中心_海外业务运营处_海外教学服务运营组 | **K2 inputselect autocomplete** |
| 需要添加的豌豆币/魔力币数总共 | 由 `积分汇总.xlsx` 自动求和 | 文本 |
| 附件 | `发放豌豆币文档填写模板_YYYYMMDD.xlsx` | 文件 |

K2 inputselect 控件的处理路径（`fill_inputselect_field`）见 [SKILL.md](SKILL.md)。

## 安全注意

- `scripts/oa_login/auth_state.json` 含会话 Cookie，**严禁上传**。已在 `.gitignore` 中。
- `01_bi_exports/` 和 `03_output/<期次目录>/` 含真实学员数据，**严禁上传**。已在 `.gitignore` 中。
- `submit_step*.png`、`record_events*.json` 是调试产物，可能含截图/DOM 数据，已忽略。

## 依赖

- Python 3.12+
- openpyxl
- playwright（chrome channel）
- bi_skill（内部 BI 报表下载工具）

## License

内部项目，仅供学习参考。
