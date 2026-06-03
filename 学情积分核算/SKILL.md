# 学情积分核算 Skill

自动从BI提取海外思维学情课包数据，核算学员积分发放，自动填发放模板，并自动提交 OA 豌豆币添加申请。

## 功能特性

- ✅ 自动从BI下载续费规划表和上课明细
- ✅ 4条件智能筛选：在课包池 + 是否预习=1 + 线上作业已提交 + 消耗课时>=1
- ✅ 自动计算积分：(基础课时消耗 + 赠送课时消耗) × 500
- ✅ 分文件输出：积分汇总、上课明细带标注、续费规划表筛选
- ✅ 自动累积课包池：每期自动使用上期的课包池数据
- ✅ 自动填写「发放豌豆币文档填写模板」，文件名带当天日期
- ✅ Windows定时任务：每月1号和16号自动触发
- ✅ OA 豌豆币添加申请自动提交（Playwright，复用扫码登录态，含 K2 inputselect 控件处理）
- ✅ 一条龙：取数 → 核算 → 填模板 → 提交 OA

## 重要规则

### 报表下载规则（必读）

⚠️ **每次核算不同时间段的积分时，必须重新下载报表**

- ✅ **默认行为**：每次运行 `run` 命令都会从BI重新下载报表
- ❌ **禁止**：不同时间段的核算使用同一份报表文件
- ✅ **唯一例外**：同一天内，对同一时间段重复核算时，可使用 `--skip-fetch` 跳过下载

**示例**：
```bash
# 第一次核算 5月16-31日 → 自动下载报表 ✅
python xueqing_credit_skill.py run --start 2026-05-16 --end 2026-05-31

# 同一天内，发现数据有问题，重新核算同一时间段 → 可跳过下载 ✅
python xueqing_credit_skill.py run --start 2026-05-16 --end 2026-05-31 --skip-fetch

# 第二天核算 6月1-15日 → 必须重新下载报表 ✅
python xueqing_credit_skill.py run --start 2026-06-01 --end 2026-06-15

# ❌ 错误：用昨天的报表算今天的数据
python xueqing_credit_skill.py run --start 2026-06-01 --end 2026-06-15 --skip-fetch
```

**原因**：BI报表数据会实时更新，使用旧报表会导致积分核算结果不准确。

## 快速开始

### 一条龙（推荐）

```bash
# 取数 → 核算 → 生成模板 → 自动提交 OA（前提：已扫码登录）
python xueqing_credit_skill.py run --auto --submit-oa

# 手动指定区间也行
python xueqing_credit_skill.py run --start 2026-05-01 --end 2026-05-15 --submit-oa

# 无人值守（跳过提交前确认，定时任务用）
python xueqing_credit_skill.py run --auto --submit-oa --yes
```

> 提交 OA 前会列出期次目录、附件、积分汇总，并询问 `数据是否无误？(y/N)`。输 `y` 才会调用 Playwright 提交，其它键取消。无人值守可加 `--yes` 跳过。

### 分步运行

```bash
# 只跑核算 + 模板生成（不提交 OA）
python xueqing_credit_skill.py run --auto

# 跳过 BI 取数，用已下载的报表重跑
python xueqing_credit_skill.py run --auto --skip-fetch

# 单独提交 OA（自动找最新期次最新模板）
python xueqing_credit_skill.py submit-oa

# 指定期次目录提交
python xueqing_credit_skill.py submit-oa --output-dir "03_output/20260501-20260515学情积分发放明细"
```

### 配置定时任务

```bash
python xueqing_credit_skill.py setup-schedule
```

定时任务逻辑：
- **每月1号 09:30**：计算上月16号 ~ 上月最后一天
- **每月16号 09:30**：计算本月1号 ~ 本月15号

## 完整流程

`run --auto --submit-oa` 串起 5 步：

1. **BI 取数** —— 下载续费规划表 + 上课明细到 `01_bi_exports/`
2. **数据处理** —— 4条件筛选，输出 `03_output/<YYYYMMDD-YYYYMMDD学情积分发放明细>/积分汇总.xlsx` 等 3 个文件
3. **生成发放模板** —— 用积分汇总的「获得积分明细」填模板，输出 `03_output/<期次>/发放豌豆币文档填写模板_YYYYMMDD.xlsx`（**文件名带当天日期，方便溯源**）
4. **提交 OA**（仅 `--submit-oa` 时）—— Playwright 自动登录、点入口、填表、上传附件、提交
5. **保留浏览器 60 秒** —— 便于人工核对

## OA 表单字段映射

| 字段 | 值 | 控件类型 |
|---|---|---|
| 申请类型 | 豌豆商城豌豆币添加 | radio |
| 申请原因 | 海外益智学情积分（预习课后习题消课）| 文本 |
| 科目 | 益智 | radio |
| 是否合同赠送 | 否 | radio |
| 虚拟币类型 | 豌豆币 | radio |
| 活动批准审批单号 | 无 | 文本 |
| 用户ID | 详情见附件 | 文本 |
| **积分成本归属部门** | 欢乐童年_海外直播业务线_海外业务中心_海外业务运营处_海外教学服务运营组 | **K2 inputselect autocomplete** |
| 需要添加的豌豆币/魔力币数总共 | 由 `积分汇总.xlsx` 计算 | 文本 |
| 附件 | `发放豌豆币文档填写模板_YYYYMMDD.xlsx` | 文件 |

### 「积分成本归属部门」自动化关键

这是个 K2 `inputselectsgl` autocomplete 控件，不是普通 input，**靠录制用户真实操作得出可行路径**：

1. 点 `.inputselectsgl` 占位 DIV 让控件展开
2. 找展开后的 `input.mp_input:visible` 并 click 聚焦
3. `inp_loc.press_sequentially("海外教学服务运营组", delay=60)` 输入**搜索关键词**（不要用完整带下划线的全路径，K2 不分词）
4. evaluate dispatch `input/keyup/keydown/change` 强制让 K2 刷新候选下拉
5. 等 2.5s 让 ajax 候选回来
6. 按 Enter 让 K2 选中 highlighted 项
7. 验证 chip 出现在 `ol.mf_list li.mf_item`

填写顺序按表单从上到下，部门字段已并入 `FIELD_VALUES`，不再特殊优先。

注意：`scripts/oa_login/auth_state.json` 含会话 Cookie，**不要 git commit、不要回显内容**。如失效，重跑 `python scripts/oa_login/login_oa_qr.py` 扫码刷新。

## 核算规则

### 发放条件（4个条件必须全部满足）

1. **在课包池中**：(学生ID, 课时包ID) 在累计课包池中
2. **是否预习 = 1**：学员完成了预习
3. **线上作业提交状态 = 已提交**：学员提交了作业
4. **消耗课时 >= 1**：基础课时消耗 + 赠送课时消耗 >= 1

### 积分计算公式

```
发放积分 = (基础课时消耗 + 赠送课时消耗) × 500
```

## 输出文件

输出到 `03_output/YYYYMMDD-YYYYMMDD学情积分发放明细/` 文件夹：

| 文件 | 内容 |
|---|---|
| 积分汇总.xlsx | 汇总 / 获得积分明细 / 学情课包ID池 三个 sheet |
| 上课明细_带标注.xlsx | 完整上课明细 + 是否学情课包 / 是否符合发放条件 / 发放积分数量 |
| 续费规划表_学情筛选.xlsx | 本期学情课包的续费规划数据 |
| **发放豌豆币文档填写模板_YYYYMMDD.xlsx** | OA 上传附件，文件名带生成日期 |

## 目录结构

```
学情积分核算/
├── xueqing_credit_skill.py       # Skill 主入口
├── SKILL.md                      # 本文档
├── scripts/
│   ├── run.py                    # 一键运行（取数 → 处理 → 填模板）
│   ├── fetch_reports.py          # BI 取数
│   ├── process_xueqing.py        # 数据处理
│   ├── fill_wandou_template.py   # 填发放模板（自动加日期后缀）
│   ├── setup_scheduled_task.ps1  # 定时任务配置
│   └── oa_login/
│       ├── login_oa_qr.py        # 扫码登录（生成 auth_state.json）
│       ├── submit_oa.py          # 自动提交 OA（Playwright）
│       ├── record_dept_flow.py   # 录制工具（debug 用，记录用户在 K2 表单的真实操作）
│       └── auth_state.json       # 登录态（敏感，勿提交 git）
├── 01_bi_exports/                # BI 原始报表
├── 03_output/                    # 输出
│   ├── 发放豌豆币文档填写模板.xlsx       # 原始空模板（拷贝源）
│   └── YYYYMMDD-YYYYMMDD学情积分发放明细/
│       ├── 积分汇总.xlsx
│       ├── 上课明细_带标注.xlsx
│       ├── 续费规划表_学情筛选.xlsx
│       └── 发放豌豆币文档填写模板_YYYYMMDD.xlsx
└── docs/
    └── 流程.md
```

## 高级用法

### 指定历史课包池

```bash
python xueqing_credit_skill.py run --start 2026-05-01 --end 2026-05-15 \
  --pool-sheet "学情课包ID池" \
  --pool-source "上期输出/积分汇总.xlsx"
```

### 管理定时任务

```bash
schtasks /Query /TN "学情积分核算_月初"
schtasks /Run   /TN "学情积分核算_月初"
schtasks /Delete /TN "学情积分核算_月初" /F
```

### OA 调试 / 录制

如果 K2 改了控件行为导致提交失败，用录制脚本采集真实操作：

```bash
python scripts/oa_login/record_dept_flow.py --keyword 海外
# 浏览器自动到表单页，手动操作目标字段，回到终端 Ctrl+C 收集
# 事件落在 scripts/oa_login/record_events.json
```

## 依赖

- Python 3.12+
- openpyxl
- playwright（含 chrome channel：`playwright install chrome` 或用本地 Chrome）
- bi_skill (BI 报表下载工具)

## 常见问题

### Q: OA 提交失败，URL 还是 `method=add`？
A: 说明字段验证未通过。先看 `scripts/oa_login/submit_step3_filled.png` 截图判断哪个字段空。多数是 K2 widget 初始化没就绪 —— 重跑一次通常就过了。

### Q: 部门字段填了但被误识别为「首页」之类？
A: 双列布局陷阱 + autocomplete 没刷新。已修复（`fill_inputselect_field`），靠 `INPUTSELECT_SEARCH["积分成本归属部门"] = "海外教学服务运营组"` 关键词触发。

### Q: 定时任务没执行？
A:
1. 检查注册：`schtasks /Query /TN "学情积分核算_月初"`
2. 任务计划程序 → 查看任务历史
3. 手动触发：`schtasks /Run /TN "学情积分核算_月初"`

### Q: BI 报表下载提示 `extendsion_loginlock_forbid` / "账号在其他地方登录"？
A: SmartBI 限制同账号同时登录。两次连续运行（或 headful 调试 + 正式运行）会冲突。**等 5-10 秒**让会话释放后重试即可。

### Q: 报表显示「下载完成 0 bytes」，但程序没报错？
A: 历史 bug，已修复（v3.1）。原因是当 `rowCount > max_rows` 时浏览器跳过导出但程序不报错。现在会明确 raise `SmartbiBrowserExportError: Export skipped`。如果再遇到，先调大 `max_rows`，或对该报表配置 `split_days` 分段下载。

### Q: 上课明细下载报 `JavaScript heap out of memory`？
A: 浏览器 V8 堆 ~2GB 上限不足以装 9w+ 行数据。**Playwright `args: ["--js-flags=--max-old-space-size=8192"]` 不生效**（实测）。解决方案：在 `configs/smartbi_simple_report_tasks.json` 给报表配 `"split_days": 4`，按 4 天分段下载然后合并。已为「上课明细」默认开启。

### Q: 合并后的 xlsx 只有几 KB，数据全没了？
A: openpyxl `read_only=True` 模式读 SmartBI 导出的 xlsx 时 `iter_rows()` 会返回空。`merge_xlsx_files` 已改为非 read_only 模式（v3.1）。同时增加合并后大小校验：合并文件 < 段文件总和 30% 会 raise 并保留段文件，方便人工排查/重合并。

### Q: 处理时报 `PermissionError: ...上课明细_带标注.xlsx`？
A: 输出文件被 Excel 占用。看 `01_bi_exports/` 或 `03_output/<期次>/` 是否有 `~$xxx.xlsx` 临时锁文件 —— 关闭对应 Excel 窗口再跑。`run.py` 的 `find_latest_xlsx` 已自动跳过 `~$` 文件。

### Q: 数据处理结果是 0 行 / 0 积分，但 BI 报表是有数据的？
A: 大概率是 `01_bi_exports/` 里有同名旧文件被错用。**v3.1 起 BI 文件名带时间段后缀**（如 `海外思维学员上课明细_20260516-20260531.xlsx`），`run.py` 也按后缀精确匹配。如果文件名没后缀，删了重跑。

### Q: 提交 OA 时点完提交按钮，脚本报 `TargetClosedError: Target page, context or browser has been closed`？
A: OA 系统在提交成功后会跳转/关闭弹窗，脚本最后等待 5s 验证时浏览器已关闭。**通常意味着提交已成功**，但脚本无法自动确认。处理：
1. 登录 OA → 「我的申请」查看是否有刚提交的豌豆币添加申请
2. 没看到再 `python xueqing_credit_skill.py submit-oa --yes` 重提

## 更新日志

### v3.1 (2026-06-03)
- 🐛 修复 `smartbi_browser_export.py` 在 `rowCount > max_rows` 时静默失败：现在会 raise `Export skipped` 错误
- ✨ BI 导出文件名加时间段后缀（如 `海外思维学员上课明细_20260516-20260531.xlsx`），不同期次互不覆盖
- ✨ `run.py` 的 `find_latest_xlsx` 按时间段后缀精确匹配，并自动跳过 `~$` Excel 锁文件
- ✨ 默认 `max_rows` 从 50000 提到 200000
- ✨ **大报表分段下载并合并**：在 task config 加 `"split_days": N` 即可按 N 天分段下载，合并到单一 xlsx。「上课明细」默认 `split_days: 4`，解决 9w+ 行浏览器 OOM 问题
- ✨ 段文件复用：再跑时已下载且非空的段会跳过下载，只补未完成的段
- ✨ 合并失败保护：合并后大小 < 段文件总和 30% 时 raise 并保留段文件
- 🐛 `merge_xlsx_files` 改为非 read_only 模式打开 SmartBI 导出的 xlsx（read_only 下 iter_rows 返回空）

### v3.0 (2026-05-28)
- ✨ 一条龙：`run --auto --submit-oa` 串起取数 → 核算 → 填模板 → 提交 OA
- ✨ 提交 OA 前自动询问确认（列出期次目录 + 附件 + 积分汇总），`--yes` 跳过
- ✨ 模板文件名加日期后缀：`发放豌豆币文档填写模板_YYYYMMDD.xlsx`
- ✨ `submit_oa.py` 自动找最新期次的最新模板（无硬编码路径）
- ✨ `fill_wandou_template.py` 参数化，被 `run.py` 自动调用
- ✨ 「积分成本归属部门」按真实录制路径处理（K2 inputselectsgl autocomplete：点开 → 输关键词 → dispatch keyup → Enter 选中），靠 `record_dept_flow.py` 录制得出
- 🔧 废弃"部门字段先填"特殊规则，统一按表单从上到下顺序

### v2.1 (2026-05-27)
- ✨ OA 豌豆币添加申请自动提交（`submit-oa` 子命令）
- ✨ 复用扫码登录态（auth_state.json）
- 🐛 修复双列布局下「积分成本归属部门」与「活动批准审批单号」相互覆盖的问题
- 🔧 「积分成本归属部门」改为直接文本输入（不走通讯录对话框）

### v2.0 (2026-05-26)
- ✨ 整合为完整的 skill，统一入口
- ✨ 新增 `setup-schedule` 命令配置定时任务
- ✨ 自动使用上期课包池
- ✨ 输出拆分为3个文件
- 🐛 修复4条件筛选逻辑

### v1.0 (2026-05-25)
- 🎉 初始版本
