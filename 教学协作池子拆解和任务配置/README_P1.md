# 教学协作 P1 服务池协作自动化

> 创建日期：2026-06-07
> 
> 基于 P0 学位预警流程改造，专门用于每月教学协作池 P1 服务池协作任务自动化。

---

## 一、快速开始

### 一键执行（推荐）

```bash
python fuwuchi_auto.py all --month 2026-06 --teacher-complete-time "2026-06-30 23:59:59"
```

参数说明：
- `--month`：目标月份（YYYY-MM 格式，例：2026-06）
- `--teacher-complete-time`：教师任务完成时间（YYYY-MM-DD HH:mm:ss 格式）

### 单步执行

当需要调试或部分执行时：

```bash
# 1. 导出报表（复用P0的export.py）
python fuwuchi_auto.py export --month 2026-06

# 2. 筛选（手动传入导出的文件路径）
python fuwuchi_auto.py filter --input exports/学位预警_20260607/海外思维续费规划表_新版_26年启用_20260607_143022.xlsx

# 3. 创建标签
python fuwuchi_auto.py tag --month 2026-06

# 4. 创建用户群
python fuwuchi_auto.py group --month 2026-06

# 5. 克隆北极星任务
python fuwuchi_auto.py polaris --target-month 6 --teacher-complete-time "2026-06-30 23:59:59"

# 6. 标签数据同步
python fuwuchi_auto.py sync --month 2026-06
```

---

## 二、P1任务完整流程

### 2.1 报表导出（export）

**目标**：从 SmartBI 导出「海外思维续费规划表_新版_26年启用」

**自动填充筛选项**：
- 开课M计算时间 = 当月1号（例：6月 → 2026-06-01）
- 退费结束时间 = 上月最后一天 23:59:59（例：6月 → 2026-05-31 23:59:59）
- 学管大区 = 海外教学服务部
- 池子节点3 = 服务月

**输出**：`exports/学位预警_{YYYYMMDD}/海外思维续费规划表_新版_26年启用_{YYYYMMDD}_{HHMMSS}.xlsx`

**注意**：P1 复用 P0 的 export.py，报表导出逻辑完全相同

---

### 2.2 二次筛选（filter_p1.py）

**筛选条件（与P0的关键差异）**：
- `是否可续学员` == 1
- `月初是否续费` == 空（NaN）
- ⚠️ **P1 不筛选「月初剩余总课时」**（P0 筛选 1-12 课时）

**输出**：
- `output/p1_dadou_ids_{YYYYMMDD}.xlsx` → 豌豆大账号ID（普通标签/用户群用）
- `output/p1_user_ids_{YYYYMMDD}.xlsx` → 学员ID（益智标签/用户群用）

---

### 2.3 创建标签（create_tag_p1.py）

**标签命名规则（与P0的差异）**：
- 普通标签：`【海外】26年X月教学协作池-X月服务池`
- 益智标签：`【海外】26年X月教学协作池-X月服务池（益智）`

**上传ID类型**：
- 普通标签 ← 豌豆大账号ID（p1_dadou_ids_*.xlsx）
- 益智标签 ← 学员ID（p1_user_ids_*.xlsx）

**业务参数**：
- bizCode = WANDOU
- tagType = 531（关键行为）
- dataFrom = 2（按导入用户ID筛选）

**输出**：`output/p1_tag_ids_{YYYYMMDD}.json`

---

### 2.4 创建用户群（create_group_p1.py）

**用户群命名规则**：
- 普通用户群：`【海外】26年X月教学协作池-X月服务池`
- 益智用户群：`【海外】26年X月教学协作池-X月服务池（益智）`

**模板**：
- 普通用户群 ← 模板 16811（豌豆大账号模板）
- 益智用户群 ← 模板 16812（学员ID小账号模板）

**输出**：`output/p1_group_ids_{YYYYMMDD}.json`

---

### 2.5 北极星任务（update_task_p1.py）

**操作方式**：复制历史任务并修改

**搜索关键词**：`P1（学情反馈）`（P0 搜索「P0（学位预警）」）

**任务名称**：`【海外】26年X月教学协作池-X月服务池`

**修改字段**：
- ⚠️ 任务名称：末尾月份改为目标月份
- ⚠️ 用户群：挂普通用户群（不挂益智群）
- ⚠️ 教师任务完成时间：手动传入（格式 YYYY-MM-DD HH:mm:ss）
- 🔒 其他字段：全部保持原样，不做任何修改

**输出**：`output/p1_polaris_task_{YYYYMMDD}.json`

---

### 2.6 标签数据同步（sync_tag_data_p1.py）

**配置项**：
- 业务类型：豌豆
- 同步业务系统：豌豆数合表
- 同步用户群：**益智群**（不是普通群）
- 同步数据频率：每天
- 状态：启用

**操作**：
1. 访问用户标签页面刷新
2. 访问用户群页面刷新
3. 回到标签数据同步页面
4. 点击新增按钮
5. 填写表单
6. 点击确认
7. 点击手动同步

**输出**：`output/p1_sync_tag_final_{YYYYMMDD}.png`（截图）

---

## 三、P0 vs P1 对比

| 维度 | P0（学位预警） | P1（服务池协作） |
|------|--------------|----------------|
| **报表导出** | SmartBI同一张表 | SmartBI同一张表 |
| **筛选条件** | 课时1-12 + 可续 + 未续费 | 可续 + 未续费（不限课时）|
| **标签命名** | 【海外】26年X月教学协作池-1-12课时低活 | 【海外】26年X月教学协作池-X月服务池 |
| **用户群命名** | 同标签 | 同标签 |
| **北极星搜索** | P0（学位预警） | 服务池 |
| **北极星任务名** | 【海外】26年X月教学协作池-P0（学位预警） | 【海外】26年X月教学协作池-X月服务池 |

---

## 四、前置准备

### 4.1 首次使用

P1 任务完全复用 P0 的登录态和环境配置：

1. **登录六一工作台**（一次即可，后续复用）

```bash
python liuyi_login/login_liuyi.py
```

登录成功后会保存 `liuyi_login/auth_state.json`。

2. **登录北极星外呼平台**（一次即可，后续复用）

```bash
python polaris_login/login_polaris.py
```

登录成功后会保存 `polaris_login/auth_state.json`。

3. **配置 SmartBI 凭证**（环境变量）

```bash
# Windows PowerShell
$env:SMARTBI_USERNAME = "你的用户名"
$env:SMARTBI_PASSWORD = "你的密码"

# 或者写入系统环境变量（永久生效）
```

---

### 4.2 目录结构

```
教学协作池子拆解和任务配置/
├── README.md                              # P0 任务文档
├── README_P1.md                           # P1 任务文档（本文档）
├── PROCESS_RECORD.md                      # 完整流程记录
├── 教学协作池子拆解和任务配置.md              # 原需求文档
├── xuewei_warning.py                      # P0 顶层入口
├── fuwuchi_auto.py                        # P1 顶层入口
├── export.py                              # 报表导出（P0/P1 共用）
├── filter.py                              # P0 二次筛选
├── filter_p1.py                           # P1 二次筛选
├── smartbi_tasks.json                     # SmartBI 报表配置
├── liuyi_login/                           # 六一工作台登录（P0/P1 共用）
│   ├── login_liuyi.py
│   └── auth_state.json
├── liuyi_tag/                             # 标签/用户群/同步
│   ├── create_tag.py                      # P0 标签创建
│   ├── create_tag_p1.py                   # P1 标签创建
│   ├── create_group.py                    # P0 用户群创建
│   ├── create_group_p1.py                 # P1 用户群创建
│   ├── sync_tag_data.py                   # P0 标签同步
│   ├── sync_tag_data_p1.py                # P1 标签同步
│   ├── latest_inputs.json                 # P0 筛选产物清单
│   └── latest_inputs_p1.json              # P1 筛选产物清单
├── polaris_login/                         # 北极星登录（P0/P1 共用）
│   ├── login_polaris.py
│   └── auth_state.json
├── polaris_task/                          # 北极星任务克隆
│   ├── update_task.py                     # P0 任务克隆
│   └── update_task_p1.py                  # P1 任务克隆
├── templates/                             # 标签上传模板
│   ├── usertag_template-标签模板（豌豆大账号）.xlsx
│   └── usertag_template-标签模板（豌豆账号）.xlsx
├── exports/                               # 报表导出目录
│   └── 学位预警_{YYYYMMDD}/
└── output/                                # 产物目录
    ├── p1_dadou_ids_{YYYYMMDD}.xlsx
    ├── p1_user_ids_{YYYYMMDD}.xlsx
    ├── p1_tag_ids_{YYYYMMDD}.json
    ├── p1_group_ids_{YYYYMMDD}.json
    ├── p1_polaris_task_{YYYYMMDD}.json
    └── p1_sync_tag_final_{YYYYMMDD}.png
```

---

## 五、注意事项

1. **标签命名不含课时**  
   P1 标签名为「X月服务池」，不含「1-12课时低活」字样。

2. **筛选不限课时**  
   P1 筛选条件只看「可续」+「未续费」，不限制月初剩余总课时。

3. **教师任务完成时间**  
   每次手动传入，格式 `YYYY-MM-DD HH:mm:ss`，例如 `2026-06-30 23:59:59`。

4. **同步用户群选益智群**  
   标签数据同步时选择「益智群」（带"（益智）"后缀），不是普通群。

5. **复用 P0 基础设施**  
   P1 复用 P0 的登录模块、报表导出、模板文件，无需重复配置。

---

## 六、常见问题

**Q1: P1 和 P0 可以同时执行吗？**  
→ 可以，两个任务产物文件名不同（p1_ 前缀），不会冲突。

**Q2: P1 的北极星任务搜索关键词是什么？**  
→ 「服务池」，与 P0 的「P0（学位预警）」不同。

**Q3: P1 筛选结果比 P0 多很多正常吗？**  
→ 正常，P1 不限制课时，所有未续费的可续学员都会被筛选出来。

**Q4: P1 标签名称格式？**  
→ `【海外】26年6月教学协作池-6月服务池` / `【海外】26年6月教学协作池-6月服务池（益智）`

---

## 七、更新日志

| 日期 | 操作 | 说明 |
| --- | --- | --- |
| 2026-06-07 | 创建 P1 自动化脚本 | 基于 P0 流程改造，适配服务池协作需求 |
| 2026-06-07 | 完成流程记录文档 | 记录 P0/P1 完整流程到 PROCESS_RECORD.md |
