# 教学协作 P0 学位预警自动化

> 创建日期：2026-06-05
> 
> 基于停课唤醒目标流程搭建，专门用于每月教学协作池 P0 学位预警任务自动化。

## 一、快速开始

### 一键执行（推荐）

```bash
python xuewei_warning.py all --month 2026-06 --teacher-complete-time "2026-06-30 23:59:59"
```

参数说明：
- `--month`：目标月份（YYYY-MM 格式，例：2026-06）
- `--teacher-complete-time`：教师任务完成时间（YYYY-MM-DD HH:mm:ss 格式）

### 单步执行

当需要调试或部分执行时：

```bash
# 1. 导出报表
python xuewei_warning.py export --month 2026-06

# 2. 筛选（手动传入导出的文件路径）
python xuewei_warning.py filter --input exports/学位预警_20260605/海外思维续费规划表_新版_26年启用_20260605_143022.xlsx

# 3. 创建标签
python xuewei_warning.py tag --month 2026-06

# 4. 创建用户群
python xuewei_warning.py group --month 2026-06

# 5. 克隆北极星任务
python xuewei_warning.py polaris --target-month 6 --teacher-complete-time "2026-06-30 23:59:59"

# 6. 标签数据同步
python xuewei_warning.py sync --month 2026-06
```

---

## 二、完整流程

### 2.1 报表导出（export）

**目标**：从 SmartBI 导出「海外思维续费规划表_新版_26年启用」

**自动填充筛选项**：
- 开课M计算时间 = 当月1号（例：6月 → 2026-06-01）
- 退费结束时间 = 上月最后一天 23:59:59（例：6月 → 2026-05-31 23:59:59）
- 学管大区 = 海外教学服务部（默认）

**输出**：`exports/学位预警_{YYYYMMDD}/海外思维续费规划表_新版_26年启用_{YYYYMMDD}_{HHMMSS}.xlsx`

---

### 2.2 二次筛选（filter）

**筛选条件**：
- `月初剩余总课时` in [1, 2, 3, ..., 12]
- `是否可续学员` == 1
- `月初是否续费` == 空（NaN）

**输出**：
- `output/dadou_ids_{YYYYMMDD}.xlsx` → 豌豆大账号ID（普通标签/用户群用）
- `output/user_ids_{YYYYMMDD}.xlsx` → 学员ID（益智标签/用户群用）

---

### 2.3 创建标签（tag）

**标签命名规则**：
- 普通标签：`【海外】26年X月教学协作池-1-12课时低活`
- 益智标签：`【海外】26年X月教学协作池-1-12课时低活（益智）`

**上传ID类型**：
- 普通标签 ← 豌豆大账号ID（dadou_ids_*.xlsx）
- 益智标签 ← 学员ID（user_ids_*.xlsx）

**业务参数**：
- bizCode = WANDOU
- tagType = 531（关键行为）
- dataFrom = 2（按导入用户ID筛选）

**输出**：`output/tag_ids_{YYYYMMDD}.json`

---

### 2.4 创建用户群（group）

**用户群命名规则**：
- 普通用户群：`【海外】26年X月教学协作池-1-12课时低活`
- 益智用户群：`【海外】26年X月教学协作池-1-12课时低活（益智）`

**模板**：
- 普通用户群 ← 模板 16811（豌豆大账号模板）
- 益智用户群 ← 模板 16812（学员ID小账号模板）

**输出**：`output/group_ids_{YYYYMMDD}.json`

---

### 2.5 北极星任务（polaris）

**操作方式**：复制历史任务并修改

**搜索关键词**：`P0（学位预警）`

**任务名称**：`【海外】26年X月教学协作池-P0（学位预警）`

**修改字段**：
- ⚠️ 任务名称：末尾月份改为目标月份
- ⚠️ 用户群：挂普通用户群（不挂益智群）
- ⚠️ 教师任务完成时间：手动传入（格式 YYYY-MM-DD HH:mm:ss）
- 🔒 其他字段：全部保持原样，不做任何修改

**输出**：`output/polaris_task_{YYYYMMDD}.json`

---

### 2.6 标签数据同步（sync）

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

**输出**：`output/sync_tag_final_{YYYYMMDD}.png`（截图）

---

## 三、前置准备

### 3.1 首次使用

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

### 3.2 目录结构

```
教学协作池子拆解和任务配置/
├── README.md                              # 本文档
├── SKILL.md                               # skill 元信息
├── 教学协作池子拆解和任务配置.md              # 原需求文档
├── xuewei_warning.py                      # 顶层入口（一键 / 单步）
├── export.py                              # 导出报表
├── filter.py                              # 二次筛选
├── smartbi_tasks.json                     # SmartBI 报表配置
├── liuyi_login/                           # 六一工作台登录
│   ├── login_liuyi.py
│   └── auth_state.json                    # 登录态（首次运行生成）
├── liuyi_tag/                             # 标签 / 用户群 / 同步
│   ├── create_tag.py
│   ├── create_group.py
│   ├── sync_tag_data.py
│   └── latest_inputs.json                 # filter 产物清单
├── polaris_login/                         # 北极星登录
│   ├── login_polaris.py
│   └── auth_state.json                    # 登录态（首次运行生成）
├── polaris_task/                          # 北极星任务克隆
│   └── update_task.py
├── templates/                             # 标签上传模板
│   ├── usertag_template-标签模板（豌豆大账号）.xlsx
│   └── usertag_template-标签模板（豌豆账号）.xlsx
├── exports/                               # 报表导出目录
│   └── 学位预警_{YYYYMMDD}/
└── output/                                # 产物目录
    ├── dadou_ids_{YYYYMMDD}.xlsx
    ├── user_ids_{YYYYMMDD}.xlsx
    ├── tag_ids_{YYYYMMDD}.json
    ├── group_ids_{YYYYMMDD}.json
    ├── polaris_task_{YYYYMMDD}.json
    └── sync_tag_final_{YYYYMMDD}.png
```

---

## 四、注意事项

1. **标签命名加月份**  
   所有标签、用户群、北极星任务名称都按月份命名，例如"26年6月"。

2. **教师任务完成时间**  
   每次手动传入，格式 `YYYY-MM-DD HH:mm:ss`，例如 `2026-06-30 23:59:59`。

3. **同步用户群选益智群**  
   标签数据同步时选择"益智群"（带"（益智）"后缀），不是普通群。

4. **不挂定时任务**  
   本流程为手动触发，不设置定时任务。

5. **SmartBI 凭证安全**  
   环境变量 `SMARTBI_USERNAME` 和 `SMARTBI_PASSWORD` 仅在 export 步骤使用，不会写入日志或文件。

---

## 五、常见问题

**Q1: 报表导出失败？**  
→ 检查 SmartBI 凭证是否配置正确（环境变量），以及网络连接。

**Q2: 标签创建失败？**  
→ 确认 `liuyi_login/auth_state.json` 是否有效（有效期约 7 天），过期需重新登录。

**Q3: 北极星任务克隆失败？**  
→ 确认 `polaris_login/auth_state.json` 是否有效，以及搜索关键词"P0（学位预警）"能否匹配到历史任务。

**Q4: 标签数据同步无法选到用户群？**  
→ 同步前需先访问"用户标签页面"和"用户群页面"刷新，脚本已自动执行此步骤。

---

## 六、更新日志

| 日期 | 操作 | 说明 |
| --- | --- | --- |
| 2026-06-05 | 创建 skill 骨架 | 参照停课唤醒流程搭建，适配教学协作 P0 学位预警需求 |
