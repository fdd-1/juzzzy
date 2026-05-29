# 停课唤醒目标 SKILL — 操作说明

## 这个 SKILL 干什么

每月运营要做的事：把 BI 报表里「历史停课未唤醒」学员筛出来 → 在六一工作台建标签和用户群 → 打企微标签 → 在北极星建外呼任务，让LP跟进。

整套流程被拆成 7 个原子步骤，可以一键跑全流程，也可以从任意中间步骤补跑。

## 触发方式（与 Claude 对话）

跟 Claude 说类似这些话就会触发：

- `拉 6 月停课唤醒目标名单` / `跑一次停课唤醒` / `建停课唤醒目标`
- `停课唤醒外呼任务` / `圈一批停课学员准备外呼`

Claude 会执行：

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py all --month 2026-06
```

`--month YYYY-MM` 同时控制：
- BI 报表的签到时间窗口（动态算 [M-4 月初, M-1 月底]）
- 六一标签 / 用户群命名（`2026年6月海外益智停课学员` / 同名 + `（大账号）`）
- 北极星任务克隆目标月份（任务名末尾 `-6月`）

不传 `--month` 时按当前 Windows 时间的当月。

## 7 个步骤一览

| # | 子命令 | 干什么 | 主要输出 |
|---|---|---|---|
| 1 | `export` | BI 拉「思维停课学员执行明细」→ 筛「历史停课未唤醒」分组 | `output/raw_{今天}.xlsx` |
| 2 | `filter` | 二次筛选：是否停课唤醒=0 + 月初前最近一次签到 ∈ 距今前 1~4 个月 | `output/filtered_{今天}.xlsx` |
| 3 | `query-dadou-id` | 平台中心批量用户查询：学员ID → 豌豆大账号ID 映射 | `output/dadou_mapping_{今天}.xlsx` |
| 4 | `liuyi-tag` | 六一工作台新建 2 个标签（小账号 + 大账号），上传 ID 文件 | `output/tag_ids_{今天}.json` |
| 5 | `liuyi-group` | 六一工作台新建 2 个用户群（复制 16811/16812 模板，替换主标签） | `output/group_ids_{今天}.json` |
| 6 | `wechat-tag` | 把大账号用户群挂到企微「【益智】长期标签」标签组 | `output/wechat_tag_{今天}.json` |
| 7 | `polaris-task` | 北极星克隆「【海外】停课120天以内-N月」任务到目标月份 | `output/polaris_task_{今天}.json` |

## 一键跑全流程

```bash
# 默认按 Windows 当月
python ~/Desktop/停课唤醒目标/tingke_wakeup.py all

# 指定目标月份
python ~/Desktop/停课唤醒目标/tingke_wakeup.py all --month 2026-06
```

任一步骤失败会停下，保留前面的产物，修好之后**从失败那一步单独跑**就能续上（每步都会读上一步的产物）。

## 单步骤跑

```bash
python tingke_wakeup.py export --month 2026-06
python tingke_wakeup.py filter --month 2026-06
python tingke_wakeup.py query-dadou-id
python tingke_wakeup.py liuyi-tag --month 2026-06
python tingke_wakeup.py liuyi-group --month 2026-06
python tingke_wakeup.py wechat-tag
python tingke_wakeup.py polaris-task --target-month 6
```

每步都支持 `--help` 看完整参数。

`--dry-run` 在 `liuyi-group` / `wechat-tag` / `polaris-task` 上可用，只打印请求体不真提交。

## 环境配置

### 依赖

```bash
pip install playwright pandas openpyxl
playwright install chromium
```

`bi_skill` 必须存在于 `~/.workbuddy/skills/bi_skill/bi_skill.py`。

### 登录态（首次运行前必做）

三套独立登录态：

```bash
# 六一统一登录平台（覆盖步骤 1-2-3-4-5-6）
python liuyi_login/login_liuyi.py    # 浏览器扫码

# 北极星（覆盖步骤 7）
python polaris_login/login_polaris.py  # 浏览器扫码
```

> BI（步骤 1）走 `bi_skill` 自带的登录态，按 `bi_skill` 自身的说明配置。

登录态文件保存在：
- `liuyi_login/auth_state.json`（六一域，复用于平台中心、六一工作台、企微）
- `polaris_login/auth_state.json`（北极星）

JWT 通常 7~30 天过期。如果某步报「鉴权失败」/「未找到 token」，重跑对应的 login 脚本扫一次码即可。

## 关键约定

### 命名规则

- **六一标签 / 用户群**：`2026年{月}月海外益智停课学员` 和 `2026年{月}月海外益智停课学员（大账号）`
- **北极星任务**：`【海外】停课120天以内-{月}月`（前缀固定，**只换月份数字**，其他字段全部从原任务详情拷贝）
- **企微标签**：用户群名直接作为企微标签名

`{月}` 默认按 Windows 当月，可用 `--month` 显式覆盖。

### 北极星任务的「修改」语义

虽然口语叫「修改任务」，接口实际是 `POST /task/taskTemplate/add` —— 克隆 + 新建一条新月份的任务，**原任务不删**。月底跑下个月版本，是运营的常规节奏。

### 用户群的固定业务标签

大账号用户群（模板 16811）保留 5 个固定运营标签：
- 交集：`8026`、`2160`、`14272`
- 差集：`8025`、`8936`

每月复制时这 5 个不动，只替换主标签 `tagIds[0]` 为本月新建的大账号标签。

## 各步骤产物

### `output/raw_{今天}.xlsx`
BI 导出原始明细，表头在第 10 行（`header_row=9`）。

### `output/filtered_{今天}.xlsx`
二次筛选后的目标人群。「学员id」在第一列，行数即本批次人数。

### `output/dadou_mapping_{今天}.xlsx`
两列：「豌豆账号id」（=学员id）→ 「豌豆大账号id」。来自平台中心批量用户查询导出。

### `output/tag_ids_{今天}.json`
```json
{
  "user_tag":  {"tagId": 19014, "name": "2026年X月海外益智停课学员"},
  "dadou_tag": {"tagId": 19015, "name": "2026年X月海外益智停课学员（大账号）"}
}
```

### `output/group_ids_{今天}.json`
```json
{
  "user_group":  {"name": "...",  "code": "...", "groupId": 17229},
  "dadou_group": {"name": "...",  "code": "...", "groupId": 17230}
}
```

### `output/wechat_tag_{今天}.json`
```json
{
  "wechat_tag_id": 8480,
  "corp_tag_group_id": "etN7IECgAAkW39vv9E__scZlJAnXZFzw",
  "corp_tag_name": "2026年X月海外益智停课学员（大账号）",
  "user_group_id": 17230
}
```

### `output/polaris_task_{今天}.json`
```json
{
  "source_task_id": 3477,
  "source_task_name": "【海外】停课120天以内-5月",
  "new_task_id": 3504,
  "new_task_name": "【海外】停课120天以内-X月"
}
```

## 故障排查

### 「找不到 auth_state.json」 / 「鉴权失败」
跑对应的 login 脚本重新扫码：
```bash
python liuyi_login/login_liuyi.py
python polaris_login/login_polaris.py
```

### `query-dadou-id` 卡在轮询 401
六一登录态过期，重新跑 `login_liuyi.py`。

### `liuyi-tag` 报 `code: 101 系统繁忙`
通常是 multipart 字段对不上。检查上传的 xlsx：
- 表头必须是「用户id」（不是「导入用户id」）
- 数据从 A2 开始，单列数字

### `polaris-task` 报 「Failed to fetch」
浏览器控制台 CORS 拦截。当前实现已经处理（不带 `credentials: 'include'`）。如果还报错，看 `polaris_login/auth_state.json` 是否过期，重新登录。

### 月份匹配不上 / 任务名没改
默认行为：
- `liuyi-tag` / `liuyi-group`：按 Windows 当月命名（可被 `--month` 覆盖）
- `polaris-task`：默认下月（运营节奏：月底为下月准备）

如果用户说「拉 6 月名单」但今天是 5/28 想测试，直接 `--month 2026-06` 覆盖。

### 测试数据如何清理

跑完测试如果产生了不该上线的数据，需要去对应平台手动删除：
- 六一标签：`https://home.61info.cn/` → 标签管理
- 六一用户群：同上 → 用户群
- 企微标签：六一工作台 → 标签管理 → 企微标签
- 北极星任务：`https://sh-center.vipthink.cn/` → 外呼任务

## 项目目录结构

```
停课唤醒目标/
├── tingke_wakeup.py          # 主入口
├── SKILL.md                  # SKILL 元信息（步骤定义）
├── README.md                 # 本文件 — 操作说明
│
├── liuyi_login/              # 六一登录态
│   ├── login_liuyi.py
│   └── auth_state.json
│
├── pingtai_query/            # 步骤 3：平台中心批量用户查询
│   ├── prepare_template.py   #   生成上传 xlsx
│   └── query_dadou.py        #   Playwright 自动化全流程
│
├── liuyi_tag/                # 步骤 4-6：六一工作台 + 企微
│   ├── prepare_csv.py        #   生成两个 ID 文件
│   ├── create_tag.py         #   创建标签 ×2
│   ├── create_group.py       #   创建用户群 ×2
│   └── create_wechat_tag.py  #   挂企微标签 ×1
│
├── polaris_login/            # 北极星登录态
│   ├── login_polaris.py
│   └── auth_state.json
│
├── polaris_task/             # 步骤 7：北极星任务
│   └── update_task.py        #   克隆任务（实为 POST add）
│
└── output/                   # 所有产物（每天一份）
    ├── raw_*.xlsx
    ├── filtered_*.xlsx
    ├── dadou_mapping_*.xlsx
    ├── tag_ids_*.json
    ├── group_ids_*.json
    ├── wechat_tag_*.json
    └── polaris_task_*.json
```
