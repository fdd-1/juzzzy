---
name: tingke-wakeup
description: 停课唤醒目标自动化全流程。从 BI 导出停课学员明细，筛选目标人群，自动在六一工作台新建标签 / 用户群 / 企微标签，最后在北极星外呼平台建立停课唤醒外呼任务。
version: 0.2.0
agent_created: true
status: draft
---

> **技术路线**：Playwright + 真实 Chrome 用户态（`launch_persistent_context`，复用本地 Chrome profile，免登录）。每个步骤的脚本统一输出 `{"ok": true, "data": {...}}` / `{"ok": false, "error": {...}}` 契约（借鉴 x-cli），方便链式调用和断点续跑。

# 停课唤醒目标自动化

将运营每周（或按需）执行的「停课唤醒目标圈选 + 任务下发」操作脚本化，从一份 BI 报表出发，一键完成 标签 → 用户群 → 企微标签 → 外呼任务 的串联。

## 触发方式

用户说以下类似内容时触发，**默认调 `all` 子命令一键跑完整套**：

- "拉 6 月停课唤醒目标名单" / "拉 5 月停课唤醒目标"
- "跑一次停课唤醒" / "建停课唤醒目标"
- "停课唤醒外呼任务" / "圈一批停课学员准备外呼"

**Claude 的行为约定**：
- 听到含明确月份的口令（如「拉 6 月停课唤醒」），执行 `python tingke_wakeup.py all --month 2026-06`
- 没有月份时按 Windows 当月，执行 `python tingke_wakeup.py all`
- 全流程是 7 步：`export → filter → query-dadou-id → liuyi-tag → liuyi-group → wechat-tag → polaris-task`
- 任一步失败就停下，已产生的产物保留，修好后**从失败那步单独跑**就能续上
- 完整使用说明见 [README.md](./README.md)

## 全流程总览

```
┌────────────┐  ┌────────────┐  ┌──────────────┐  ┌────────────┐  ┌────────────┐  ┌────────────┐  ┌────────────┐
│ 1. export  │→ │ 2. filter  │→ │ 3. query-    │→ │ 4. liuyi-  │→ │ 5. liuyi-  │→ │ 6. wechat- │→ │ 7. polaris-│
│  BI 导出   │  │  二次筛选  │  │  dadou-id   │  │  tag       │  │  group     │  │  tag       │  │  task      │
│            │  │            │  │  ID 映射    │  │  建标签×2  │  │  建用户群×2│  │  企微标签  │  │  克隆任务  │
└────────────┘  └────────────┘  └──────────────┘  └────────────┘  └────────────┘  └────────────┘  └────────────┘
     │                │                  │                │                │                │              │
     ▼                ▼                  ▼                ▼                ▼                ▼              ▼
 raw_*.xlsx     filtered_*.xlsx   dadou_mapping_   tag_ids_       group_ids_        wechat_tag_     polaris_task_
                                  *.xlsx           *.json         *.json            *.json          *.json
```

也可以用 `tingke_wakeup.py all --month YYYY-MM` 一键跑完，每步产物落到 `output/`。

---

## 步骤 1：BI 数据导出 ✅ 已确认

**目标**：从 SmartBI 拉取停课学员明细。

| 项 | 内容 |
|---|---|
| 报表名 | **思维停课学员执行明细** |
| 目录树路径 | 海外直播业务线 / 海外后端 / 思维-后端 / 服务 / 停课唤醒 / 思维停课学员执行明细 |
| 筛选项 | 「停课学员分组」= **历史停课未唤醒**（开始/结束日期保持默认不动） |
| 输出 | `Desktop/停课唤醒目标/output/raw_{yyyymmdd}.xlsx` |
| 实现 | 复用 `bi_skill`（Playwright 自动化登录 + 导出） |

> 注意：BI 的「开始日期」「结束日期」不需要修改，保持报表默认值。签到时间的筛选在步骤 2 中通过 Excel 二次过滤完成。

**用法**：

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py export
# 指定月份（默认当月）
python ~/Desktop/停课唤醒目标/tingke_wakeup.py export --month 2026-06
```

---

## 步骤 2：数据筛选（Excel 二次过滤）✅ 已确认

**目标**：在导出的 Excel 中按业务规则筛出本次要外呼的目标人群。

| 项 | 内容 |
|---|---|
| 输入 | 步骤 1 的 `raw_{yyyymmdd}.xlsx`（表头在第 10 行，即 header_row=9） |
| 筛选 1 | 「是否停课唤醒」= 0（实际上「历史停课未唤醒」分组已隐含此条件，全部为 0） |
| 筛选 2 | 「月初前最近一次签到时间」∈ **距今前 1~4 个月** |
| 输出 | `Desktop/停课唤醒目标/output/filtered_{yyyymmdd}.xlsx` |

**「月初前最近一次签到时间」日期窗口规则（动态计算）**：

| 导出当月 | 起始（前 4 月 1 号） | 结束（前 1 月最后一天） |
|---|---|---|
| 5 月 | 1/1 | 4/30 |
| 6 月 | 2/1 | 5/31 |
| 7 月 | 3/1 | 6/30 |

> 即：以导出月 M 为基准，签到时间 ∈ `[M-4 月 1 号, M-1 月最后一天]`（距今前 1~4 个月）。

**用法**：

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py filter
# 指定输入文件
python ~/Desktop/停课唤醒目标/tingke_wakeup.py filter --input path/to/raw.xlsx
# 指定月份
python ~/Desktop/停课唤醒目标/tingke_wakeup.py filter --month 2026-06
```

---

## 步骤 2.5：学员ID → 豌豆大账号ID 映射 ✅ 已实现

**目标**：六一/北极星都用「豌豆大账号 ID」识别用户，但 BI 导出的是「学员 ID（豌豆账号 ID）」，所以要先去 **平台中心 → 用户管理 → 批量用户查询** 拿到映射关系。

| 项 | 内容 |
|---|---|
| 入口 | https://bizcenter-h5-cms.61info.cn/ → 用户管理 → 批量用户查询 |
| 登录 | 复用六一统一登录态 `liuyi_login/auth_state.json` |
| 输入 | `output/filtered_{yyyymmdd}.xlsx` 的「学员id」列 |
| 中间产物 | `pingtai_query/upload_{yyyymmdd_HHMMSS}.xlsx`（A1=「导入用户id」，A2..学员ID） |
| 输出 | `output/dadou_mapping_{yyyymmdd}.xlsx`（两列：豌豆账号id / 豌豆大账号id） |
| 实现 | Playwright 浏览器自动化：新建 → 导入类型=「豌豆用户」→ 上传文件 → 导出类型=「豌豆大账号」→ 保存 → 等 10s → 调 `queryExportRecords` 接口轮询 → 下载 OSS 导出文件 |
| 鉴权 | 接口需 JWT，从浏览器 localStorage 自动取 |

**用法**：

```bash
# 一键跑完
python ~/Desktop/停课唤醒目标/tingke_wakeup.py query-dadou-id

# 指定 filtered xlsx
python ~/Desktop/停课唤醒目标/tingke_wakeup.py query-dadou-id --input output/filtered_20260527.xlsx

# 也可以分步跑
python ~/Desktop/停课唤醒目标/pingtai_query/prepare_template.py --input output/filtered_20260527.xlsx
python ~/Desktop/停课唤醒目标/pingtai_query/query_dadou.py
```

---

## 步骤 2（旧版重复段，已合并到上方）

**目标**：在原始明细基础上按业务规则筛出本次要外呼的目标人群。

| 项 | 内容 |
|---|---|
| 筛选规则 | _待用户提供_（例：停课天数 ≥ 14 且 ≤ 60、未做转介绍、剩余课时 > 0、最近 7 天无 1V1 触达 等） |
| 去重逻辑 | 默认按 `user_id` 去重；如有黑名单（已退费、投诉用户）需排除 |
| 抽样/上限 | 是否限制单次任务人数上限（如最多 500 人） |
| 输出 | `Desktop/停课唤醒/filtered_{yyyymmdd}.xlsx`，至少包含 `user_id, 学员姓名, 手机号, 所属老师, 停课天数, 备注` |

**用法**：

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py --step filter --input raw_20260527.xlsx
```

---

## 步骤 3：六一工作台 — 新建标签 + 新建用户群 ✅ 已实现

**目标**：在六一工作台创建本批次的两个标签 + 两个用户群（小账号版 + 大账号版）。

| 项 | 内容 |
|---|---|
| 入口 | https://home.61info.cn/ → 标签管理 |
| 登录 | 复用六一统一登录态 `liuyi_login/auth_state.json` |
| 实现 | Playwright 进六一工作台 → 浏览器内 `page.evaluate(fetch ...)` 调 `userTag/addwithfile`（multipart）和 `userGroup/add`（JSON） |
| 鉴权 | 接口需 JWT，从浏览器 localStorage 自动取 |

### 3.1 新建标签 ×2 — `liuyi-tag`

| 项 | 内容 |
|---|---|
| 输入 | `output/filtered_{yyyymmdd}.xlsx`（学员id列） + `output/dadou_mapping_{yyyymmdd}.xlsx`（豌豆大账号id列） |
| 命名 | `2026年{月}月海外益智停课学员` 和 `2026年{月}月海外益智停课学员（大账号）`，月份默认按当前月，可用 `--month YYYY-MM` 覆盖 |
| 业务参数 | `bizCode=WANDOU` / `tagType=531`（关键行为/活跃行为）/ `dataFrom=2`（按导入用户id筛选）/ `tagUpdateType=1` / `onceUpdate=1` |
| 文件格式 | xlsx，A 列表头「用户id」（与官方模板 `usertag_template.csv` 一致） |
| 中间产物 | `liuyi_tag/user_ids_{stamp}.xlsx`、`liuyi_tag/dadou_ids_{stamp}.xlsx` |
| 输出 | `output/tag_ids_{yyyymmdd}.json`（小账号 tagId / 大账号 tagId） |

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-tag
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-tag --month 2026-07
```

### 3.2 新建用户群 ×2 — `liuyi-group`

复制 5 月模板（小账号 16812 / 大账号 16811）的 `tagsConfig` 结构，把第一项 `tagIds` 替换成新建的标签 id（19011/19012/...），其余「业务侧固定标签」（8026/2160/14272 交集；8025/8936 差集）原样保留。

| 项 | 内容 |
|---|---|
| 依赖 | `output/tag_ids_{yyyymmdd}.json`（由 `liuyi-tag` 产出） |
| 模板 | 小账号 16812（单标签）/ 大账号 16811（含 5 个固定业务标签） |
| 接口 | `POST userGroup/checkData`（带模板 id 预检人群差异） + `POST userGroup/add`（不带 id 创建） |
| 输出 | `output/group_ids_{yyyymmdd}.json`（小账号 groupId / 大账号 groupId） |

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-group
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-group --month 2026-07
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-group --dry-run   # 只构造请求体打印不调 add
```

---

## 步骤 3：六一工作台 — 新建标签 + 新建用户群（待实现旧段，已合并到上方）

**目标**：把筛选后的目标学员灌进六一工作台，生成本次唤醒批次的标签和用户群。

### 3.1 新建标签

| 项 | 内容 |
|---|---|
| 入口 URL | _待用户提供_ |
| 登录方式 | _待用户提供_ |
| 标签命名规则 | 建议：`停课唤醒_YYYYMMDD` |
| 必填字段 | _待用户确认_ |
| 输出 | `tag_id` |

### 3.2 新建用户群

| 项 | 内容 |
|---|---|
| 入口路径 | _待用户提供_ |
| 群命名规则 | 建议：`停课唤醒_YYYYMMDD` |
| 人员来源 | 筛选后的 `学员id` 列表 |
| 输出 | `group_id` |

---

## 步骤 4：新建企微标签 ✅ 已实现

**目标**：把上一步创建的「大账号用户群」关联到企微「【益智】长期标签」标签组下，让企业微信侧能拉到这批人群打标签触达。注意：**只关联大账号**，小账号不需要。

| 项 | 内容 |
|---|---|
| 入口 | 六一工作台 → 标签管理 → 企微标签 |
| 接口 | `POST /corporate-wechat-backend/o/v1/tagGroup/create` |
| 依赖 | `output/group_ids_{yyyymmdd}.json`（取 dadou_group） |
| 业务参数 | `bizCode=WANDOU`，`corpTagGroupId=etN7IECgAAkW39vv9E__scZlJAnXZFzw`（「【益智】长期标签」固定 id） |
| 输出 | `output/wechat_tag_{yyyymmdd}.json`（企微 tag id） |

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py wechat-tag
python ~/Desktop/停课唤醒目标/tingke_wakeup.py wechat-tag --dry-run
```

> 命名要点：步骤 3 标签和用户群都按 **Windows 当月** 命名（`create_tag.py` / `create_group.py` 已强制锁死，`--month` 参数被忽略）。

---

## 步骤 4：新建企微标签（待实现旧段，已合并到上方）

| 项 | 内容 |
|---|---|
| 入口 | _待用户提供_ |
| 标签结构 | _待用户提供_ |
| 输出 | `wechat_tag_id` |

---

## 步骤 5：北极星外呼平台 — 克隆「停课120天以内」任务到目标月份 ✅ 已实现

**目标**：在北极星把既有的「【海外】停课120天以内-N月」外呼任务模板克隆出一个新月份的版本，让客服坐席继续跟进。

| 项 | 内容 |
|---|---|
| 入口 | https://passport.vipthink.cn/#/account/login → https://sh-center.vipthink.cn/#/ |
| 登录 | 钉钉扫码（独立登录态 `polaris_login/auth_state.json`，不复用六一登录态） |
| 接口 | `POST /task/taskTemplate/16/list`（搜任务）+ `GET /task/taskTemplate/getDetail?id={id}`（拉详情）+ `POST /task/taskTemplate/add`（克隆新建） |
| 鉴权 | `Authorization: Bearer eyJ...`，从浏览器 localStorage / sessionStorage 自动找（兼容 `Bearer xxx` 或纯 `eyJxxx` 两种存储） |

> 注：用户口径是「修改」，但接口实际是 **克隆 + 新建**（POST add），原任务保留不删。前缀「【海外】停课120天以内」固定，**只改末尾月份数字**，其他字段全部从原任务详情拷贝。

**默认目标月份 = 下月**（运营节奏：月底为下月准备）。可用 `--target-month 1-12` 覆盖。

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py polaris-task           # 默认下月
python ~/Desktop/停课唤醒目标/tingke_wakeup.py polaris-task --target-month 7
python ~/Desktop/停课唤醒目标/tingke_wakeup.py polaris-task --dry-run  # 只构造 payload 不调 add
```

输出：`output/polaris_task_{yyyymmdd}.json`（包含 `source_task_id` / `new_task_id` / 新任务名）

---

## 步骤 5：北极星外呼平台 — 建立停课唤醒任务（待实现旧段，已合并到上方）

| 项 | 内容 |
|---|---|
| 入口 URL | _待用户提供_ |
| 登录方式 | _待用户提供_ |
| 任务类型 | _待用户确认_ |
| 必填字段 | _待用户确认_ |
| 输出 | `task_id` |

---

## 命令行总入口

```bash
# 端到端一把梭
python ~/Desktop/停课唤醒目标/tingke_wakeup.py all

# 只跑某一步（断点续跑）
python ~/Desktop/停课唤醒目标/tingke_wakeup.py export
python ~/Desktop/停课唤醒目标/tingke_wakeup.py filter
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-tag
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-group
python ~/Desktop/停课唤醒目标/tingke_wakeup.py wechat-tag
python ~/Desktop/停课唤醒目标/tingke_wakeup.py polaris-task

# 干跑（不真正提交，只演练流程并截图）
python ~/Desktop/停课唤醒目标/tingke_wakeup.py export --dry-run
```

## 失败重试 & 日志

- 每一步成功后把 `tag_id / group_id / wechat_tag_id / task_id` 落到 `run_log_*.json`
- 任意一步失败：截图 + 写日志 + 退出码 ≠ 0；下次带 `--step` 从断点续跑
- 所有 Playwright 操作默认 `headless=False` 便于排错，跑稳后切 `--headless`

## 依赖

- Python 3.10+
- pandas / openpyxl
- playwright（含 chromium）
- requests（如某些平台有 OpenAPI 可用，优先 API 而非浏览器）

## 待办（需用户提供，详见对话）

- [x] BI 报表名 + 业务筛选条件（思维停课学员执行明细，已确认）
- [ ] 数据筛选规则 + 单次人数上限
- [ ] 六一工作台 URL / 登录方式 / 新建标签&用户群的菜单路径
- [ ] 企微标签的入口和命名规范
- [ ] 北极星外呼平台 URL / 登录方式 / 任务模板
- [ ] 是否有现成的 OpenAPI 可绕过浏览器
