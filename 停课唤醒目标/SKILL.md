---
name: tingke-wakeup
description: 停课唤醒目标自动化全流程。从 BI 导出停课学员明细，筛选目标人群，自动在六一工作台新建标签 / 用户群 / 企微标签，在北极星外呼平台建立停课唤醒外呼任务，最后配置标签数据同步。
version: 0.4.0
agent_created: true
status: stable
---

> **技术路线**：Playwright + 真实 Chrome 用户态（`launch_persistent_context`，复用本地 Chrome profile，免登录）。每个步骤的脚本统一输出 `{"ok": true, "data": {...}}` / `{"ok": false, "error": {...}}` 契约，方便链式调用和断点续跑。

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
- 全流程是 8 步：`export → filter → query-dadou-id → liuyi-tag → liuyi-group → wechat-tag → polaris-task → sync-tag-data`
- 任一步失败就停下，已产生的产物保留，修好后**从失败那步单独跑**就能续上
- 完整使用说明见 [README.md](./README.md)

**⚠️ 首次使用必读**：
- 第一次运行时，务必先执行 `python tingke_wakeup.py all --dry-run`
- 检查所有参数是否正确（月份、文件路径、接口响应）
- 确认无误后，去掉 `--dry-run` 正式运行

## 全流程总览

```
┌────────────┐  ┌────────────┐  ┌──────────────┐  ┌────────────┐  ┌────────────┐  ┌────────────┐  ┌────────────┐  ┌──────────────┐
│ 1. export  │→ │ 2. filter  │→ │ 3. query-    │→ │ 4. liuyi-  │→ │ 5. liuyi-  │→ │ 6. wechat- │→ │ 7. polaris-│→ │ 8. sync-tag- │
│  BI 导出   │  │  二次筛选  │  │  dadou-id   │  │  tag       │  │  group     │  │  tag       │  │  task      │  │  data        │
│            │  │            │  │  ID 映射    │  │  建标签×2  │  │  建用户群×2│  │  企微标签  │  │  克隆任务  │  │  数据同步    │
└────────────┘  └────────────┘  └──────────────┘  └────────────┘  └────────────┘  └────────────┘  └────────────┘  └──────────────┘
     │                │                  │                │                │                │              │                │
     ▼                ▼                  ▼                ▼                ▼                ▼              ▼                ▼
 raw_*.xlsx     filtered_*.xlsx   dadou_mapping_   tag_ids_       group_ids_        wechat_tag_     polaris_task_   sync_tag_
                                  *.xlsx           *.json         *.json            *.json          *.json          done_*.png
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
| **失败场景** | **⚠️ 如果下载超时** → 重试 3 次 → 失败则记录日志并终止 |
| | **⚠️ 如果文件大小 < 1KB** → 视为无效，报错并终止 |
| | **⚠️ 如果登录态失效** → 提示运行 `bi_skill` 重新登录 |

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

🔴 **CHECKPOINT 1：筛选前确认**
- 导出月份是否正确？（显示在文件名中）
- raw 文件行数是否 > 0？

| 项 | 内容 |
|---|---|
| 输入 | 步骤 1 的 `raw_{yyyymmdd}.xlsx`（表头在第 10 行，即 header_row=9） |
| 筛选 1 | 「是否停课唤醒」= 0（实际上「历史停课未唤醒」分组已隐含此条件，全部为 0） |
| 筛选 2 | 「月初前最近一次签到时间」∈ **距今前 1~4 个月** |
| 输出 | `Desktop/停课唤醒目标/output/filtered_{yyyymmdd}.xlsx` |
| **失败场景** | **⚠️ 如果输入文件不存在** → 检查步骤 1 是否执行成功 |
| | **⚠️ 如果筛选后人数为 0** → 检查日期窗口规则是否正确 |
| | **⚠️ 如果筛选后人数 > 5000** → 停下来检查筛选逻辑是否异常 |

**「月初前最近一次签到时间」日期窗口规则（动态计算）**：

| 导出当月 | 起始（前 4 月 1 号） | 结束（前 1 月最后一天） |
|---|---|---|
| 5 月 | 1/1 | 4/30 |
| 6 月 | 2/1 | 5/31 |
| 7 月 | 3/1 | 6/30 |

> 即：以导出月 M 为基准，签到时间 ∈ `[M-4 月 1 号, M-1 月最后一天]`（距今前 1~4 个月）。

**筛选后验证**：
- 筛选后人数是否在合理范围（100-2000）？

**用法**：

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py filter
# 指定输入文件
python ~/Desktop/停课唤醒目标/tingke_wakeup.py filter --input path/to/raw.xlsx
# 指定月份
python ~/Desktop/停课唤醒目标/tingke_wakeup.py filter --month 2026-06
```

---

## 步骤 3：学员ID → 豌豆大账号ID 映射 ✅ 已实现

**目标**：六一/北极星都用「豌豆大账号 ID」识别用户，但 BI 导出的是「学员 ID（豌豆账号 ID）」，所以要先去 **平台中心 → 用户管理 → 批量用户查询** 拿到映射关系。

| 项 | 内容 |
|---|---|
| 入口 | https://bizcenter-h5-cms.61info.cn/ → 用户管理 → 批量用户查询 |
| 登录 | 复用六一统一登录态 `liuyi_login/auth_state.json` |
| 输入 | `output/filtered_{yyyymmdd}.xlsx` 的「学员id」列 |
| 中间产物 | `pingtai_query/upload_{yyyymmdd_HHMMSS}.xlsx`（A1=「导入用户id」，A2..学员ID） |
| 输出 | `output/dadou_mapping_{yyyymmdd}.xlsx`（两列：豌豆账号id / 豌豆大账号id） |
| 实现 | Playwright 浏览器自动化：新建 → 导入类型=「豌豆用户」→ 上传文件 → 导出类型=「豌豆大账号」→ 保存 → 动态轮询直到 status=completed（每 3s 查询一次，最多等 60s）→ 下载 OSS 导出文件 |
| 鉴权 | 接口需 JWT，从浏览器 localStorage 自动取 |
| **失败场景** | **⚠️ 如果文件上传失败** → 检查文件格式是否正确（A1=导入用户id） |
| | **⚠️ 如果轮询超时（60s）** → 检查平台中心是否正常响应 |
| | **⚠️ 如果导出文件为空** → 检查输入的学员ID是否有效 |
| | **⚠️ 如果接口返回 401** → 刷新 `liuyi_login/auth_state.json` 登录态 |

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

## 步骤 4：六一工作台 — 新建标签 ✅ 已实现

**目标**：在六一工作台创建本批次的两个标签（小账号版 + 大账号版）。

| 项 | 内容 |
|---|---|
| 入口 | https://home.61info.cn/ → 标签管理 |
| 登录 | 复用六一统一登录态 `liuyi_login/auth_state.json` |
| 实现 | Playwright 进六一工作台 → 浏览器内 `page.evaluate(fetch ...)` 调 `userTag/addwithfile`（multipart） |
| 鉴权 | 接口需 JWT，从浏览器 localStorage 自动取 |
| 输入 | `output/filtered_{yyyymmdd}.xlsx`（学员id列） + `output/dadou_mapping_{yyyymmdd}.xlsx`（豌豆大账号id列） |
| 命名 | `2026年{月}月海外益智停课学员` 和 `2026年{月}月海外益智停课学员（大账号）`，月份默认按当前月，可用 `--month YYYY-MM` 覆盖 |
| 业务参数 | `bizCode=WANDOU` / `tagType=531`（关键行为/活跃行为）/ `dataFrom=2`（按导入用户id筛选）/ `tagUpdateType=1` / `onceUpdate=1` |
| 文件格式 | xlsx，A 列表头「用户id」（与官方模板 `usertag_template.csv` 一致） |
| 中间产物 | `liuyi_tag/user_ids_{stamp}.xlsx`、`liuyi_tag/dadou_ids_{stamp}.xlsx` |
| 输出 | `output/tag_ids_{yyyymmdd}.json`（小账号 tagId / 大账号 tagId） |
| **失败场景** | **⚠️ 如果接口返回 401** → 刷新 `liuyi_login/auth_state.json` 登录态 |
| | **⚠️ 如果文件格式错误** → 检查 xlsx 第一列表头是否为「用户id」 |
| | **⚠️ 如果标签名已存在** → 在标签名后加 `_v2` 后缀重试 |

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-tag
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-tag --month 2026-07
```

---

## 步骤 5：六一工作台 — 新建用户群 ✅ 已实现

**目标**：基于步骤 4 的标签创建两个用户群（小账号版 + 大账号版）。复制 5 月模板（小账号 16812 / 大账号 16811）的 `tagsConfig` 结构，把第一项 `tagIds` 替换成新建的标签 id（19011/19012/...），其余「业务侧固定标签」（8026/2160/14272 交集；8025/8936 差集）原样保留。

| 项 | 内容 |
|---|---|
| 入口 | https://home.61info.cn/ → 标签管理 → 用户群 |
| 登录 | 复用六一统一登录态 `liuyi_login/auth_state.json` |
| 依赖 | `output/tag_ids_{yyyymmdd}.json`（由步骤 4 产出） |
| 模板 | 小账号 16812（单标签）/ 大账号 16811（含 5 个固定业务标签） |
| 接口 | `POST userGroup/checkData`（带模板 id 预检人群差异） + `POST userGroup/add`（不带 id 创建） |
| 输出 | `output/group_ids_{yyyymmdd}.json`（小账号 groupId / 大账号 groupId） |
| **失败场景** | **⚠️ 如果 tag_ids json 不存在** → 先运行步骤 4 创建标签 |
| | **⚠️ 如果接口返回 401** → 刷新 `liuyi_login/auth_state.json` 登录态 |
| | **⚠️ 如果用户群名已存在** → 在名称后加 `_v2` 后缀重试 |

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-group
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-group --month 2026-07
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-group --dry-run   # 只构造请求体打印不调 add
```

---

## 步骤 6：新建企微标签 ✅ 已实现

**目标**：把上一步创建的「大账号用户群」关联到企微「【益智】长期标签」标签组下，让企业微信侧能拉到这批人群打标签触达。注意：**只关联大账号**，小账号不需要。

| 项 | 内容 |
|---|---|
| 入口 | 六一工作台 → 标签管理 → 企微标签 |
| 接口 | `POST /corporate-wechat-backend/o/v1/tagGroup/create` |
| 依赖 | `output/group_ids_{yyyymmdd}.json`（取 dadou_group） |
| 业务参数 | `bizCode=WANDOU`，`corpTagGroupId=etN7IECgAAkW39vv9E__scZlJAnXZFzw`（「【益智】长期标签」固定 id） |
| 输出 | `output/wechat_tag_{yyyymmdd}.json`（企微 tag id） |
| **失败场景** | **⚠️ 如果 group_ids json 不存在** → 先运行步骤 5 创建用户群 |
| | **⚠️ 如果接口返回 401** → 刷新 `liuyi_login/auth_state.json` 登录态 |
| | **⚠️ 如果 corpTagGroupId 失效** → 联系企微管理员确认「【益智】长期标签」id 是否变更 |

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py wechat-tag
python ~/Desktop/停课唤醒目标/tingke_wakeup.py wechat-tag --dry-run
```

> 命名要点：步骤 4、5 的标签和用户群都按 **Windows 当月** 命名（`create_tag.py` / `create_group.py` 已强制锁死，`--month` 参数被忽略）。

---

## 步骤 7：北极星外呼平台 — 克隆「停课120天以内」任务到目标月份 ✅ 已实现

**目标**：在北极星把既有的「【海外】停课120天以内-N月」外呼任务模板克隆出一个新月份的版本，让客服坐席继续跟进。

🔴 **CHECKPOINT 2：克隆前确认**
- 目标月份是否正确？（默认下月）
- 源任务名称是否包含「【海外】停课120天以内」？

| 项 | 内容 |
|---|---|
| 入口 | https://passport.vipthink.cn/#/account/login → https://sh-center.vipthink.cn/#/ |
| 登录 | 钉钉扫码（独立登录态 `polaris_login/auth_state.json`，不复用六一登录态） |
| 接口 | `POST /task/taskTemplate/16/list`（搜任务）+ `GET /task/taskTemplate/getDetail?id={id}`（拉详情）+ `POST /task/taskTemplate/add`（克隆新建） |
| 鉴权 | `Authorization: Bearer eyJ...`，从浏览器 localStorage / sessionStorage 自动找（兼容 `Bearer xxx` 或纯 `eyJxxx` 两种存储） |
| **失败场景** | **⚠️ 如果钉钉扫码登录失效** → 重跑扫码登录脚本刷新 `polaris_login/auth_state.json` |
| | **⚠️ 如果搜不到源任务** → 检查任务名称前缀「【海外】停课120天以内」是否匹配 |
| | **⚠️ 如果新任务名已存在** → 检查 target-month 是否重复，已建则跳过 |
| | **⚠️ 如果接口返回 401** → 刷新北极星登录态 |

> 注：用户口径是「修改」，但接口实际是 **克隆 + 新建**（POST add），原任务保留不删。前缀「【海外】停课120天以内」固定，**只改末尾月份数字**，其他字段全部从原任务详情拷贝，**不做任何修改**。

**默认目标月份 = 下月**（运营节奏：月底为下月准备）。可用 `--target-month 1-12` 覆盖。

```bash
python ~/Desktop/停课唤醒目标/tingke_wakeup.py polaris-task           # 默认下月
python ~/Desktop/停课唤醒目标/tingke_wakeup.py polaris-task --target-month 7
python ~/Desktop/停课唤醒目标/tingke_wakeup.py polaris-task --dry-run  # 只构造 payload 不调 add
```

输出：`output/polaris_task_{yyyymmdd}.json`（包含 `source_task_id` / `new_task_id` / 新任务名）

---

## 步骤 8：六一工作台 — 标签数据同步配置 ✅ 已实现

**目标**：配置标签数据同步，将新建的用户群数据定期同步到豌豆数合表。

🔴 **CHECKPOINT 3：同步前确认**
- 选择的用户群名称是否**不带**"大账号"后缀？
- 同步频率是否设置为"每天"？

| 项 | 内容 |
|---|---|
| 入口 | 六一工作台 → 标签管理 → 标签数据同步 |
| 登录 | 复用六一统一登录态 `liuyi_login/auth_state.json` |
| 配置项 | 业务类型=豌豆，同步业务系统=豌豆数合表，同步用户群=新建的用户群（小账号版本），同步频率=每天，状态=启用 |
| 关键点 | **必须选择小账号版本用户群**（不带"大账号"后缀），例如 `2026年6月海外益智停课学员新` |
| 操作流程 | 1. 访问用户标签页面刷新 → 2. 访问用户群页面刷新 → 3. 回到标签数据同步页面 → 4. 点击新增 → 5. 填写表单 → 6. 点击确认 → 7. 点击手动同步 |
| 输出 | `sync_tag_final_{yyyymmdd}.png`（配置完成截图） |
| **失败场景** | **⚠️ 如果用户群下拉框找不到新建的群** → 重新刷新用户标签和用户群页面 |
| | **⚠️ 如果误选了大账号版本** → 删除已创建的同步配置，重新选择小账号版本 |
| | **⚠️ 如果接口返回 401** → 刷新 `liuyi_login/auth_state.json` 登录态 |
| | **⚠️ 如果手动同步按钮点击无响应** → 刷新页面后重试 |

**用法**：

```bash
python ~/Desktop/停课唤醒目标/liuyi_tag/sync_tag_data_final.py --month 2026-06
```

> 注意：标签数据同步配置需要在创建标签和用户群之后进行。脚本会自动访问标签和用户群页面刷新，确保新建的资源在下拉框中可见。

---

## 命令行总入口

```bash
# 端到端一把梭
python ~/Desktop/停课唤醒目标/tingke_wakeup.py all

# 只跑某一步（断点续跑）
python ~/Desktop/停课唤醒目标/tingke_wakeup.py export
python ~/Desktop/停课唤醒目标/tingke_wakeup.py filter
python ~/Desktop/停课唤醒目标/tingke_wakeup.py query-dadou-id
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-tag
python ~/Desktop/停课唤醒目标/tingke_wakeup.py liuyi-group
python ~/Desktop/停课唤醒目标/tingke_wakeup.py wechat-tag
python ~/Desktop/停课唤醒目标/tingke_wakeup.py polaris-task
python ~/Desktop/停课唤醒目标/tingke_wakeup.py sync-tag-data

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

## 反模式（❌ 不要这样做）

以下操作会导致流程失败、数据错误或下游触达异常，严禁执行：

1. ❌ **不要在 BI 导出时修改「开始日期」「结束日期」**
   - 签到时间筛选在步骤 2 中通过 Excel 完成
   - 修改 BI 筛选条件会导致数据范围错误

2. ❌ **不要手动修改 `output/` 目录下的中间产物文件**
   - 中间产物用于断点续跑，手动修改会破坏数据一致性
   - 如需调整，重新运行对应步骤生成新文件

3. ❌ **不要在企微标签步骤关联小账号用户群**
   - 企微标签只关联大账号用户群（步骤 5 产出）
   - 关联小账号会导致企业微信侧无法正确识别用户

4. ❌ **不要同时挂载小账号和大账号到同一个企微标签**
   - 企微侧用户实体是大账号粒度；同时挂载会产生重复打标签和触达
   - 步骤 6 的 `tagGroup/create` 必须只传 dadou_group，禁止把 wandou_group 一并塞进去
   - 如果历史已经误挂，先在企微后台解除小账号关联，再重跑步骤 6

5. ❌ **不要在北极星任务克隆时修改任务详情字段**
   - 只改末尾月份数字，其他字段全部从原任务拷贝
   - 修改其他字段可能导致外呼流程异常

6. ❌ **不要在标签数据同步时选择大账号用户群**
   - 必须选择小账号版本用户群（不带"大账号"后缀）
   - 选错会导致豌豆数合表同步失败

7. ❌ **不要在流程中途切换月份参数**
   - 同一次执行的所有步骤必须使用相同的 `--month` 参数
   - 切换月份会导致中间产物不匹配

8. ❌ **不要跳过首次使用的 `--dry-run` 演练**
   - 首次执行务必先 dry-run 确认参数
   - 直接正式跑可能因配置错误产生脏数据

9. ❌ **不要把学员ID当作豌豆大账号ID直接传入北极星/企微**
   - BI 导出的「学员id」= 豌豆账号 id，与豌豆大账号 id 是多对一关系
   - 必须经过步骤 3 的映射，下游系统识别的是大账号
   - 跳过映射会出现「找得到人但触达不到」的静默失败

10. ❌ **不要在登录态失效时重试同一步骤刷状态**
    - 6 个步骤里有 5 个依赖 `liuyi_login/auth_state.json`，1 个依赖 `polaris_login/auth_state.json`
    - 401 出现时先重跑对应登录脚本，再续跑业务步骤；盲目重试只会消耗接口配额

## 待办（未完成项已标注为 🚨 阻塞风险）

> 状态说明：以下未完成项不影响当前 8 步主流程在已实现路径上的正常运行（status=stable 即针对当前实现），但会在「换人接手」「换月调整规则」「平台改版」时成为阻塞点，必须在触发场景出现前补齐。

- [x] BI 报表名 + 业务筛选条件（思维停课学员执行明细，已确认）
- [ ] 🚨 **阻塞风险**：数据筛选规则 + 单次人数上限
  - 影响：步骤 2 当前硬编码「距今前 1~4 个月」窗口；如运营调整窗口或加上限，没有文档化规则就只能改代码
  - 触发场景：换运营负责人、调整外呼节奏
- [ ] 🚨 **阻塞风险**：六一工作台 URL / 登录方式 / 新建标签&用户群的菜单路径
  - 影响：登录态失效或平台改版时，没有手工兜底路径，无法复现自动化操作
  - 触发场景：六一工作台改版、`liuyi_login/auth_state.json` 长期失效
- [ ] 🚨 **阻塞风险**：企微标签的入口和命名规范
  - 影响：`corpTagGroupId=etN7IECgAAkW39vv9E__scZlJAnXZFzw` 是硬编码 id，没有人工查找路径
  - 触发场景：企微管理员变更「【益智】长期标签」结构
- [ ] 🚨 **阻塞风险**：北极星外呼平台 URL / 登录方式 / 任务模板
  - 影响：钉钉扫码登录态失效时只能依赖脚本重跑，没有手工确认任务模板的路径
  - 触发场景：北极星改版、模板任务名前缀变更
- [ ] 🚨 **阻塞风险**：是否有现成的 OpenAPI 可绕过浏览器
  - 影响：当前全链路依赖 Playwright + 浏览器登录态，稳定性受平台前端影响
  - 触发场景：高并发批量场景、需要在无头服务器上跑
