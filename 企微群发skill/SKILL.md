---
name: qiwei-broadcast-builder
description: 自动化六一工作台「企微话术模板管理 + 企微群发任务配置」。用户给一份本地 docx（话术）和同目录附件（PDF/图片/视频），skill 解析后批量建话术模板，再按用户群和执行时间批量建群发任务。触发词：「建话术模板」「批量建群发任务」「跑企微群发」「按这份文档建模板」。
---

# 企微群发 SKILL

## TL;DR

```
[阶段1] python parse_docx.py     → modules.json
[阶段2] python create_template.py → 在六一建话术模板
[阶段3] python create_task.py    → 在六一建群发任务
```

每阶段独立可重跑。失败有截图、日志和可回退状态。

---

## 触发场景

✅ 该用：
- 用户给了一份 docx（按模块分话术）+ 同目录附件，要批量建模板
- 要在「企微群发任务配置-豌豆素质 Tab」批量起群发任务
- 任务发送方式 = 手动群发，科目 = 豌豆益智，任务类型 = 续费

❌ 不该用：
- 其它 Tab（咕比团队 / 画啦啦团队 / 派培优 / 个微 / 小灯塔 / 派斯）— 本 skill 只覆盖豌豆素质
- 自动群发（本 skill 只跑手动群发）
- 删除已建模板/任务（六一 UI 手动删）
- 钉钉文档 API 拉取（已弃用，只支持本地文件）

---

## 执行前必须问用户拿的信息（CHECKPOINT）

🔴 **CHECKPOINT 1**：执行前必须确认下面 5 项，缺一项就停下问。

| # | 信息 | 例 | 用途 |
|---|------|----|------|
| 1 | docx 文件绝对路径 | `D:\发送话术\0611素材.docx` | 阶段1 解析 |
| 2 | 附件目录（PDF/图片）绝对路径 | `D:\发送话术\附件\` 或同 docx 目录 | 阶段1 匹配附件 |
| 3 | 用户群名（精确匹配） | `【S8】香港益智在读粤语班学员` | 阶段3 选用户群 |
| 4 | 执行时间（YYYY-MM-DD HH:MM） | `2026-06-13 14:00` | 阶段3 定时执行 |
| 5 | 执行团队（可选，不传按规则推断） | `港澳1组,港澳2组,港澳组` | 阶段3 |

**STOP**：以上任意一项空缺，要么向用户追问，要么提示用户「我需要 X 才能跑」。不要靠默认值/猜测填写。

---

## 阶段 1：解析 docx → modules.json

```bash
python parse_docx.py --input <docx路径> --attachments-dir <附件目录>
```

### 文档约定（用户写 docx 时遵守）

- 用 Word 标题样式或「数字. 文字」开头（如 `1. S5级别` `2. S6级别`）切分模块
- 每个模块标题下放话术正文，或用两列表格（第一列 = 话术，第二列 = 物料标签）
- 模块名取标题去掉「`1.` / `2.`」前缀，例：`S5级别`、`S8级别`
- 附件文件名包含模块关键字（如 `S5模拟卷.pdf` 会自动归到 `S5级别` 模块）

### 输出

```
exports/{YYYYMMDD}/modules.json
```

结构：
```json
[
  {
    "name": "S8级别",
    "texts": ["寶貝家長您好！..."],
    "attachments": [
      {"path": "<绝对路径>", "type": "file|image|video", "filename": "S8模擬卷.pdf"}
    ]
  }
]
```

### 失败模式（if-then）

| 触发 | 一线修复 | 仍失败兜底 |
|---|---|---|
| 解析出 0 个模块 | 检查 docx 是否用了「标题样式」或「数字. 文字」起头 | 让用户改 docx 标题，重跑 |
| 附件 0 个匹配 | 检查附件文件名是否含模块关键字（S5/S6 等） | 让用户重命名附件后重跑 |
| `python-docx` 报错 | `pip install python-docx` | 提示用户装依赖 |

🔴 **CHECKPOINT 2**：解析完打印计划摘要（每模块 N 段文字 + M 个附件），让用户瞄一眼再进阶段2。

---

## 阶段 2：建话术模板

```bash
python create_template.py --modules-json exports/<日期>/modules.json [--only 1,3] [--keep-open]
```

### 表单字段（已确认，硬编码在脚本里）

| 字段 | 值 |
|---|---|
| 团队 Tab | 豌豆素质 |
| 模板名 | `{模块名}{MMDD}`（如 `S8级别0611`，超 20 字截断） |
| 选择科目 | 豌豆益智 |
| 可调用功能 | 企微群发任务配置 |
| 话术类型 | 公共话术 |
| 添加话术 | 文档里有什么形式就加什么；文字 + 文件，每种 ≤30 |

### 失败模式（if-then）

| 触发 | 一线修复 | 仍失败兜底 |
|---|---|---|
| 浏览器跳到登录页 | 扫码登录六一（首次必扫） | profile 损坏 → 删 `browser_profile/` 重扫码 |
| 文字子弹窗保存后 editor 是空的 | 脚本已自动重试 3 次 | 仍失败：截图 `screenshots/err_*.png`，人工补 |
| 弹窗 z-index 遮挡按钮点不到 | 脚本用 `evaluate` 直接点击绕开 | 仍失败：减小并发，去掉 `--only`，单个跑 |
| 模板名重复 | 六一会拒绝 | 改 `--date` 或先去 UI 删旧的 |
| 附件路径不存在 | 脚本预检会 `sys.exit(2)` | 检查 `modules.json` 里附件路径，重跑阶段1 |

### 输出

```
exports/{YYYYMMDD}/templates_log_{stamp}.csv  # index, module, template_name, status, error
exports/{YYYYMMDD}/screenshots/               # 失败步骤截图
```

🔴 **CHECKPOINT 3**：阶段2 跑完打开 csv 看 status 是否全 OK；有 FAIL 看截图诊断后再决定是否进阶段3。

---

## 阶段 3：建群发任务

```bash
python create_task.py \
    --template-name <模板名> \
    --user-group "<用户群名精确>" \
    --exec-at "YYYY-MM-DD HH:MM" \
    [--teams "组1,组2"] \
    [--submit] [--keep-open]
```

### 表单字段

| 字段 | 值 / 来源 |
|---|---|
| 团队 Tab | 豌豆素质 |
| 发送方式 | 手动群发（先选这个，点确定才会展开剩余字段） |
| 选择科目 | 豌豆益智 |
| 任务类型 | 续费 |
| 任务名称 | 与话术模板同名（`--template-name`） |
| 优先使用 | 高级群发 |
| 选用户群 | `--user-group`，每次跑由用户给 |
| 选择话术 | 同任务名搜出来 |
| 沟通中和未回复 | 过滤 |
| 任务时效 | 脚本实时计算：start = 现在+15min，end = 现在+10000min |
| 执行团队 | `--teams` 显式给，或按用户群名自动推（见下） |
| 执行方式 | 定时执行 |
| 执行时间 | `--exec-at`，必须落在任务时效区间内 |

### 任务时效六一硬约束

- start ≥ 现在 + 10 分钟
- end ≤ 现在 + 10080 分钟（= 7 天 = 168 小时）

脚本在「点开新建任务后、填日期前」实时重算，给 5 min 起始 buffer + 80 min 结束 buffer。**不要在脚本启动时就算时间**——从启动到填日期可能过去几分钟，会踩 10 分钟下限。

### 执行团队自动推断规则

| 用户群名包含 | 默认执行团队 |
|---|---|
| 港澳 / 香港 / 澳门 | 港澳1组、港澳2组、港澳组 |
| 亚欧 / 美澳 / 欧美 | 美澳1组、美澳2组、美澳3组、美澳4组、美澳5组 |
| 其它 | 留空，必须显式 `--teams` 传 |

cascader 是层级（公司 → 业务线 → 中心 → 服务部 → 服务区 → 组），脚本用 cascader filterable 搜索能力直接搜组名，不用一层层点。

### 提交流程

| 开关 | 行为 |
|---|---|
| 默认（无开关） | 填完表停在表单页，方便人工核对，不点预览 |
| `--click-preview` | 填完点「预览」按钮，停在预览页，**不**真创建 |
| `--submit` | 填完点「预览」→ 预览页点「**确认创建**」（4 字按钮，不是 2 字「确认」），真落库 |

🔴 **CHECKPOINT 4**：第一次跑某个用户群时，强烈建议先 `--click-preview` 看一遍预览页，确认无误再加 `--submit` 重跑。

### 失败模式（if-then）

| 触发 | 一线修复 | 仍失败兜底 |
|---|---|---|
| 「任务时效只能选择10分钟之后到10080分钟之内」红字 | 时间被脚本跑慢了 → 重跑 | 脚本里 buffer 加大（改 fill_form 里的 15min/10000min） |
| 用户群搜不到 | 名字打错或用户群已下线 | 让用户在六一 UI 确认用户群名 |
| 执行团队没勾上 | 用户群名不含港澳/亚欧关键字 | 加 `--teams "组1,组2"` 显式传 |
| 找不到「确认创建」按钮 | 按钮叫法变了 | dump `预览页可见按钮` 看实际文字，更新 create_task.py 的正则 |
| 浏览器中途关闭 | TargetClosedError | 重跑（profile 还在，不用重新扫码） |

### 输出

```
exports/{YYYYMMDD}/screenshots/
├── task_landing_{stamp}.png       # 进群发任务配置页
├── task_filled_{stamp}.png        # 表单填好（预览前）
├── task_preview_{stamp}.png       # 预览页
└── task_submitted_{stamp}.png     # 提交后任务列表
```

---

## 反例黑名单（什么不要做）

| # | 反模式 | 为什么不要做 | 替代做法 |
|---|---|---|---|
| 1 | 在脚本启动时就算 `任务时效.start = now+15min` | 跑到填日期可能过去 5+ 分钟，踩 10 分钟下限被驳回 | 在 `fill_form` 内部 `任务时效` 这一步实时重算 |
| 2 | `start = now+22min, end = start+6days` | end 距 now 超过 7 天上限被驳回 | end 必须以 `now` 为基准，不是 `start` |
| 3 | 把 `预览` 按钮当成 footer 的「确定」按钮点 | 文字不一样：`确定` vs `预览` vs `确认创建`，正则要分清 | 各按钮用独立正则，不要混 |
| 4 | 文字子弹窗保存只看 dialog 是否关掉 | 关掉不等于内容保存了（S7/S8 之前翻车） | 保存前用 `evaluate` 读 editor.innerText 验证非空 |
| 5 | 用 `Escape` 关 date-range-picker popper | 焦点已不在 input，Escape 不生效 | 用 `evaluate` 强制 `display:none + pointer-events:none` |
| 6 | `cascader` 用「点 → 展开父节点 → 勾叶子」 | 层级深（6 级）容易踩空 | 用 cascader filterable 直接搜组名 |
| 7 | 删 `browser_profile/` 试图修登录问题 | 会丢扫码登录态，下次必须重新扫 | 先重跑确认是不是其它原因 |
| 8 | 跑完不看 `templates_log_*.csv` 就进阶段3 | 阶段2 失败的模板进阶段3 一定搜不到话术 | 每阶段都看日志/截图确认 |

---

## 跨电脑使用约束（runtime neutrality）

- ✅ 所有 Python 脚本只用相对路径（`Path(__file__).resolve().parent`），不写 `C:\Users\xxx`
- ✅ 命令里不固化用户名 / 工作目录
- ✅ 第三方依赖：`playwright`、`python-docx`，跨平台都能装
- ✅ Windows / macOS / Linux 都能跑（六一是 Web 的）
- ⚠️ `browser_profile/` 是 Chrome 数据目录，跨电脑迁移要重新扫码登录

---

## 资源清单

| 文件 | 用途 |
|---|---|
| `parse_docx.py` | 阶段1：本地 docx → modules.json |
| `create_template.py` | 阶段2：UI 自动化建话术模板 |
| `create_task.py` | 阶段3：UI 自动化建群发任务 |
| `browser_profile/` | Playwright 持久化 Chrome profile（首次扫码后保留） |
| `exports/{YYYYMMDD}/` | 每天的 modules.json + 日志 + 截图 |
| `README.md` | 给交接同事看的操作手册 |
| `SKILL.md` | 本文件，给 AI agent 看的执行指引 |
