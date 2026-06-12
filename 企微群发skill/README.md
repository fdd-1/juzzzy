# 企微群发 SKILL — 操作手册

这份手册写给「**接手这个 skill 的人**」。看完应该能从 0 到 1 跑完一次「读 docx → 建话术模板 → 建群发任务」。

---

## 一、它是干嘛的

把下面这个原本要在六一工作台手点 30+ 次的流程自动化：

```
本地 docx + 附件
    ↓ 阶段1 解析
按模块拆出来的话术 + 附件清单（modules.json）
    ↓ 阶段2 建模板（自动开浏览器点）
六一-企微话术模板管理-豌豆素质 Tab 下批量出 N 个话术模板
    ↓ 阶段3 建任务（自动开浏览器点）
六一-企微群发任务配置-豌豆素质 Tab 下批量出 N 个群发任务
```

每个阶段一个 Python 脚本：[parse_docx.py](parse_docx.py)、[create_template.py](create_template.py)、[create_task.py](create_task.py)。

---

## 二、第一次使用：环境准备

需要装的东西，**一次装好以后不用再装**：

### 1. Python 3.10+

打开 PowerShell 输入 `python --version` 看一下。没装就到 https://www.python.org/downloads/ 下一个，安装时勾上「Add to PATH」。

### 2. 装依赖（在 skill 文件夹下打开 PowerShell）

```powershell
pip install playwright python-docx
python -m playwright install chromium
```

第三行会下载一个 200MB 左右的 Chromium 浏览器内核，用作 UI 自动化。装一次以后不用动。

### 3. 第一次跑会要扫码登录六一

任何阶段第一次跑，会自动开一个 Chrome 窗口跳到六一登录页。**用钉钉扫码登录一次**就行，登录态会存到当前文件夹下的 `browser_profile/` 里，以后不用再扫。

🔴 **注意**：`browser_profile/` 是个文件夹，跨电脑迁移这个 skill 时要把这个文件夹一起拷过去（或者拷过去后第一次跑时重新扫一次码也行）。

---

## 三、每次使用要给 AI 什么

跑这个 skill 之前，**必须**准备好下面 5 样东西，缺一样就跑不了：

### 必给 5 项

| # | 名字 | 例子 | 说明 |
|---|------|----|------|
| 1 | **docx 文件路径** | `D:\发送话术\0611素材.docx` | Word 文档，按模块写话术 |
| 2 | **附件目录路径** | 一般和 docx 同目录就行 | PDF / 图片 / 视频物料放这里 |
| 3 | **用户群名（精确）** | `【S8】香港益智在读粤语班学员` | 在六一用户群列表里能搜到的精确名字 |
| 4 | **执行时间** | `2026-06-13 14:00` | 任务的定时执行时间，YYYY-MM-DD HH:MM |
| 5 | **执行团队**（可选） | `港澳1组,港澳2组,港澳组` | 不传时按用户群名自动推（见下表） |

### 执行团队自动推断规则

如果用户群名里包含下面的关键字，可以不传 `--teams`，脚本自动推：

| 用户群名包含 | 默认执行团队 |
|---|---|
| 港澳 / 香港 / 澳门 | 港澳1组、港澳2组、港澳组 |
| 亚欧 / 美澳 / 欧美 | 美澳1组、美澳2组、美澳3组、美澳4组、美澳5组 |
| 其它（不含关键字） | **必须**显式 `--teams "组1,组2"` 传 |

---

## 四、docx 怎么写

文档结构必须满足：

1. **每个模块用 Word「标题样式」隔开**，或者每个模块开头一行写「`数字. 文字`」（如「`1. S5级别`」）。
2. 标题下放话术正文（一段或多段都行）。
3. 一个 docx 里可以多个模块，每个模块独立成一个话术模板。

### 例子（4 个模块的 docx）

```
1. S5级别
寶貝家長您好！「校內模擬評估」活動已經圓滿結束啦～...

2. S6级别
寶貝家長您好！「校內模擬評估」活動已經圓滿結束啦～...

3. S7级别
寶貝家長您好！...

4. S8级别
寶貝家長您好！...
```

### 附件（PDF / 图片 / 视频）

- 跟 docx 放同一个文件夹（或单独一个目录）
- 文件名里**带模块关键字**（脚本靠文件名自动归类）

例：

```
0611素材/
├── 0611素材.docx
├── S5模擬卷B卷試題&答案.pdf   ← 自动归到 S5级别 模块
├── S6模擬卷B卷試題&答案.pdf   ← 自动归到 S6级别 模块
├── S7模擬卷B卷試題&答案.pdf
└── S8模擬卷B卷試題&答案.pdf
```

---

## 五、跑起来：完整命令

### 阶段 1：解析 docx

```powershell
cd <skill 所在路径>
python parse_docx.py --input "D:\发送话术\0611素材\0611素材.docx" --attachments-dir "D:\发送话术\0611素材"
```

跑完会在 `exports\YYYYMMDD\modules.json` 看到解析结果，控制台会打印类似：

```
[OK] 已写入 exports\20260611\modules.json
[NEXT] 看一眼 modules.json 没问题就跑 create_template.py
```

打开 `modules.json` 看一眼模块名 / 附件路径对不对。

### 阶段 2：建话术模板

```powershell
python create_template.py --modules-json exports\20260611\modules.json
```

第一次跑会开浏览器要求扫码（钉钉扫六一登录二维码），扫完会自动进话术模板管理页，按 modules.json 里的模块逐个建模板。

跑完看 `exports\20260611\templates_log_<时间戳>.csv`：

```
index,module,template_name,status,error
1,S5级别,S5级别0611,OK,
2,S6级别,S6级别0611,OK,
3,S7级别,S7级别0611,OK,
4,S8级别,S8级别0611,OK,
```

status 全 OK 就可以进阶段3。有 FAIL 的看 `screenshots/` 里的截图分析。

#### 阶段 2 常用开关

| 开关 | 作用 |
|---|---|
| `--only 1,3` | 只跑第 1 和第 3 个模块（从 1 开始计数） |
| `--date 0611` | 强制模板名后缀为 0611（默认是今天 MMDD） |
| `--keep-open` | 跑完不关浏览器，方便检查 |
| `--dry-run` | 只打印计划不跑（看用户群匹配 / 附件路径对不对） |

### 阶段 3：建群发任务（一个模板对应一个任务）

```powershell
python create_task.py --template-name "S8级别0611" --user-group "【S8】香港益智在读粤语班学员" --exec-at "2026-06-13 14:00" --submit
```

参数解释：

- `--template-name`：阶段2 建好的话术模板名（要和那个模板**一模一样**）
- `--user-group`：用户群名，必须精确（含【】等符号）
- `--exec-at`：定时执行时间，格式 `YYYY-MM-DD HH:MM`
- `--submit`：跑到底真创建任务（不加这个开关只到表单填好就停下来，方便人工确认）

#### 阶段 3 常用开关

| 开关 | 作用 |
|---|---|
| `--submit` | 真创建任务（默认不真创建） |
| `--click-preview` | 跑完点「预览」按钮，停在预览页（不真创建），用来人工确认 |
| `--teams "组1,组2"` | 显式指定执行团队，绕过自动推断 |
| `--keep-open` | 跑完不关浏览器 |
| `--start-at "YYYY-MM-DD HH:MM"` | 显式指定任务时效开始时间（一般不用） |
| `--end-days 6` | 任务时效结束 = 现在+N天（默认 6） |

#### 多个用户群批量起任务

每个用户群跑一次脚本就行。例如 4 个级别要发 4 个任务：

```powershell
python create_task.py --template-name "S5级别0611" --user-group "【S5】香港益智在读粤语班学员" --exec-at "2026-06-13 14:00" --submit
python create_task.py --template-name "S6级别0611" --user-group "【S6】香港益智在读粤语班学员" --exec-at "2026-06-13 14:00" --submit
python create_task.py --template-name "S7级别0611" --user-group "【S7】香港益智在读粤语班学员" --exec-at "2026-06-13 14:00" --submit
python create_task.py --template-name "S8级别0611" --user-group "【S8】香港益智在读粤语班学员" --exec-at "2026-06-13 14:00" --submit
```

---

## 六、跟 AI 说话的标准开场（怎么让 AI 帮我跑）

把下面的模板复制改填后丢给 AI（Claude / Cursor / Codex 都行）：

```
帮我跑企微群发 skill，建话术模板 + 建群发任务。

docx 路径：D:\发送话术\0611素材\0611素材.docx
附件目录：同 docx 目录
要建任务的用户群（每行一个，对应一个模板）：
- S5级别 → 【S5】香港益智在读粤语班学员
- S6级别 → 【S6】香港益智在读粤语班学员
- S7级别 → 【S7】香港益智在读粤语班学员
- S8级别 → 【S8】香港益智在读粤语班学员
执行时间：2026-06-13 14:00
执行团队：默认（按用户群名自动推港澳1/2/组）

按 SKILL.md 流程跑，每阶段给我看结果再进下一步。
```

AI 会按 SKILL.md 的 CHECKPOINT 流程：跑完阶段1 给你看 modules.json → 跑完阶段2 给你看 templates_log.csv → 跑完阶段3 给你看任务列表截图。

---

## 七、常见错误与处理

| 报错 / 现象 | 原因 | 处理 |
|---|---|---|
| `ModuleNotFoundError: No module named 'playwright'` | 没装依赖 | `pip install playwright python-docx` |
| `Executable doesn't exist at ...chromium-...` | 没装 Chromium 内核 | `python -m playwright install chromium` |
| 浏览器一直停在登录页 | 第一次用，要扫码 | 用钉钉扫码登录六一，扫完脚本会继续 |
| `任务时效只能选择10分钟之后到10080分钟之内` 红字 | 时间 buffer 不够 | 重跑（脚本有 5min 起始 buffer，正常情况下不会踩） |
| 阶段3 提示「找不到用户群」 | 用户群名打错 / 用户群已下线 | 在六一 UI 搜一下用户群名，复制后重跑 |
| 阶段3 任务建成功但没勾上执行团队 | 用户群名不含港澳/亚欧关键字 | 加 `--teams "港澳1组,港澳2组,港澳组"` 显式传 |
| 模板创建后预览看到文字是空的 | 文字编辑器失焦了（脚本应该已经修过这个，再翻车就报） | 看 `screenshots/err_*.png`，截图给我 |

---

## 八、文件结构速查

```
企微群发skill/
├── SKILL.md                # AI 看的执行指引
├── README.md               # 这份手册
├── parse_docx.py           # 阶段1
├── create_template.py      # 阶段2
├── create_task.py          # 阶段3
├── browser_profile/        # Playwright Chrome 持久化数据（含登录态）
├── exports/                # 每次跑的产物
│   └── 20260611/
│       ├── modules.json                # 阶段1 输出
│       ├── templates_log_*.csv         # 阶段2 输出
│       └── screenshots/
│           ├── task_landing_*.png      # 阶段3 落地页
│           ├── task_filled_*.png       # 阶段3 表单填好
│           ├── task_preview_*.png      # 阶段3 预览页
│           └── task_submitted_*.png    # 阶段3 提交后
└── 话术模板资料/             # 历史素材（参考用，不影响脚本）
```

---

## 九、能做什么 / 不能做什么

### ✅ 能做

- 解析按模块拆分的 docx + 附件
- 在六一-企微话术模板管理-**豌豆素质** Tab 下批量建话术模板
- 在六一-企微群发任务配置-**豌豆素质** Tab 下批量建群发任务
- 任务发送方式 = 手动群发，科目 = 豌豆益智，类型 = 续费
- 自动按用户群名推执行团队（港澳/美澳）
- 任务时效自动满足六一约束（10min ≤ start, end ≤ 10080min）

### ❌ 不能做（往里扩展前先想清楚）

- 别的 Tab（咕比 / 画啦啦 / 派培优 / 个微 / 小灯塔 / 派斯）
- 自动群发（只跑手动群发）
- 删除已建的模板/任务（去六一 UI 手动删）
- 修改已建的模板（六一不支持改，只能新建）
- 钉钉文档 API 拉取（已弃用，只支持本地 docx）
- 自动群发的「自动执行」（只跑「定时执行」）

---

## 十、找谁问

- **脚本报错 / 跑不通**：把 `exports/<日期>/screenshots/` 里相关时间戳的截图 + 控制台日志贴给我
- **新增 Tab 支持 / 流程改动**：先在六一 UI 手动跑一遍，截全流程图，再让我改脚本
- **登录态过期**：删 `browser_profile/` 重新扫码登录
