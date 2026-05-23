---
name: crm-lesson-package
description: 在豌豆思维 CRM「财务 → 课时包管理」批量创建课时包。基于 Playwright 自动化 Element UI 表单，输入是 Excel SKU 配置表，输出是逐条 CSV 日志 + 失败截图。
---

# CRM 课时包批量创建 Skill

## 1. 何时使用本 Skill

触发条件（任一即可）：

- 用户给出 SKU / 套餐 / 课时包 配置表（Excel / 飞书多维表格），要求把这些套餐配置到豌豆思维 CRM 里。
- 用户提到「批量创建课时包」「批量配置套餐」「按 Excel 跑课时包」「自动填课时包表单」。
- 用户给出 `https://crm.vipthink.cn/.../ClassPackageManage` 链接并要求按列表添加。

不要用本 skill 处理：

- 编辑或删除已有课时包（脚本只覆盖「添加」流程，删除/修改是高风险操作，没写）。
- 课时包以外的 CRM 表单（学员、订单、退费等）。

## 2. 项目位置（不要复制到别处）

实现工程位于：`c:\Users\fengjianyi\Desktop\crm-lesson-package-skill\`

```
crm-lesson-package-skill/
├── crm_create_lesson_package.py        # 单条：调试 / 创建一条
├── crm_batch_create_lesson_packages.py # 批量：从 Excel 跑
├── config.example.json
├── config.local.json                   # gitignored，含真实 URL/字段
├── auth_state.json                     # gitignored，登录态
├── utils/
│   ├── auth.py        # ensure_login / first_time_login
│   └── element_ui.py  # fill_text / fill_filterable_dropdown / fill_cascade_dropdown / fill_multi_select / click_button_with_text
├── debug_after_type.py     # 探针：填完课包类型后看动态字段
├── debug_save_button.py    # 探针：弹窗按钮文案
├── debug_dropdown_options.py  # 探针：级联下拉的真实可选项
└── logs/                   # gitignored，CSV + 截图
```

接到批量任务时，**直接复用上面的脚本**，不要在别处重写一份。配置改 `config.local.json`，数据来源用 `--xlsx` 切换。

## 3. 标准流程

### 3.1 准备

1. 读 `config.local.json` 确认 `crm.home_url` / `crm.class_package_url` 仍指向目标环境。
2. 确认 `auth_state.json` 存在；不存在则先跑一次 `python crm_create_lesson_package.py --dry-run` 让用户扫码。
3. 拿到用户的 Excel 路径（默认是 `c:\Users\fengjianyi\Desktop\skill学习文档\港澳6月SKU课时包配置.xlsx`，每次任务可被覆盖）。

### 3.2 解析 Excel

Excel 不是「表头 + 行」结构，而是**每行是 (label, value, label, value, ...) 交替排列**的横向 KV。脚本里 `parse_excel` 已实现：

- 以 `课时包名称` 作为行起始标记，跳过非数据行（标题行、备注行 "年课包pro请填写成年课包" / "其他填写为其他类型课包"）。
- 字段映射见 `HEADER_MAP`；`课包类型` 是叶子（如「中课包」「年课包」），通过 `TYPE_PARENT` 补一级父「常规正课」拼成 `[parent, leaf]`，给级联下拉。
- `适用课类` 用顿号 / 中文逗号 / 英文逗号都能切。

**新增列时**改 `HEADER_MAP` 即可；**新增叶子节点**改 `TYPE_PARENT`。

### 3.3 跑批量

```powershell
cd c:\Users\fengjianyi\Desktop\crm-lesson-package-skill
$env:PYTHONIOENCODING = "utf-8"
python crm_batch_create_lesson_packages.py --xlsx "<excel 路径>"
```

可选参数：
- `--start N`：从第 N 条开始（1-based），断点续跑用。
- `--limit N`：最多跑 N 条，先小批验证用。
- `--skip-existing`：列表搜索框查重，已存在则跳过。重跑前必加，避免重复创建。

输出：
- `logs/batch-<timestamp>.csv`：序号 / 名称 / 结果 / 详情 四列，UTF-8 BOM，可直接 Excel 打开。
- 每个失败条都落 `logs/submit-fail-<safe_name>-<ts>.png` 截图。

### 3.4 关键已知坑（按踩过的顺序）

| 坑 | 现象 | 处理 |
|---|---|---|
| 成功 toast 文案是 `OK`，不是「保存成功」 | 全跑结果显示 16/16 FAIL，但 CRM 里实际已创建 | 成功关键词列表必须包含 `.el-message: OK` / `OK` / `Success` 变体（见 `submit_and_verify`） |
| `el-message` 自动消失 ~3s | 单次 `wait_for_selector` 抓不到 | 用 deadline 轮询循环（10s）+ `_collect_messages`，每 200ms 扫一次 |
| 课包类型选完后被课包分类反向清空 | 提交时报「课包类型必填」 | `fill_form` 里先填类型再填分类，填完读 `_read_select_value` 校验，被清空就重选 |
| 「打卡次数 / 停课次数」字段是**选完课包类型后才出现** | 一开始填表没这俩字段 | 顺序必须：名称 → 类型 → 分类 → 数字字段；后者要 `if "checkin_count" in data` 守 |
| 级联二级选项文案与口语不一致 | 「两年包」点不到，报 `第 2 级找不到选项: 两年包` | 用 `debug_dropdown_options.py` 探出真实文案（可能是「两年课包」「两年课包pro」之类），改 Excel 或在 `TYPE_PARENT` 里做别名 |
| Windows GBK stdout | 打印中文/emoji 报 `UnicodeEncodeError` | 入口加 `sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")`，或外部 `set PYTHONIOENCODING=utf-8` |
| 单次浏览器会话脏状态 | 上一条失败后弹窗残留 | 每条循环开头 `goto_list` 重新刷列表页；失败分支调用 `close_dialog`（取消 → Escape → 二次确认） |

详细可执行片段见 [references/known-pitfalls.md](references/known-pitfalls.md)。

## 4. 表单字段速查

| 中文 label | config 字段 | 控件类型 | 备注 |
|---|---|---|---|
| 课时包名称 | `name` | `el-input` | 全局唯一，重名会失败 |
| 课包类型 | `package_type` | 二级级联 `el-select-dropdown` | 必须传 `[parent, leaf]`，如 `["常规正课", "中课包"]` |
| 课包分类 | `package_category` | 可搜索 `el-select` | 例：`小班直播课` |
| 有效期 | `valid_months` | `el-input-number` | 单位：月 |
| 补课次数 | `makeup_chances` | `el-input-number` | |
| 普通课时 | `normal_lessons` | `el-input-number` | |
| 赠送课时 | `gift_lessons` | `el-input-number` | |
| 原价 | `original_price` | `el-input-number` | 单位：元 |
| 优惠价 | `discount_price` | `el-input-number` | |
| 试学期 | `trial_days` | `el-input-number` | 0 也要传 |
| 打卡次数 | `checkin_count` | `el-input-number` | 选完类型后出现 |
| 停课次数 | `suspend_count` | `el-input-number` | 选完类型后出现 |
| 适用课类 | `applicable_classes` | `el-select multiple` | 数组，例：`["豌豆思维2024", "豌豆明思2024"]` |
| 录播营 | `recorded_course` | 可搜索 `el-select` | 可空 |

字段名一旦从 CRM 那头加新的，先用 [`debug_after_type.py`](../../Desktop/crm-lesson-package-skill/debug_after_type.py) 探一遍真实 label，再改 `HEADER_MAP` 和 `fill_form`。

## 5. 安全 / 边界

- 默认 `headless=False`，肉眼监控；自动化跑长批量也保持有头模式。
- 不实现「删除课时包」流程，避免误操作。
- `auth_state.json` / `config.local.json` / `logs/` 全部 gitignore。
- 重跑前永远先评估「上一次有多少条已实际创建」，必须用 `--skip-existing` 防重复。

## 6. 探针脚本（调试时用）

| 脚本 | 用途 |
|---|---|
| `debug_after_type.py` | 选完课包类型后，打印弹窗里所有 `el-form-item__label`，发现新字段 |
| `debug_save_button.py` | 打印弹窗 footer 按钮文案，确认「确定」「保存」具体叫什么 |
| `debug_dropdown_options.py` | 打开任意级联，打印当前可见的真实选项文本，对齐 Excel 的口径 |

写新探针时复用这三个的模板：`sync_playwright` → `storage_state=auth_state.json` → 直链跳 `class_package_url` → 等 `添加课时包` 出现 → 复现到目标状态 → `all_text_contents()` 打印。

## 7. 失败诊断流程

接到「跑完发现 N 条 FAIL」时按顺序排查：

1. 打开 `logs/batch-<ts>.csv`，看「详情」列里的 toast 文本。
2. 如果详情是 `（未抓到 toast）`：去 `logs/submit-fail-*.png` 看实际界面，**红色 inline 错误**优先于 toast。
3. 如果详情包含 `OK` 但仍判 FAIL：成功关键词列表漏了，补 `submit_and_verify` 里的 `kw` 元组。
4. 如果详情是 `级联下拉 ... 第 2 级找不到选项: X`：跑 `debug_dropdown_options.py` 把真实选项打出来对齐。
5. 如果详情包含「课包类型不能为空」：是被课包分类反向清空了，检查 `_read_select_value` 校验逻辑是否生效。
6. 任意阶段 `LoginExpiredError` / 跳到登录页：删 `auth_state.json` 重扫码。

## 8. 一次完整任务的最小命令序列

```powershell
# 0. 切到工程
cd c:\Users\fengjianyi\Desktop\crm-lesson-package-skill
$env:PYTHONIOENCODING = "utf-8"

# 1. 确认登录态
python crm_create_lesson_package.py --dry-run

# 2. 小批验证（前 2 条）
python crm_batch_create_lesson_packages.py --xlsx "<path>" --limit 2

# 3. 看 logs/batch-*.csv，确认成功；有 FAIL 先按 §7 修

# 4. 全量跑（带去重）
python crm_batch_create_lesson_packages.py --xlsx "<path>" --skip-existing

# 5. 失败重试（如果有）
python crm_batch_create_lesson_packages.py --xlsx "<path>" --start <失败行号> --skip-existing
```
