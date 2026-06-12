# 完整工作流

CRM-SKU 的标准两步流程 + UI 改版时的录制方法。

## 1. 标准流程：先课包再套餐

### 0. 准备

1. 拿到 Excel 配置表（如 `美澳-6月.xlsx`），按 [excel-schema.md](excel-schema.md) 检查格式
2. 确认 CRM 中所需的赠送礼品已存在（人工确认或在 CRM 搜一次）
3. 确认 `config.local.json` 的 URL 指向目标环境（默认是生产 `https://crm.vipthink.cn`）

### 1. 创建课时包

```bash
cd .claude/skills/CRM-SKU/scripts

# 首跑：用账号密码登录（首次会保存 auth_state.json）
PYTHONIOENCODING=utf-8 python crm_batch_create_lesson_packages.py \
  --xlsx "<Excel 路径>" --limit 1 --use-password
```

观察第 1 条结果：
- 浏览器中能看到表单完整填充并保存成功
- `logs/batch-<ts>.csv` 第 1 行为 `OK`

第 1 条通过后，跑全量：

```bash
PYTHONIOENCODING=utf-8 python crm_batch_create_lesson_packages.py \
  --xlsx "<Excel 路径>" --skip-existing --use-password
```

如果有失败：
- 看 `logs/batch-<ts>.csv` 的「详情」列
- 看 `logs/submit-fail-*.png` 截图
- 对照 [known-pitfalls.md](known-pitfalls.md)

### 2. 创建套餐

```bash
PYTHONIOENCODING=utf-8 python crm_batch_create_packages.py \
  --xlsx "<Excel 路径>" --limit 1 --use-password
```

第 1 条通过后跑全量（套餐脚本不需要 `--skip-existing`，因为套餐管理列表的去重逻辑暂未实现，重跑会创建重复套餐，注意从 `--start N` 跑剩余的）：

```bash
PYTHONIOENCODING=utf-8 python crm_batch_create_packages.py \
  --xlsx "<Excel 路径>" --use-password
```

### 3. 验证

打开 CRM：
- `财务 → 商品管理 → 课时包管理`，搜索每个课时包名称
- `财务 → 商品管理 → 套餐管理`，搜索每个套餐名称（与课时包同名）
- 抽查 1-2 个套餐，点编辑确认所有字段正确（赠送礼品 / 是否换课 / 重复购买次数 / 转介绍规则）

### 4. 失败重试

```bash
# 课时包失败重试（从第 N 条）
python crm_batch_create_lesson_packages.py --xlsx "<path>" --start <N> --skip-existing --use-password

# 套餐失败重试
python crm_batch_create_packages.py --xlsx "<path>" --start <N> --use-password
```

## 2. UI 改版时的录制方法

CRM 改版后定位失效，用 `record_user_actions_v2.py` 录制真实操作，对照修复脚本。

### 录制步骤

```bash
cd .claude/skills/CRM-SKU/scripts
PYTHONIOENCODING=utf-8 python record_user_actions_v2.py
```

录制脚本会：
- 用账号密码登录
- 跳转到「套餐管理」页
- 注入 JS 监听 click / input / change / keydown 事件
- 每 3 秒打印新增事件，并实时把日志写入 `logs/recording-v2-<ts>.json`

### 操作

在浏览器中**手动**完成一个套餐的完整创建（按业务要求）。每次点击 / 输入都会被记录，包括：
- 元素 tag / class / id
- 文本 / placeholder / value
- 所在 form-item 的 label 和子元素状态（switch=ON/OFF, input=val/ph/ro/dis）
- 元素 outerHTML 前 200 字符
- 是否在下拉/弹层中

操作完成后 Ctrl+C 结束录制。

### 分析录制日志

```python
import json
with open('logs/recording-v2-<ts>.json', encoding='utf-8') as f:
    events = json.load(f)
for i, e in enumerate(events, 1):
    t = e['target']
    print(f"{i:3d}. {e['event']:6s} {t['tag']:6s} [{t['label'] or t['placeholder']}] '{t['text'][:40]}'")
    if t['outer']:
        print(f"     outer: {t['outer']}")
```

重点观察：
- 哪些元素是真正的可点击目标（不是 SPAN 文本而是 .el-checkbox__inner 这种）
- 字段定位的祖先关系（form-item label 是空的还是有具体文本）
- 控件类型（el-cascader / el-tree / el-input-number / el-switch）
- 关键时机（开关切换后哪些 input 才出现）

### 修复脚本

对照录制日志修改 `crm_batch_create_packages.py` 的对应函数：
- `add_class_package` - 课时包搜索
- `add_gift` - 赠送礼品（树形 checkbox）
- `fill_course_change` - 是否换课开关 + 规则
- `fill_repeat_purchase` - 重复购买次数开关 + 数字
- `fill_search_select` / `fill_simple_select` - 各类下拉

修改完后跑 `--limit 1` 验证。

## 3. 跨月 / 跨区域复用

每月 / 每区域只需准备新的 Excel 配置表，无需改脚本。

约定的目录组织（示例）：

```
CRM的课包配置/
├── 5月/
│   ├── 美澳-5月.xlsx
│   └── 港澳-5月.xlsx
├── 6月/
│   ├── 美澳-6月.xlsx
│   └── 港澳-6月.xlsx
└── ...
```

跑批命令统一：
```bash
PYTHONIOENCODING=utf-8 python crm_batch_create_lesson_packages.py \
  --xlsx "<具体月份的 xlsx>" --skip-existing --use-password
PYTHONIOENCODING=utf-8 python crm_batch_create_packages.py \
  --xlsx "<具体月份的 xlsx>" --use-password
```

仅当业务规则变化（默认套餐配置、礼品名称等）需要改脚本中的 `package_conf` 字典。
