# 已知坑与可复用片段

本文档记录 CRM 课时包自动化在实战中踩过的坑、根因、定位手法和最终用的代码片段。新会话接到「跑这个 skill 失败」时按本文档对照。

## 1. 成功 toast 文案是 `OK`

**现象**：批量跑完 16 条全部判 FAIL，但 CRM 后台实际看 12 条已创建。

**根因**：CRM 自定义了 `el-message`，文案就是字符串 `OK`，不是常见的「保存成功 / 创建成功 / 操作成功」。`submit_and_verify` 的关键词列表里漏了。

**修复**：

```python
if any(
    kw in m
    for m in captured
    for kw in (
        "保存成功", "创建成功", "操作成功", "新增成功", "添加成功",
        ".el-message: OK", ".el-message: ok",
        ".el-message: Success", ".el-message: success",
    )
):
    success = True
    break
```

**教训**：判定成功别只信「成功」二字。先抓一次失败截图，确认 toast 真正长啥样，再写关键词。

## 2. el-message 自动消失 ~3s

**现象**：`page.wait_for_selector(".el-message")` 偶尔抓不到。

**根因**：Element UI 的 `el-message` 默认 3s auto-dismiss，单次同步等待容易错过窗口。

**修复**：用 deadline + 200ms 轮询循环，10s 内累积所有 toast：

```python
captured = []
deadline = page.evaluate("Date.now()") + 10000
while page.evaluate("Date.now()") < deadline:
    for m in _collect_messages(page):
        if m not in captured:
            captured.append(m)
            print(f"  [MSG] {m}")
    if any(kw in m for m in captured for kw in SUCCESS_KEYWORDS):
        success = True
        break
    page.wait_for_timeout(200)
```

`_collect_messages` 同时扫 `.el-message` / `.el-notification` / `.el-form-item__error` / `.el-message-box`，因为不同失败模式 toast 可能落在不同容器。

## 3. 课包类型被课包分类反向清空

**现象**：填完类型再填分类，提交报「课包类型不能为空」。

**根因**：CRM 的 `el-select` 在级联场景下有联动重置逻辑，分类的 change 事件会清类型。

**修复**：填完分类后回读类型 input 值校验，不一致就重选：

```python
expected_leaf = pkg_type[-1] if isinstance(pkg_type, list) else pkg_type
if expected_leaf and expected_leaf not in _read_select_value(page, "课包类型"):
    print(f"  [WARN] 课包类型 被清空, 重新选择: {expected_leaf}")
    fill_cascade_dropdown(page, "课包类型", pkg_type)
    page.wait_for_timeout(500)
```

`_read_select_value`：

```python
def _read_select_value(page, label_text):
    return page.locator(
        f".el-form-item:has-text('{label_text}') input"
    ).first.input_value() or ""
```

## 4. 「打卡次数 / 停课次数」是动态字段

**现象**：一开始按现有字段填表，提交后报「打卡次数必填」。

**根因**：这俩字段是**选完课包类型后**才挂到 DOM 上，初始 `el-dialog` 里没有。

**修复**：

- 顺序硬编码：名称 → 课包类型 → 课包分类 → 数字字段。
- 写入时用 `if "checkin_count" in data` 守，缺字段就跳过。

诊断动态字段是否新增的一次性脚本：[`debug_after_type.py`](../../../Desktop/crm-lesson-package-skill/debug_after_type.py)，逻辑是「填完类型后打 `.el-dialog:visible .el-form-item__label` 全量内容」。

## 5. 级联二级选项文案不一致

**现象**：跑「两年包」时报 `级联下拉 '课包类型' 第 2 级找不到选项: 两年包`。

**根因**：Excel 里写「两年包」，CRM 真实选项可能叫「两年课包」「两年课包pro」之类。

**定位**：[`debug_dropdown_options.py`](../../../Desktop/crm-lesson-package-skill/debug_dropdown_options.py) 模板：

```python
type_form = pg.locator(".el-form-item:has-text('课包类型')").first
type_form.locator("input").first.click()
pg.wait_for_timeout(500)
pg.locator(".el-select-dropdown:visible .el-select-dropdown__item") \
  .filter(has_text="常规正课").first.click()
pg.wait_for_timeout(800)
for t in pg.locator(".el-select-dropdown:visible .el-select-dropdown__item").all_text_contents():
    print(repr(t.strip()))
```

**修复策略**：
- 优选改 Excel，保持脚本简单。
- 如果 Excel 的口径必须是用户语言，给 `TYPE_PARENT` 同时加一个 `LEAF_ALIAS`：`{"两年包": "两年课包"}` 之类，在 `parse_excel` 里替换。

## 6. Windows 控制台 GBK 编码

**现象**：脚本打印中文 / emoji 时 `UnicodeEncodeError: 'gbk' codec can't encode ...`。

**修复**（任选）：

入口加：
```python
import sys, io
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
```

或 PowerShell：
```powershell
$env:PYTHONIOENCODING = "utf-8"
```

**注意**：`python -c "..."` 一行式里用 `TextIOWrapper` 会触发 `I/O operation on closed file`，改用环境变量法。

## 7. 弹窗脏状态拖累后续条目

**现象**：第 N 条失败后，第 N+1 条直接走偏。

**根因**：上一条提交失败时 `el-dialog` 还开着，可能盖住列表的「添加课时包」按钮。

**修复**：

- 每条循环开头 `goto_list(page, url)` 重新刷列表。
- 失败分支无脑调 `close_dialog(page)`：先点「取消」，不行按 Escape，再不行就「确定」（应付二次确认弹窗）。

```python
def close_dialog(page):
    if page.locator(".el-dialog:visible").count() > 0:
        try:
            click_button_with_text(page, "取消")
        except Exception:
            page.keyboard.press("Escape")
        page.wait_for_timeout(300)
        if page.locator(".el-dialog:visible").count() > 0:
            try:
                click_button_with_text(page, "确定")
            except Exception:
                pass
            page.wait_for_timeout(300)
```

## 8. Excel 是「KV 横向交替」，不是表头 + 行

**现象**：第一次直接按 `ws.iter_rows` 当表头处理，全错位。

**根因**：用户的 Excel 每行是 `(label1, val1, label2, val2, ...)` 形式，从「课时包名称」开始一直到「适用课类」。还混杂了备注行如 `年课包pro请填写成年课包`。

**修复**：

```python
for raw in ws.iter_rows(values_only=True):
    if not raw or raw[0] is None:
        continue
    if str(raw[0]).strip() != "课时包名称":  # 行起始锚点
        continue
    pkg = {}
    for i in range(0, len(raw), 2):
        label = raw[i]
        value = raw[i + 1] if i + 1 < len(raw) else None
        if label is None:
            break
        field = HEADER_MAP.get(str(label).strip())
        if field:
            pkg[field] = value
    rows.append(pkg)
```

新增字段：扩 `HEADER_MAP` 即可，不用改 parse 逻辑。

## 9. 重跑前永远先评估「上次实际成功了几条」

**教训**：第一次跑 16 条因为 toast 关键词漏了，CSV 显示 16 FAIL，但其中 12 条 CRM 已实际创建。如果直接重跑就会有 12 条重名失败 + 4 条仍然失败。

**规则**：
1. 跑完先肉眼看 1~2 张 `submit-fail-*.png` 确认 toast 实际文本。
2. 任何重跑都加 `--skip-existing`。
3. `is_name_exists` 用列表搜索框查重，搜不到才允许新建。
