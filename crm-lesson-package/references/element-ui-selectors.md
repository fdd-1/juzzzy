# Element UI 选择器手册

本文档列出在 CRM 课时包表单上常用的 Playwright 选择器和写法。新会话遇到 CRM 其他表单想复用思路，先看这里。

## 1. 表单容器

```python
form_item = page.locator(f".el-form-item:has-text('{label}')").first
```

`:has-text` 比按 `for` 属性匹配更稳，因为 CRM 自定义 label 经常没有 `for`。

## 2. 文本框（el-input）

```python
form_item.locator("input").first.fill(value)
```

数字输入框（`el-input-number`）也是 `input` 标签，同上。注意一定 `fill(str(value))`，传 int 会报错。

## 3. 可搜索下拉（el-select filterable）

```python
input_locator = form_item.locator("input").first
input_locator.click()
page.wait_for_timeout(300)
is_readonly = input_locator.get_attribute("readonly") is not None
if not is_readonly:
    page.keyboard.press("Control+A")
    page.keyboard.press("Delete")
    page.keyboard.type(value, delay=50)
    page.wait_for_timeout(300)

# 在可见的下拉面板里点
for opt in page.locator(".el-select-dropdown:visible .el-select-dropdown__item").all():
    if value in opt.inner_text().strip():
        opt.click()
        break
```

`readonly` 区分 filterable / 非 filterable：filterable 要打字搜索，否则直接点。

## 4. 二级级联（el-select 嵌套，不是 el-cascader）

CRM 的「课包类型」用的是**两层 el-select**，不是标准 `el-cascader`：点完一级后，下拉会**替换**为二级选项（DOM 还是同一个 `.el-select-dropdown`）。所以写法是：

```python
input_locator.click()
page.wait_for_timeout(500)

# 一级
for opt in page.locator(".el-select-dropdown:visible .el-select-dropdown__item").all():
    if level1 in opt.inner_text().strip():
        opt.click()
        page.wait_for_timeout(400)
        break

# 二级（同一选择器，DOM 已被替换）
for opt in page.locator(".el-select-dropdown:visible .el-select-dropdown__item").all():
    if level2 in opt.inner_text().strip():
        opt.click()
        page.wait_for_timeout(400)
        break
```

如果换了真正的 `el-cascader`，备用选择器：`.el-cascader-panel:visible .el-cascader-node`。

## 5. 多选下拉（el-select multiple）

多选下拉点选项后**不会自动收起**，逐个点完后用 Escape 收起：

```python
input_locator.click()
page.wait_for_timeout(300)
for v in values:
    opt = page.locator(".el-select-dropdown:visible .el-select-dropdown__item") \
            .filter(has_text=v).first
    opt.click()
    page.wait_for_timeout(150)
page.keyboard.press("Escape")
```

## 6. 弹窗按钮

```python
def click_button_with_text(page, text):
    # 「确定」可能渲染成「确 定」（中文 2 字按钮 letter-spacing）
    candidates = [text]
    if len(text) == 2:
        candidates.append(f"{text[0]} {text[1]}")
    for t in candidates:
        for sel in ["button:has-text('{}')", ".el-button:has-text('{}')"]:
            try:
                page.locator(sel.format(t)).first.click(timeout=5000)
                return
            except Exception:
                continue
    raise ValueError(f"找不到按钮: {text}")
```

按钮文案不确定时用 [`debug_save_button.py`](../../../Desktop/crm-lesson-package-skill/debug_save_button.py) 探：

```python
btns = pg.locator(".el-dialog:visible .el-dialog__footer button").all_text_contents()
```

## 7. Toast / 错误消息

四个常用容器，按可见性扫一遍：

```python
for sel in [".el-message", ".el-notification",
            ".el-form-item__error", ".el-message-box"]:
    for el in page.locator(sel).all():
        if el.is_visible():
            txt = (el.inner_text() or "").strip()
            if txt:
                yield f"{sel}: {txt}"
```

`el-form-item__error` 是 inline 红色错误，提交校验失败优先看这个。

## 8. 列表查重

```python
search_input = page.locator(
    ".el-form-item:has-text('课时包名称') input, "
    "input[placeholder*='课时包名称']"
).first
search_input.fill(name)
click_button_with_text(page, "查询")
page.wait_for_timeout(800)
exists = page.locator(f".el-table:has-text('{name}')").first.is_visible()
```

完整名匹配比模糊更准；如果名字里含特殊字符（CSS 选择器会爆），换 `:text-is`：

```python
page.locator(f".el-table tr").filter(has_text=name).first.is_visible()
```
