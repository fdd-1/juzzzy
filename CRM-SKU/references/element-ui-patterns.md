# Element UI 自动化模式

CRM 用的是 Element UI（Vue 2），各类控件的自动化定位方式不一样。本文档列出本 skill 实战验证过的模式，供未来类似项目参考。

## 1. 通用原则

- **优先用 label 文本定位**：`.el-form-item__label:text-is('字段名')` 比 class 更稳
- **避开 `:has-text` 在长 label 上**：会匹配到嵌套元素，用 `:text-is` 精确匹配
- **大量字段时用 `.el-dialog:visible.last`** 限定弹窗范围，避免点到背景列表的元素
- **JS evaluate 比 Playwright locator 更灵活**：复杂 DOM 结构（树/嵌套）直接用 `page.evaluate()` 写一次性查找逻辑
- **用 `.el-form-item.is-error` 检查必填校验**：保存失败时优先看红色边框

## 2. 控件模式

### el-input（普通文本框）

```python
def fill_text(page, label_text, value):
    form_item = get_form_item_by_label(page, label_text)
    inp = form_item.locator("input").first
    inp.click()
    inp.fill(str(value))
```

### el-input-number（数字框）

同 el-input，但注意：
- `aria-disabled="true"` 但 `disabled=false` 时，仍可填值（绕过检查）
- 直接用 JS 注入：

```python
page.evaluate("""
(value) => {
    const inp = ...找到目标 input...;
    const desc = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value');
    desc.set.call(inp, String(value));
    inp.dispatchEvent(new Event('input', {bubbles: true}));
    inp.dispatchEvent(new Event('change', {bubbles: true}));
    inp.blur();
}
""", value)
```

### el-select（普通下拉）

点击 input → 等待下拉 → 点匹配项：

```python
def fill_simple_select(page, label_text, value):
    form_item = get_form_item_by_label(page, label_text)
    inp = form_item.locator("input").first
    inp.click(force=True)
    page.wait_for_timeout(800)
    # 在 .el-select-dropdown:visible 中点匹配项
    for opt in page.locator(".el-select-dropdown:visible .el-select-dropdown__item").all():
        if value in opt.inner_text().strip():
            opt.click()
            return
```

注意：
- 用 `force=True` 避免 input 被 cascader/其他元素遮挡
- 长名字下拉可能滚动，多次重试
- 选完后下拉自动关，不需要手动关

### el-select（搜索式下拉）

如「服务协议」：

```python
def fill_search_select(page, label_text, value, search_keyword):
    form_item = get_form_item_by_label(page, label_text)
    inp = form_item.locator("input").first
    inp.click(force=True)
    is_readonly = inp.get_attribute("readonly") is not None
    if not is_readonly and search_keyword:
        page.keyboard.type(search_keyword, delay=80)
        page.wait_for_timeout(1200)
    # 点匹配项
    click_dropdown_item(page, value)
```

服务协议要输入「0403」过滤后才能看到「豌豆益智直播课合同（海外-不含教具）-20260403」。

### el-cascader（级联选择器）

如「课包类型」（常规正课 → 中课包）：

**关键差异**：el-cascader **不是 el-select**，要点 `.el-cascader` 容器而不是 input：

```python
form_item = page.locator(f".el-form-item:has-text('{label_text}')").first
cascader = form_item.locator(".el-cascader").first
cascader.click(timeout=timeout)
page.wait_for_timeout(1000)

# 一级
nodes = page.locator(".el-cascader-node").all()
for node in nodes:
    if "常规正课" in node.inner_text():
        node.click()
        break

# 二级（点完一级后会出现新的 el-cascader-node）
page.wait_for_timeout(500)
nodes = page.locator(".el-cascader-node").all()
for node in nodes:
    if "中课包" in node.inner_text():
        node.click()
        break
```

### el-autocomplete（输入建议）

如「课时包」搜索：

```python
search = nested_dialog.locator("input[placeholder*='课时包']").first
search.click()
search.fill(package_name)
page.wait_for_timeout(1800)

# 建议项是 li，文本格式「常规正课 / 中课包 / xxx」
items = page.locator(".el-autocomplete-suggestion:visible li, .el-popper:visible li").all()
for it in items:
    if package_name in it.inner_text():
        it.click()
        break
```

### el-tree + checkbox（树形复选框）

如「赠送礼品」（商城积分 → 海外特批16000豌豆币）：

**关键发现**：
- el-checkbox 本身 `innerText` 为空（文本不在 checkbox 内部）
- 礼品名是相邻的 `.el-tree-node__label` 或 `__label__text` 文本
- 必须**先找含礼品名直接文本的元素，再向上找 `.el-checkbox__inner` 祖先**

```python
js_result = page.evaluate(f"""
() => {{
    const giftName = "{gift_name}";
    // 找直接含礼品名文本的元素
    const all = document.querySelectorAll('*');
    for (const el of all) {{
        if (el.offsetParent === null) continue;
        const direct = Array.from(el.childNodes)
            .filter(n => n.nodeType === 3)
            .map(n => n.textContent.trim())
            .join('');
        if (direct === giftName || direct.includes(giftName)) {{
            // 向上找含 .el-checkbox__inner 的祖先
            let cur = el;
            for (let i = 0; i < 8; i++) {{
                cur = cur.parentElement;
                if (!cur) break;
                const cb = cur.querySelector('.el-checkbox__inner');
                if (cb && (cur.innerText || '').trim().length < giftName.length + 30) {{
                    cb.click();
                    return {{ ok: true }};
                }}
            }}
        }}
    }}
    return {{ ok: false }};
}}
""")
```

**验证**：点击后查 `.el-checkbox.is-checked` 是否包含礼品名（也是向上找祖先文本）。

### el-switch（开关）

如「是否换课」、「此套餐允许重复购买次数」：

**关键发现**：开关旁的字段名 form-item 的 label 经常是空的，文本是兄弟节点：

```python
# JS 找含字段名直接文本的元素，向上找 .el-switch
result = page.evaluate("""
() => {
    const all = document.querySelectorAll('*');
    let target = null;
    for (const el of all) {
        const direct = Array.from(el.childNodes)
            .filter(n => n.nodeType === 3)
            .map(n => n.textContent || '')
            .join('');
        if (direct.includes('是否换课')) {
            target = el;
            break;
        }
    }
    if (!target) return { ok: false };

    let cur = target;
    for (let i = 0; i < 8; i++) {
        cur = cur.parentElement;
        if (!cur) break;
        const sw = cur.querySelector('.el-switch');
        if (sw) {
            const isOn = sw.classList.contains('is-checked');
            if (!isOn) sw.click();
            return { ok: true };
        }
    }
    return { ok: false };
}
""")
```

开后等 1.2 秒让 UI 渲染新出现的字段（如换课规则下拉、次数 input）。

## 3. 弹窗管理

- `page.locator(".el-dialog:visible").first` - 主弹窗
- `page.locator(".el-dialog:visible").last` - 当前最顶层弹窗（嵌套时是子弹窗）
- 关闭弹窗：找「确定」/「取消」按钮，找不到就 `page.keyboard.press("Escape")`
- 礼品弹窗特殊：没有确认按钮，点主弹窗标题即可关

## 4. 提交验证

提交后：
1. 等弹窗关闭：`page.locator(".el-dialog:visible").count() == 0`
2. 看成功 toast：`.el-message: OK` / `保存成功` / `创建成功`
3. 看错误：`.el-form-item.is-error`、`.el-message`、`.el-form-item__error`

实战发现成功 toast 的文案是 `OK` 而不是中文 → 关键词列表必须包含变体。

## 5. 常见陷阱

| 陷阱 | 解决 |
|---|---|
| `inner_text()` 在 element 在动画中时超时 | 一次性 `all_inner_texts()` 或 `evaluate()` 内部处理 |
| `:visible` 在某些版本下不识别 | 用 `el.offsetParent !== null` 在 JS 中判可见 |
| 选项文案前后空格 | `.trim()` 后比对 |
| `force=True` 后仍卡住 | element 被遮挡，先 `keyboard.press("Escape")` 关其他弹层 |
| 下拉打开后立刻 `inner_text()` 拿不到 | 等 800-1200ms |
| 选完一项后下拉关闭 → 二次查找失败 | 找到立即点击，不要拆成两步 |
