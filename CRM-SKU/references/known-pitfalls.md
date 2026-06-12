# 已知坑与解决方案

按踩过的顺序记录，新增问题先来这里查。

## 1. 课时包创建

### 1.1 找不到课包类型一级选项「常规正课」

**现象**：脚本运行到课包类型时报 `级联下拉 '课包类型' 第 1 级找不到选项: 常规正课`

**原因**：CRM 课包类型用的是 `el-cascader`（级联）而不是普通 `el-select`，需要点击 `.el-cascader` 容器，遍历 `.el-cascader-node` 而不是 `.el-select-dropdown__item`

**已修**：`utils/element_ui.py` 的 `fill_cascade_dropdown` 函数。

### 1.2 课包类型「年课包pro」找不到

**现象**：Excel 写「年课包pro」，CRM 中只有「年课包」选项

**原因**：业务命名差异

**已修**：`crm_batch_create_lesson_packages.py` 的 `parse_excel` 中加了映射：
```python
mapping = {
    "年课包pro": "年课包",
    "两年包": "两年课包",
    "其他": "其他类型课包",
}
```
新类型直接扩展该字典。

### 1.3 「打卡次数 / 停课次数」字段一开始没有

**现象**：第一次填表时 form 里只有 11 个字段

**原因**：这两个字段是**选完课包类型才出现的**

**已修**：`fill_form` 顺序是「名称 → 类型 → 分类 → 数字字段」，后续字段用 `if "checkin_count" in data:` 守。

### 1.4 课包类型选完后被课包分类反向清空

**现象**：填课包分类后，提交时报「课包类型必填」

**已修**：`fill_form` 里先填类型再填分类，填完读 `_read_select_value` 校验，被清空就重选。

### 1.5 成功 toast 的文案是 `OK` 而不是中文

**现象**：`.el-message` 弹的是 `OK`，但脚本判断成功关键词只看「保存成功 / 创建成功」，全跑结果显示 100% FAIL，但 CRM 里实际已创建

**已修**：`submit_and_verify` 的成功关键词列表加入 `OK / ok / Success / success` 变体。

### 1.6 列表页搜索框元素被其他元素遮挡

**现象**：`Locator.click: ... intercepts pointer events`

**已修**：用 `force=True` 点击。

## 2. 套餐创建

### 2.1 服务协议下拉打不开

**现象**：点击 input 后 `el-select-dropdown:visible` 数量为 0

**原因**：服务协议是搜索式下拉，需要先输入关键字才能筛选；同时点击经常被父级 form-item 遮挡，需要 `force=True`。

**已修**：`fill_search_select` 函数；服务协议的 `search_keyword="0403"`（取协议名末尾数字过滤）。

### 2.2 课时包搜索建议没有

**现象**：在嵌套对话框输入完课时包名后下拉为空

**原因**：CRM 用的是 el-autocomplete，输入后需要等 1.5 秒才返回建议。

**已修**：`add_class_package` 函数等待 `1800ms` 后查找建议项；建议文本格式是 `常规正课 / 中课包 / 【VIPTHINK】xxx`。

### 2.3 赠送礼品 checkbox 找不到

**现象**：找到 44 个 `.el-checkbox` 但 innerText 都是空字符串

**原因**：`el-checkbox` 本身只包含图标，礼品文本在 `.el-tree-node__label` 这种相邻元素。直接搜 checkbox 文本永远找不到。

**已修**：用 JS 找**直接含礼品名文本**的元素（`childNodes` 中的 text node），再向上找 `.el-checkbox__inner` 祖先点击。详见 `add_gift` 函数。

### 2.4 礼品勾选验证 innerText 为空

**现象**：勾选成功后查 `.el-checkbox.is-checked`，但所有 `.el-checkbox.is-checked` 的 innerText 都为空，验证失败

**原因**：同上，文本不在 checkbox 元素内

**已修**：验证时也是向上找祖先看是否含礼品名。

### 2.5 关闭礼品弹窗

**现象**：礼品弹窗找不到「确定」按钮

**原因**：礼品弹窗是直接选完即生效，没有确认按钮

**已修**：点击主弹窗的 `.el-dialog__title` 区域来关闭，或按 ESC。

### 2.6 「是否换课」label 是空的

**现象**：用 `:text-is('是否换课')` 找 label，找不到

**原因**：「是否换课」开关的 form-item 没有 label，文本「是否换课」是 form-item 同级的 div / span 文本

**已修**：用 JS 找含「是否换课」直接文本的元素，向上找 `.el-switch`：
```js
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
// 向上找 .el-switch
```

### 2.7 是否换课开关切换后下拉里没有「海外换课」

**现象**：开关切换瞬间检查下拉，所有 5 个 readonly「请选择」input 都没有「海外换课」选项

**原因**：开关切换有动画 + UI 渲染时间，需要等 1.2 秒

**已修**：开关切换后 `page.wait_for_timeout(1200)`。

### 2.8 「重复购买次数」input 选错了

**现象**：填到了「使用入口」的 input，保存时报「请输入次数」

**原因**：之前用 form-item 祖先太宽泛，包含了相邻字段的 input。错把「请选择使用入口」placeholder 的 input 当成数字框。

**已修**：精确找含「此套餐允许重复购买次数」直接文本的元素的 form-item 祖先；过滤 input 的 placeholder 不能含「请选择」（5 字以上的）。

### 2.9 重复购买次数 input 是 `aria-disabled=true` 但 `disabled=false`

**现象**：开关切开后，input role=spinbutton 仍然 `aria-disabled="true"`，Playwright `click` 报 `element is not enabled` 一直重试

**原因**：el-input-number 的视觉禁用属性，但实际可填值

**已修**：去掉 `aria-disabled` 检查，直接 JS 注入 value：
```js
const desc = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value');
desc.set.call(inp, String(count));
inp.dispatchEvent(new Event('input', {bubbles: true}));
inp.dispatchEvent(new Event('change', {bubbles: true}));
```

### 2.10 转介绍规则的实际选项

**现象**：文档里写「其他 - A0B0」，但 CRM 里实际选项是「全年课（A0B0）」

**已修**：`package_conf["referral_rule"] = "全年课（A0B0"`（用部分前缀匹配，避免不同括号字符差异）。

## 3. 通用

### 3.1 Windows GBK 编码报 UnicodeEncodeError

**已修**：脚本入口加：
```python
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
```
外部加 `set PYTHONIOENCODING=utf-8`。

### 3.2 多次失败后弹窗残留

**现象**：上一条失败后弹窗没关，下一条点「添加」失败

**已修**：每条循环开头 `page.goto(class_package_url)` 强刷新；失败分支 `finally` 里 ESC 关弹窗。

### 3.3 登录态过期

**现象**：跑批中途跳转到登录页

**解决**：删 `auth_state.json`，重新跑 `--use-password`。

### 3.4 页面加载未完成就操作

**现象**：找按钮 / 字段超时

**已修**：所有关键步骤间加 `wait_for_timeout(1500-2500)`，避免用 `networkidle`（CRM 后台轮询导致一直不 idle）。

## 4. 失败诊断顺序

接到「跑完发现 N 条 FAIL」时按顺序排查：

1. 打开 `logs/<batch>-<ts>.csv`，看「详情」列
2. 详情里有 `OK` 但仍判 FAIL → 关键词列表漏了，补 `submit_and_verify` 的 `kw` 元组
3. 详情含「不能为空 / 必填」→ 是字段没填中或被反向清空
4. 详情含「找不到选项 / 找不到 checkbox」→ UI 改版或文案变化，对照 [element-ui-patterns.md](element-ui-patterns.md) 调整定位
5. 详情含「请输入次数」→ 重复购买次数 input 选错了，看 §2.8
6. 详情是「（未抓到 toast）」→ 看 `logs/<package>-fail-*.png` 截图
7. 任意阶段跳到登录页 → 删 `auth_state.json` 重跑 `--use-password`

排查不出来时：用 [workflow.md §2 录制方法](workflow.md) 重新录制，对照真实操作修脚本。
