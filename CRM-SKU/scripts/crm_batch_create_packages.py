#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CRM 套餐批量创建脚本（基于录制 v2 日志重写）

关键发现：
- 赠送礼品：点击搜索结果是 SPAN（不是 li），选完后点主弹窗任意位置关闭
- 是否换课：form-item 没有 label，input 是 readonly 的「请选择」
- 重复购买次数：form-item 也没有 label，input 没有 placeholder，初始 val='0'
- 服务协议：搜索式，输入「0403」过滤
"""

import sys
import io
import json
import csv
import re
from pathlib import Path
from datetime import datetime
import openpyxl
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

sys.path.insert(0, str(Path(__file__).parent))
from utils.auth import get_credentials
from utils.precheck import precheck_package_excel, print_report, verify_in_list

SCRIPT_DIR = Path(__file__).parent
CONFIG_FILE = SCRIPT_DIR / "config.local.json"
AUTH_STATE_FILE = SCRIPT_DIR / "auth_state.json"
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

HEADER_MAP = {"课时包名称": "name", "赠送礼品": "gift_items"}


def load_config():
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(f"配置文件不存在: {CONFIG_FILE}")
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


def parse_excel(excel_path):
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active
    rows = []
    for raw in ws.iter_rows(values_only=True):
        if not raw or raw[0] is None:
            continue
        if str(raw[0]).strip() != "课时包名称":
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
        if pkg:
            rows.append(pkg)
    return rows


# ============= 通用工具 =============

def get_form_item_by_label(page, label_text):
    """精确按 label 文本定位 form-item"""
    scope = page.locator(".el-dialog:visible").last
    label = scope.locator(f".el-form-item__label:text-is('{label_text}')").first
    form_item = label.locator("xpath=ancestor::div[contains(@class, 'el-form-item')]").first
    return form_item


def click_dropdown_item(page, text, exact=False, timeout=5000):
    """点击当前可见下拉/弹层中匹配文本的项

    覆盖：el-select-dropdown__item / el-popper li / el-cascader-node /
           el-autocomplete-suggestion / 礼品弹窗中的 SPAN
    """
    deadline = page.evaluate("Date.now()") + timeout

    while page.evaluate("Date.now()") < deadline:
        candidates = []
        for sel in [
            ".el-select-dropdown:visible .el-select-dropdown__item",
            ".el-popper:visible li",
            ".el-cascader-panel:visible .el-cascader-node",
            ".el-autocomplete-suggestion:visible li",
        ]:
            for el in page.locator(sel).all():
                try:
                    if el.is_visible():
                        candidates.append(el)
                except:
                    pass

        for el in candidates:
            try:
                t = el.inner_text().strip()
                if exact and t == text:
                    el.click()
                    page.wait_for_timeout(400)
                    return True
                if not exact and text in t:
                    el.click()
                    page.wait_for_timeout(400)
                    return True
            except:
                continue
        page.wait_for_timeout(300)
    return False


def fill_text(page, label_text, value):
    form_item = get_form_item_by_label(page, label_text)
    inp = form_item.locator("input").first
    inp.click()
    inp.fill(str(value))
    page.wait_for_timeout(300)


def fill_simple_select(page, label_text, value):
    form_item = get_form_item_by_label(page, label_text)
    inp = form_item.locator("input").first
    inp.click(force=True)
    page.wait_for_timeout(800)
    if not click_dropdown_item(page, value):
        raise ValueError(f"找不到选项「{value}」 字段=「{label_text}」")


def fill_search_select(page, label_text, value, search_keyword):
    """搜索式：先点击 input，输入关键字过滤，再选项"""
    form_item = get_form_item_by_label(page, label_text)
    inp = form_item.locator("input").first
    inp.click(force=True)
    page.wait_for_timeout(500)

    is_readonly = inp.get_attribute("readonly") is not None
    if not is_readonly and search_keyword:
        page.keyboard.type(search_keyword, delay=80)
        page.wait_for_timeout(1200)

    if not click_dropdown_item(page, value):
        raise ValueError(f"找不到选项「{value}」 字段=「{label_text}」")


def add_class_package(page, package_name):
    print(f"  [步骤] 添加课时包: {package_name}")
    main_dialog = page.locator(".el-dialog:visible").first
    main_dialog.locator("button:has-text('添加课时包'), span:has-text('添加课时包')").first.click()
    page.wait_for_timeout(1500)

    nested = page.locator(".el-dialog:visible").last
    search = nested.locator("input[placeholder*='课时包']").first
    search.click()
    search.fill(package_name)
    page.wait_for_timeout(1800)

    if not click_dropdown_item(page, package_name, timeout=5000):
        raise ValueError(f"找不到课时包建议项: {package_name}")
    page.wait_for_timeout(1200)
    print(f"  ✓ 已选择课时包")

    nested.locator("button:has-text('确定'), .el-button:has-text('确定')").first.click()
    page.wait_for_timeout(2000)
    print(f"  ✓ 课时包已添加")


def add_gift(page, gift_name):
    """赠送礼品（树形复选框列表）

    流程（来自录制 v2 + 截图验证）：
    1. 点击赠送礼品区域的「请选择」（DIV）
    2. 在弹窗中找到「请输入关键字」输入礼品名
    3. 点击叶子节点礼品前的复选框（.el-checkbox__inner）
    4. 截图验证复选框已勾选（.el-checkbox.is-checked 包含礼品名）
    5. 关闭礼品弹窗
    """
    if not gift_name or str(gift_name).strip() in ("无", "None", ""):
        return

    print(f"  [步骤] 添加赠送礼品: {gift_name}")

    form_item = get_form_item_by_label(page, "赠送礼品")
    clicked = False
    for sel in ["div:has-text('请选择')", "span:has-text('请选择')", "input"]:
        try:
            form_item.locator(sel).first.click(force=True, timeout=2000)
            clicked = True
            break
        except:
            continue
    if not clicked:
        raise ValueError("无法点开赠送礼品弹窗")
    page.wait_for_timeout(1500)

    # 在礼品弹窗输入关键字
    search = page.locator("input[placeholder*='请输入关键字']:visible").first
    search.click()
    search.fill(str(gift_name))
    page.wait_for_timeout(1500)

    # 点击叶子节点礼品的复选框（el-tree 树组件）
    js_result = page.evaluate(f"""
    () => {{
        const giftName = "{gift_name}";

        // 全局搜索：找含礼品名的所有可见元素
        const allEls = document.querySelectorAll('*');
        const matched = [];
        for (const el of allEls) {{
            if (el.offsetParent === null) continue;
            const direct = Array.from(el.childNodes)
                .filter(n => n.nodeType === 3)
                .map(n => n.textContent.trim())
                .join('');
            if (direct === giftName) {{
                matched.push({{
                    tag: el.tagName,
                    class: (el.className || '').toString().substring(0, 80),
                    text: direct.substring(0, 50)
                }});
            }}
        }}

        // 找到所有可能的容器后，向上找 checkbox
        const candidates = [];
        for (const el of allEls) {{
            if (el.offsetParent === null) continue;
            const direct = Array.from(el.childNodes)
                .filter(n => n.nodeType === 3)
                .map(n => n.textContent.trim())
                .join('');
            if (direct === giftName || direct.includes(giftName)) {{
                candidates.push(el);
            }}
        }}

        for (const el of candidates) {{
            // 向上找含 checkbox 的祖先
            let cur = el;
            for (let i = 0; i < 8; i++) {{
                cur = cur.parentElement;
                if (!cur) break;
                const cb = cur.querySelector('.el-checkbox__inner, input[type="checkbox"]');
                if (cb) {{
                    // 限制：祖先文本不能太长（防止勾到根容器）
                    const ancestorText = (cur.innerText || '').trim();
                    if (ancestorText.length < giftName.length + 30) {{
                        try {{ cb.click(); }} catch(e) {{}}
                        return {{ ok: true, mode: 'ancestor', text: ancestorText.substring(0, 80), level: i+1 }};
                    }}
                }}
            }}
        }}

        return {{
            ok: false,
            matched_direct_count: matched.length,
            matched_examples: matched.slice(0, 5)
        }};
    }}
    """)
    print(f"  [DEBUG] 礼品 checkbox 点击: {js_result}")
    if not js_result.get("ok"):
        ts_dbg = datetime.now().strftime("%H%M%S")
        safe_dbg = re.sub(r'[^\w\-]', '_', gift_name)[:30]
        page.screenshot(path=str(LOG_DIR / f"gift-not-found-{safe_dbg}-{ts_dbg}.png"), full_page=True)
        raise ValueError(f"找不到礼品 checkbox: {gift_name}")

    page.wait_for_timeout(1000)

    # 验证：检查 .el-checkbox.is-checked 中是否包含礼品名（向上找祖先文本）
    verify_result = page.evaluate(f"""
    () => {{
        const giftName = "{gift_name}";
        const checked = Array.from(document.querySelectorAll('.el-checkbox.is-checked'))
            .filter(el => el.offsetParent !== null);
        for (const cb of checked) {{
            // 向上找含礼品名文本的祖先
            let cur = cb;
            for (let i = 0; i < 8; i++) {{
                cur = cur.parentElement;
                if (!cur) break;
                const text = (cur.innerText || '').trim();
                if (text.includes(giftName)) {{
                    // 检查这个祖先是否就是礼品所在节点（文本不太长）
                    if (text.length < giftName.length + 50) {{
                        return {{ ok: true, text: text.substring(0, 60) }};
                    }}
                }}
            }}
        }}
        return {{ ok: false, total_checked: checked.length }};
    }}
    """)

    # 截图保存
    ts = datetime.now().strftime("%H%M%S")
    safe = re.sub(r'[^\w\-]', '_', gift_name)[:30]
    page.screenshot(path=str(LOG_DIR / f"gift-checked-{safe}-{ts}.png"))

    print(f"  [DEBUG] 礼品勾选验证: {verify_result}")
    if not verify_result.get("ok"):
        raise ValueError(f"礼品勾选验证失败: {verify_result}")

    print(f"  ✓ 已勾选礼品（验证通过）: {gift_name}")

    # 关闭礼品弹窗：点主弹窗标题
    page.wait_for_timeout(500)
    try:
        main_dialog = page.locator(".el-dialog:visible").first
        main_dialog.locator(".el-dialog__title").first.click()
        page.wait_for_timeout(800)
    except:
        try:
            page.keyboard.press("Escape")
            page.wait_for_timeout(500)
        except:
            pass

    print(f"  ✓ 礼品弹窗已关闭")


def find_form_item_by_neighbor_text(page, neighbor_text):
    """通过附近文本（如开关旁的 label 文本）找 form-item

    用于无 label 字段：通过附近的文字（如「此套餐允许重复购买次数」）找 form-item
    """
    scope = page.locator(".el-dialog:visible").last
    # 找含此文本的元素，向上找 form-item
    candidates = scope.locator(f"*:has-text('{neighbor_text}')").all()
    for c in candidates:
        try:
            text = c.inner_text().strip()
            # 只取直接包含、文本不太长的元素
            if neighbor_text in text and len(text) < len(neighbor_text) + 50:
                # 向上找 form-item
                form_items = c.locator("xpath=ancestor::div[contains(@class, 'el-form-item')]").all()
                if form_items:
                    return form_items[0]
        except:
            continue
    return None


def fill_course_change(page, rule_value):
    """是否换课开关 + 规则

    录制 68: click SPAN(empty) - 开关
    录制 69: click INPUT [请选择] readonly
    录制 70: click SPAN「海外换课-0节」
    """
    print(f"  [步骤] 设置换课规则: {rule_value}")

    # 第一步：用 JS 找含「是否换课」直接文本的元素，向上找开关
    switch_result = page.evaluate("""
    () => {
        const dialogs = Array.from(document.querySelectorAll('.el-dialog'))
            .filter(d => d.offsetParent !== null);
        if (dialogs.length === 0) return { ok: false, reason: 'no dialog' };
        const dialog = dialogs[dialogs.length - 1];

        // 找直接文本包含「是否换课」的元素
        const all = dialog.querySelectorAll('*');
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
        if (!target) return { ok: false, reason: 'no 是否换课 text' };

        // 向上找含 .el-switch 的祖先（最多 8 层）
        let cur = target;
        for (let i = 0; i < 8; i++) {
            cur = cur.parentElement;
            if (!cur) break;
            const sw = cur.querySelector('.el-switch');
            if (sw) {
                const isOn = sw.classList.contains('is-checked');
                if (!isOn) {
                    sw.click();
                    return { ok: true, action: 'clicked-switch' };
                }
                return { ok: true, action: 'already-on' };
            }
        }
        return { ok: false, reason: 'no switch nearby' };
    }
    """)
    print(f"  [DEBUG] 是否换课开关: {switch_result}")
    page.wait_for_timeout(1200)

    scope = page.locator(".el-dialog:visible").last
    target_clicked = False  # 直接在打开下拉时就点击

    # 优先：用 label「换课规则」找
    try:
        labels = scope.locator(".el-form-item__label:has-text('换课规则')").all()
        for lbl in labels:
            try:
                form_item = lbl.locator("xpath=ancestor::div[contains(@class, 'el-form-item')]").first
                inputs = form_item.locator("input:visible").all()
                for inp in inputs:
                    try:
                        inp.click(force=True)
                        page.wait_for_timeout(800)
                        items = page.locator(".el-select-dropdown:visible .el-select-dropdown__item, .el-popper:visible li:visible").all()
                        # 精确匹配优先
                        exact_match = None
                        contains_match = None
                        for it in items:
                            try:
                                t = it.inner_text().strip()
                                if t == rule_value:
                                    exact_match = it
                                    break
                                if rule_value in t and contains_match is None:
                                    contains_match = it
                            except:
                                continue
                        match = exact_match or contains_match
                        if match:
                            match.click()
                            page.wait_for_timeout(500)
                            target_clicked = True
                            break
                        page.keyboard.press("Escape")
                        page.wait_for_timeout(300)
                    except:
                        continue
                if target_clicked:
                    break
            except:
                continue
    except:
        pass

    # 兜底：所有 readonly 「请选择」input 试一遍
    if not target_clicked:
        readonly_inputs = scope.locator("input[placeholder='请选择'][readonly]:visible").all()
        print(f"  [DEBUG] 兜底：{len(readonly_inputs)} 个 readonly 「请选择」input")
        for inp in readonly_inputs:
            try:
                inp.click(force=True)
                page.wait_for_timeout(800)
                items = page.locator(".el-select-dropdown:visible .el-select-dropdown__item, .el-popper:visible li:visible").all()
                # 精确匹配优先
                exact_match = None
                contains_match = None
                for it in items:
                    try:
                        t = it.inner_text().strip()
                        if t == rule_value:
                            exact_match = it
                            break
                        if rule_value in t and contains_match is None:
                            contains_match = it
                    except:
                        continue
                match = exact_match or contains_match
                if match:
                    match.click()
                    page.wait_for_timeout(500)
                    target_clicked = True
                    break
                page.keyboard.press("Escape")
                page.wait_for_timeout(300)
            except:
                continue

    if not target_clicked:
        ts = datetime.now().strftime("%H%M%S")
        page.screenshot(path=str(LOG_DIR / f"course-change-debug-{ts}.png"), full_page=True)
        raise ValueError("找不到换课规则下拉或选项「海外换课-0节」")

    print(f"  ✓ 已设置换课规则")


def fill_repeat_purchase(page, count):
    """重复购买次数：开关 + 数字（form-item 没有 label）

    录制 73: click SPAN(empty) - 开关
    录制 74: click INPUT, val='0' - 数字框
    录制 79: click SPAN「此套餐允许重复购买次数」（确认文本）
    """
    print(f"  [步骤] 设置重复购买次数: {count}")

    # 第一步：开启开关
    sw_result = page.evaluate("""
    () => {
        const dialogs = Array.from(document.querySelectorAll('.el-dialog'))
            .filter(d => d.offsetParent !== null);
        if (dialogs.length === 0) return { ok: false };
        const dialog = dialogs[dialogs.length - 1];

        // 找含「此套餐允许重复购买次数」直接文本的元素
        const all = dialog.querySelectorAll('*');
        let target = null;
        for (const el of all) {
            const direct = Array.from(el.childNodes)
                .filter(n => n.nodeType === 3)
                .map(n => n.textContent || '')
                .join('');
            if (direct.includes('此套餐允许重复购买次数') || direct.includes('重复购买次数')) {
                target = el;
                break;
            }
        }
        if (!target) return { ok: false, reason: 'no text' };

        let cur = target;
        for (let i = 0; i < 8; i++) {
            cur = cur.parentElement;
            if (!cur) break;
            const sw = cur.querySelector('.el-switch');
            if (sw) {
                const isOn = sw.classList.contains('is-checked');
                if (!isOn) {
                    sw.click();
                    return { ok: true, switched: true };
                }
                return { ok: true, switched: false };
            }
        }
        return { ok: false, reason: 'no switch' };
    }
    """)
    print(f"  [DEBUG] 重复购买开关: {sw_result}")

    if not sw_result.get("ok"):
        raise ValueError(f"找不到「此套餐允许重复购买次数」开关: {sw_result}")

    # 等开关动画 + UI 渲染
    page.wait_for_timeout(1500)

    # 第二步：用 JS 直接 set value + dispatch 事件（精确定位）
    set_result = page.evaluate(f"""
    () => {{
        const count = {count};
        const dialogs = Array.from(document.querySelectorAll('.el-dialog'))
            .filter(d => d.offsetParent !== null);
        const dialog = dialogs[dialogs.length - 1];

        // 找含「此套餐允许重复购买次数」直接文本的元素
        const all = dialog.querySelectorAll('*');
        let target = null;
        for (const el of all) {{
            const direct = Array.from(el.childNodes)
                .filter(n => n.nodeType === 3)
                .map(n => n.textContent || '')
                .join('');
            if (direct.includes('此套餐允许重复购买次数')) {{
                target = el;
                break;
            }}
        }}
        if (!target) return {{ ok: false, reason: 'no text' }};

        // 在 target 的兄弟节点 / 父节点的相邻区域找 input（el-switch 后面的 input）
        // 策略：找含「此套餐允许重复购买次数」的容器，再向后兄弟方向找 input
        // 或者：找最近的 .el-input-number
        let cur = target;
        for (let i = 0; i < 6; i++) {{
            cur = cur.parentElement;
            if (!cur) break;
            // 在这个容器内找所有 input，优先选 enabled 且非 switch 内的
            const inputs = cur.querySelectorAll('input');
            const inputDetails = [];
            for (const inp of inputs) {{
                if (inp.offsetParent === null) continue;
                if (inp.type === 'checkbox') continue;
                if (inp.closest('.el-switch')) continue;
                inputDetails.push({{
                    ph: inp.placeholder || '',
                    type: inp.type,
                    disabled: inp.disabled,
                    aria_disabled: inp.getAttribute('aria-disabled'),
                    val: inp.value,
                    role: inp.getAttribute('role') || ''
                }});
            }}

            // 找 enabled 的 number/text input
            for (const inp of inputs) {{
                if (inp.offsetParent === null) continue;
                if (inp.type === 'checkbox') continue;
                if (inp.closest('.el-switch')) continue;
                if (inp.disabled) continue;
                // 注意：el-input-number 有时会有 aria-disabled='true' 但实际可填值
                // 不检查 aria-disabled

                // 不要使用入口/下拉的 input
                const ph = inp.placeholder || '';
                if (ph.includes('请选择') && ph.length > 5) continue;

                const desc = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value');
                desc.set.call(inp, String(count));
                inp.dispatchEvent(new Event('input', {{ bubbles: true }}));
                inp.dispatchEvent(new Event('change', {{ bubbles: true }}));
                inp.blur();
                return {{ ok: true, value: inp.value, ph: ph, level: i+1, role: inp.getAttribute('role') }};
            }}

            // 如果在这一层有 input 但都不合格，记录信息
            if (inputDetails.length > 0) {{
                return {{ ok: false, reason: 'inputs found but filtered', level: i+1, details: inputDetails }};
            }}
        }}

        return {{ ok: false, reason: 'no input in any ancestor' }};
    }}
    """)
    print(f"  [DEBUG] 重复购买填值: {set_result}")
    if not set_result.get("ok"):
        raise ValueError(f"填重复购买次数失败: {set_result}")

    page.wait_for_timeout(500)
    print(f"  ✓ 已设置重复购买次数")


def click_button(page, text):
    scope = page.locator(".el-dialog:visible").last
    for sel in [f"button:has-text('{text}')", f".el-button:has-text('{text}')",
                f"span:has-text('{text}')"]:
        try:
            scope.locator(sel).first.click(timeout=3000)
            return True
        except:
            continue
    raise ValueError(f"找不到按钮: {text}")


# ============= 主流程 =============

def create_one_package(page, data, conf):
    name = data.get("name", "Unknown")
    print(f"\n=== 创建套餐: {name} ===")

    try:
        page.locator("button:has-text('添加商品套餐')").first.click()
        page.wait_for_timeout(2000)

        print("【步骤 1】商品套餐信息")
        fill_text(page, "套餐名称", name)
        fill_simple_select(page, "套餐类型", "正课套餐")
        add_class_package(page, name)

        # 服务协议（搜索关键字「0403」）
        fill_search_select(page, "服务协议", conf["service_agreement"], search_keyword="0403")

        click_button(page, "下一步")
        page.wait_for_timeout(1800)

        print("【步骤 2】商品销售信息")
        fill_simple_select(page, "使用入口", conf["usage_entry"])

        gift = data.get("gift_items")
        if gift:
            add_gift(page, gift)

        if conf.get("course_change_rule"):
            fill_course_change(page, conf["course_change_rule"])

        # 转介绍规则（有 label）
        fill_simple_select(page, "转介绍规则", conf["referral_rule"])

        if conf.get("repeat_purchase_count"):
            fill_repeat_purchase(page, conf["repeat_purchase_count"])

        print("【保存】")
        click_button(page, "保存")
        page.wait_for_timeout(3000)

        if page.locator(".el-dialog:visible").count() == 0:
            print(f"✓ 创建成功")
            return "OK", "创建成功"

        msgs = []
        for el in page.locator(".el-message:visible, .el-form-item__error:visible").all():
            try:
                t = el.inner_text().strip()
                if t:
                    msgs.append(t)
            except:
                pass
        detail = " | ".join(msgs) if msgs else "弹窗未关闭"

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe = re.sub(r'[^\w\-]', '_', name)[:50]
        page.screenshot(path=str(LOG_DIR / f"package-fail-{safe}-{ts}.png"))
        print(f"✗ 失败: {detail}")
        return "FAIL", detail

    except Exception as e:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        safe = re.sub(r'[^\w\-]', '_', name)[:50]
        try:
            page.screenshot(path=str(LOG_DIR / f"package-error-{safe}-{ts}.png"))
        except:
            pass
        print(f"✗ 异常: {e}")
        return "FAIL", str(e)
    finally:
        for _ in range(3):
            try:
                if page.locator(".el-dialog:visible").count() == 0:
                    break
                page.keyboard.press("Escape")
                page.wait_for_timeout(300)
            except:
                break


def batch_create(excel_path, start=1, limit=None, use_password=False, strict_precheck=True):
    # 执行前校验
    ok, errors, warnings = precheck_package_excel(excel_path)
    print_report("执行前校验", ok, errors, warnings)
    if not ok:
        if strict_precheck:
            print("校验未通过，已中止。修复后重跑，或加 --no-precheck 跳过（不推荐）。")
            return
        print("校验未通过但已选择跳过，继续执行。")

    print(f"解析 Excel: {excel_path}")
    packages = parse_excel(excel_path)
    print(f"共 {len(packages)} 条配置\n")
    if not packages:
        return

    config_data = load_config()
    package_url = config_data["crm"]["package_manage_url"]
    login_url = config_data["crm"]["login_url"]

    package_conf = {
        "service_agreement": "豌豆益智直播课合同（海外-不含教具）-20260403",
        "usage_entry": "续费",
        "course_change_rule": "海外换课-0节",
        "repeat_purchase_count": 2,
        "referral_rule": "全年课（A0B0",
    }

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)

        if use_password:
            context = browser.new_context()
            page = context.new_page()
            print("使用账号密码登录...")
            page.goto(login_url, timeout=30000)
            page.wait_for_timeout(2000)
            try:
                # 凭据：env > config.username > 交互式输入（密码不再从 config 读）
                username, password = get_credentials(config_data.get("auth"))

                page.locator("input[placeholder*='手机号'], input[placeholder*='账号']").first.fill(username)
                page.wait_for_timeout(400)
                page.locator("input[type='password'], input[placeholder*='密码']").first.fill(password)
                page.wait_for_timeout(400)
                page.locator("button:has-text('登录'), button:has-text('登 录')").first.click()
                page.wait_for_timeout(5000)
                print("✓ 登录完成\n")
                context.storage_state(path=str(AUTH_STATE_FILE))
            except Exception as e:
                print(f"登录失败: {e}")
                return
        else:
            if AUTH_STATE_FILE.exists():
                context = browser.new_context(storage_state=str(AUTH_STATE_FILE))
            else:
                context = browser.new_context()
            page = context.new_page()

        page.goto(package_url, timeout=30000)
        page.wait_for_timeout(3000)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = LOG_DIR / f"package-batch-{ts}.csv"

        with open(log_file, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["序号", "名称", "结果", "详情", "搜索校验"])

            for i, pkg in enumerate(packages, 1):
                if i < start:
                    continue
                if limit and i >= start + limit:
                    break

                page.goto(package_url)
                page.wait_for_timeout(2500)

                result, detail = create_one_package(page, pkg, package_conf)

                # 执行后验证：套餐列表页搜索
                verify_detail = ""
                if result == "OK":
                    page.goto(package_url)
                    page.wait_for_timeout(1500)
                    found, verify_detail = verify_in_list(
                        page, pkg.get("name", ""),
                        # 套餐列表页搜索框 placeholder 不一定固定，多个候选
                        search_input_selector=(
                            "input[placeholder*='套餐'], "
                            "input[placeholder*='商品'], "
                            "input[placeholder*='名称']"
                        ),
                    )
                    if not found:
                        result = "OK_BUT_NOT_FOUND"
                        print(f"  [WARN] 创建提示成功但列表搜不到：{pkg.get('name')} ({verify_detail})")

                writer.writerow([i, pkg.get("name", "Unknown"), result, detail, verify_detail])
                f.flush()

        print(f"\n批量完成: {log_file}")
        try:
            input("\n按 Enter 关闭...")
        except EOFError:
            page.wait_for_timeout(5000)
        browser.close()


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--xlsx", required=True)
    parser.add_argument("--start", type=int, default=1)
    parser.add_argument("--limit", type=int)
    parser.add_argument("--use-password", action="store_true")
    parser.add_argument("--no-precheck", action="store_true",
                        help="跳过执行前校验（不推荐）")
    args = parser.parse_args()
    batch_create(args.xlsx, args.start, args.limit, args.use_password,
                 strict_precheck=not args.no_precheck)
