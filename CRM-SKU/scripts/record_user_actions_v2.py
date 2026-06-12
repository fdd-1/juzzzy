#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
增强版录制脚本 v2
- 记录每个事件的 outerHTML（前 300 字符）
- 记录 form-item 结构（label + 所有 input/switch 状态）
- 在 click switch 时特别记录开关状态变化
"""

import sys
import io
import json
from pathlib import Path
from datetime import datetime
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

SCRIPT_DIR = Path(__file__).parent
CONFIG_FILE = SCRIPT_DIR / "config.local.json"
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

RECORD_SCRIPT = """
() => {
    if (window.__recording_initialized) return;
    window.__recording_initialized = true;
    window.__recordings = [];

    function getElementInfo(el) {
        if (!el) return null;
        const rect = el.getBoundingClientRect ? el.getBoundingClientRect() : null;

        let label = '';
        let formItem = el.closest('.el-form-item');
        if (formItem) {
            const labelEl = formItem.querySelector('.el-form-item__label');
            if (labelEl) label = labelEl.innerText.trim();
        }

        let dialogTitle = '';
        let dialog = el.closest('.el-dialog');
        if (dialog) {
            const titleEl = dialog.querySelector('.el-dialog__title');
            if (titleEl) dialogTitle = titleEl.innerText.trim();
        }

        // 检查是否在下拉/弹层中
        let inDropdown = false;
        if (el.closest('.el-select-dropdown__item, .el-cascader-node, .el-autocomplete-suggestion__item, .el-popper li')) {
            inDropdown = true;
        }

        // outerHTML（前 200 字符）
        let outer = '';
        try {
            outer = (el.outerHTML || '').substring(0, 200);
        } catch (e) {}

        // 父级 form-item 整体结构（label + input/switch 状态）
        let formItemHtml = '';
        if (formItem) {
            try {
                const labelEl = formItem.querySelector('.el-form-item__label');
                const switches = formItem.querySelectorAll('.el-switch');
                const inputs = formItem.querySelectorAll('input');
                const parts = [`label="${labelEl ? labelEl.innerText.trim() : ''}"`];
                switches.forEach((sw, i) => {
                    const checked = sw.classList.contains('is-checked');
                    parts.push(`switch${i}=${checked ? 'ON' : 'OFF'}`);
                });
                inputs.forEach((inp, i) => {
                    const v = inp.value || '';
                    const ph = inp.getAttribute('placeholder') || '';
                    const ro = inp.hasAttribute('readonly');
                    const dis = inp.disabled;
                    if (inp.offsetParent !== null) { // visible
                        parts.push(`input${i}[ph='${ph}',ro=${ro},dis=${dis},val='${v.substring(0, 30)}']`);
                    }
                });
                formItemHtml = parts.join(' ');
            } catch (e) {}
        }

        return {
            tag: el.tagName,
            class: el.className && typeof el.className === 'string' ? el.className.substring(0, 80) : '',
            id: el.id || '',
            text: (el.innerText || '').substring(0, 60).trim(),
            value: el.value || '',
            placeholder: el.getAttribute ? (el.getAttribute('placeholder') || '') : '',
            type: el.getAttribute ? (el.getAttribute('type') || '') : '',
            label: label,
            dialog_title: dialogTitle,
            in_dropdown: inDropdown,
            outer: outer,
            form_item: formItemHtml,
            x: rect ? Math.round(rect.x) : 0,
            y: rect ? Math.round(rect.y) : 0
        };
    }

    function record(eventType, target, extra) {
        const info = getElementInfo(target);
        if (!info) return;
        const entry = {
            time: new Date().toISOString(),
            event: eventType,
            target: info,
            ...(extra || {})
        };
        window.__recordings.push(entry);
        if (window.__recordings.length > 1500) {
            window.__recordings.shift();
        }
    }

    document.addEventListener('click', (e) => {
        record('click', e.target);
    }, true);

    document.addEventListener('input', (e) => {
        record('input', e.target, { input_value: e.target.value || '' });
    }, true);

    document.addEventListener('change', (e) => {
        record('change', e.target, { change_value: e.target.value || '' });
    }, true);

    document.addEventListener('keydown', (e) => {
        if (['Enter', 'Escape', 'Tab', 'PageDown', 'PageUp', 'ArrowDown', 'ArrowUp'].includes(e.key)) {
            record('keydown', e.target, { key: e.key });
        }
    }, true);

    console.log('[Recording v2] Initialized');
}
"""

GET_RECORDINGS = "() => window.__recordings || []"


def main():
    with open(CONFIG_FILE, encoding="utf-8") as f:
        config = json.load(f)

    login_url = config["crm"]["login_url"]
    package_url = config["crm"]["package_manage_url"]

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)
        context = browser.new_context()
        page = context.new_page()

        print("登录中...")
        page.goto(login_url, timeout=30000)
        page.wait_for_timeout(2000)

        # 从 config.local.json 读取账号密码
        username = config.get("auth", {}).get("username", "")
        password = config.get("auth", {}).get("password", "")
        if not username or not password:
            raise ValueError("config.local.json 的 auth.username / auth.password 未配置")

        page.locator("input[placeholder*='手机号'], input[placeholder*='账号']").first.fill(username)
        page.wait_for_timeout(500)
        page.locator("input[type='password'], input[placeholder*='密码']").first.fill(password)
        page.wait_for_timeout(500)
        page.locator("button:has-text('登录'), button:has-text('登 录')").first.click()
        page.wait_for_timeout(5000)
        print("✓ 登录完成\n")

        page.goto(package_url, timeout=30000)
        page.wait_for_timeout(3000)

        def inject_recording():
            try:
                page.evaluate(RECORD_SCRIPT)
                print(f"[INFO] 录制脚本已注入")
            except Exception as e:
                print(f"[WARN] 注入失败: {e}")

        page.on("framenavigated", lambda frame: inject_recording() if frame == page.main_frame else None)
        inject_recording()

        print("=" * 80)
        print("【录制模式 v2】")
        print("=" * 80)
        print("请在浏览器中手动创建套餐。")
        print("重点操作步骤 2：是否换课、转介绍规则、重复购买次数")
        print("操作完成后告诉我「录制完成」")
        print("=" * 80)

        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = LOG_DIR / f"recording-v2-{timestamp}.json"

        last_count = 0
        try:
            while True:
                page.wait_for_timeout(3000)
                try:
                    recs = page.evaluate(GET_RECORDINGS)
                    if recs and len(recs) != last_count:
                        with open(output_file, "w", encoding="utf-8") as f:
                            json.dump(recs, f, ensure_ascii=False, indent=2)
                        new_count = len(recs) - last_count
                        last_count = len(recs)
                        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 总 {len(recs)} (+{new_count})")
                        # 打印最新事件简要
                        for r in recs[-new_count:][-5:]:
                            ev = r.get('event', '')
                            t = r.get('target', {})
                            label = t.get('label', '')
                            text = t.get('text', '')[:30]
                            ph = t.get('placeholder', '')
                            tag = t.get('tag', '')
                            fi = t.get('form_item', '')[:80]
                            extra = ''
                            if 'input_value' in r:
                                extra = f" v='{r['input_value'][:20]}'"
                            elif 'key' in r:
                                extra = f" key={r['key']}"
                            label_info = label or ph or ''
                            print(f"  {ev:6s} {tag:6s} [{label_info:15s}] '{text}'{extra}")
                            if fi and 'switch' in fi:
                                print(f"        > {fi}")
                except Exception as e:
                    inject_recording()
        except KeyboardInterrupt:
            print("\n[INFO] 录制结束")
            try:
                recs = page.evaluate(GET_RECORDINGS)
                with open(output_file, "w", encoding="utf-8") as f:
                    json.dump(recs, f, ensure_ascii=False, indent=2)
                print(f"✓ 日志已保存: {output_file} 共 {len(recs)} 条")
            except:
                pass
        browser.close()


if __name__ == "__main__":
    main()
