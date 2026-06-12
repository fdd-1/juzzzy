#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""标签数据同步（修复版）

正确流程：
1. 进入标签数据同步页面
2. 点击右上角「新增」按钮
3. 弹出表单后填写：
   - 业务类型 = 豌豆
   - 同步业务系统 = 豌豆数仓
   - 同步用户群 = 用户群名称
   - 同步数据频率 = 每天 (radio)
   - 状态 = 启用 (radio)
4. 点击「确认」按钮
"""
import sys, io, time, argparse, datetime as dt
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent
# 优先用六一标签的 auth_state.json（最新登录），兜底找停课唤醒目标的
AUTH_CANDIDATES = [
    Path(r"C:\Users\fengjianyi\Desktop\六一标签\auth_state.json"),
    Path(r"C:\Users\fengjianyi\Desktop\停课唤醒目标\liuyi_login\auth_state.json"),
]
OUTPUT_DIR = SCRIPT_DIR / "logs"
OUTPUT_DIR.mkdir(exist_ok=True)
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"


def log(m): print(m, flush=True)


def find_auth():
    for p in AUTH_CANDIDATES:
        if p.exists():
            return p
    return None


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--user-group-name", required=True, help="用户群名称")
    ap.add_argument("--biz-type", default="豌豆", help="业务类型（默认豌豆）")
    ap.add_argument("--biz-system", default="豌豆数仓", help="同步业务系统（默认豌豆数仓）")
    ap.add_argument("--frequency", default="每天", help="同步数据频率（默认每天）")
    ap.add_argument("--status", default="启用", help="状态（默认启用）")
    args = ap.parse_args()

    auth_path = find_auth()
    if not auth_path:
        log(f"[ERROR] 找不到登录态文件。请先登录六一工作台")
        log(f"  python C:\\Users\\fengjianyi\\Desktop\\六一标签\\login.py")
        sys.exit(1)
    log(f"[INFO] 登录态: {auth_path}")
    log(f"[INFO] 用户群名称: {args.user_group_name}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=150)
        ctx = browser.new_context(storage_state=str(auth_path))

        log("[STEP 1] 打开 portal 进入六一工作台")
        page = ctx.new_page()
        page.goto("https://dingding.61info.cn/sys/portal/page.jsp", timeout=30000)
        page.locator("p:has-text('六一工作台')").first.wait_for(timeout=30000)
        with ctx.expect_page(timeout=30000) as new_page_info:
            page.locator("p:has-text('六一工作台')").first.click()
        liuyi = new_page_info.value
        liuyi.wait_for_load_state("domcontentloaded", timeout=30000)
        liuyi.wait_for_timeout(3000)

        log("[STEP 2] 导航到标签数据同步页面")
        liuyi.goto("https://home.61info.cn/#/tagDataSync", timeout=30000)
        liuyi.wait_for_timeout(5000)

        # 截图: 进入页面后的状态
        liuyi.screenshot(path=str(OUTPUT_DIR / f"01_after_load_{dt.datetime.now().strftime('%H%M%S')}.png"))

        log("[STEP 5] 点击「新增」按钮（页面右上角）")
        click_add_js = """
        () => {
            const buttons = document.querySelectorAll('button');
            const found = [];
            for (let btn of buttons) {
                const text = (btn.textContent || '').trim();
                found.push(text);
                if (text === '新增' || text === '+ 新增' || text.replace(/\\s+/g, '') === '新增') {
                    btn.click();
                    return {ok: true, text: text};
                }
            }
            return {ok: false, buttons: found};
        }
        """
        result = liuyi.evaluate(click_add_js)
        if not result.get("ok"):
            log(f"  [ERROR] 未找到「新增」按钮")
            log(f"  页面上的按钮: {result.get('buttons')}")
            liuyi.screenshot(path=str(OUTPUT_DIR / f"err_no_add_btn_{dt.datetime.now().strftime('%H%M%S')}.png"))
            sys.exit(1)
        log(f"  -> 已点击「{result.get('text')}」按钮")
        liuyi.wait_for_timeout(2000)

        # 截图: 弹出表单后
        liuyi.screenshot(path=str(OUTPUT_DIR / f"02_dialog_open_{dt.datetime.now().strftime('%H%M%S')}.png"))

        log("[STEP 6] 填写表单")

        # 6.1 业务类型 = 豌豆
        log(f"  [6.1] 业务类型 = {args.biz_type}")
        select_dropdown_js = """
        ({label, value}) => {
            // 在弹窗内查找指定 label 的下拉框
            const dialog = document.querySelector('.el-dialog__wrapper:not([style*="display: none"]) .el-dialog')
                        || document.querySelector('.el-dialog');
            if (!dialog) return {ok: false, err: 'dialog not found'};

            const formItems = dialog.querySelectorAll('.el-form-item');
            for (let item of formItems) {
                const lbl = item.querySelector('.el-form-item__label');
                if (lbl && lbl.textContent.includes(label)) {
                    const input = item.querySelector('input.el-input__inner');
                    if (input) {
                        input.click();
                        input.dispatchEvent(new MouseEvent('mousedown', {bubbles: true}));
                        return {ok: true, opened: true};
                    }
                }
            }
            return {ok: false, err: `label ${label} not found`};
        }
        """
        click_option_js = """
        (value) => {
            // 找当前可见的下拉面板里的选项
            const dropdowns = document.querySelectorAll('.el-select-dropdown:not([style*="display: none"])');
            for (let dd of dropdowns) {
                const items = dd.querySelectorAll('.el-select-dropdown__item');
                for (let it of items) {
                    if (it.textContent.trim() === value || it.textContent.includes(value)) {
                        it.click();
                        return {ok: true, text: it.textContent.trim()};
                    }
                }
            }
            return {ok: false};
        }
        """

        # 业务类型
        r = liuyi.evaluate(select_dropdown_js, {"label": "业务类型", "value": args.biz_type})
        log(f"    open dropdown: {r}")
        liuyi.wait_for_timeout(800)
        r = liuyi.evaluate(click_option_js, args.biz_type)
        log(f"    click option: {r}")
        liuyi.wait_for_timeout(800)

        # 同步业务系统
        log(f"  [6.2] 同步业务系统 = {args.biz_system}")
        r = liuyi.evaluate(select_dropdown_js, {"label": "同步业务系统", "value": args.biz_system})
        log(f"    open dropdown: {r}")
        liuyi.wait_for_timeout(800)
        r = liuyi.evaluate(click_option_js, args.biz_system)
        log(f"    click option: {r}")
        liuyi.wait_for_timeout(800)

        # 同步用户群（可输入搜索）
        log(f"  [6.3] 同步用户群 = {args.user_group_name}")
        fill_user_group_js = """
        (value) => {
            const dialog = document.querySelector('.el-dialog__wrapper:not([style*="display: none"]) .el-dialog')
                        || document.querySelector('.el-dialog');
            if (!dialog) return {ok: false, err: 'dialog not found'};

            const formItems = dialog.querySelectorAll('.el-form-item');
            for (let item of formItems) {
                const lbl = item.querySelector('.el-form-item__label');
                if (lbl && lbl.textContent.includes('同步用户群')) {
                    const input = item.querySelector('input.el-input__inner');
                    if (input) {
                        input.click();
                        // 使用 native setter 触发 React/Vue 的 value 更新
                        const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value').set;
                        setter.call(input, value);
                        input.dispatchEvent(new Event('input', {bubbles: true}));
                        input.dispatchEvent(new Event('change', {bubbles: true}));
                        return {ok: true};
                    }
                }
            }
            return {ok: false};
        }
        """
        r = liuyi.evaluate(fill_user_group_js, args.user_group_name)
        log(f"    fill input: {r}")
        liuyi.wait_for_timeout(2000)  # 等待搜索结果

        # 点击搜索结果中匹配的选项
        r = liuyi.evaluate(click_option_js, args.user_group_name)
        log(f"    click option: {r}")
        if not r.get("ok"):
            log(f"    [WARN] 搜索结果中未找到完全匹配，尝试模糊匹配...")
            # 简化：再点击下拉里第一项
            click_first_option_js = """
            () => {
                const dropdowns = document.querySelectorAll('.el-select-dropdown:not([style*="display: none"])');
                for (let dd of dropdowns) {
                    const items = dd.querySelectorAll('.el-select-dropdown__item:not(.is-disabled)');
                    if (items.length > 0) {
                        items[0].click();
                        return {ok: true, text: items[0].textContent.trim()};
                    }
                }
                return {ok: false};
            }
            """
            r = liuyi.evaluate(click_first_option_js)
            log(f"    fallback first option: {r}")
        liuyi.wait_for_timeout(800)

        # 同步数据频率（radio button）
        log(f"  [6.4] 同步数据频率 = {args.frequency}")
        click_radio_js = """
        ({label, value}) => {
            const dialog = document.querySelector('.el-dialog__wrapper:not([style*="display: none"]) .el-dialog')
                        || document.querySelector('.el-dialog');
            if (!dialog) return {ok: false, err: 'dialog not found'};

            const formItems = dialog.querySelectorAll('.el-form-item');
            for (let item of formItems) {
                const lbl = item.querySelector('.el-form-item__label');
                if (lbl && lbl.textContent.includes(label)) {
                    const radios = item.querySelectorAll('.el-radio');
                    const found = [];
                    for (let radio of radios) {
                        const txt = radio.textContent.trim();
                        found.push(txt);
                        if (txt === value || txt.includes(value)) {
                            // 点击 radio 的标签 (不是 input)
                            radio.click();
                            return {ok: true, text: txt};
                        }
                    }
                    return {ok: false, options: found};
                }
            }
            return {ok: false, err: `label ${label} not found`};
        }
        """
        r = liuyi.evaluate(click_radio_js, {"label": "同步数据频率", "value": args.frequency})
        log(f"    click radio: {r}")
        liuyi.wait_for_timeout(500)

        # 状态（radio button）
        log(f"  [6.5] 状态 = {args.status}")
        r = liuyi.evaluate(click_radio_js, {"label": "状态", "value": args.status})
        log(f"    click radio: {r}")
        liuyi.wait_for_timeout(500)

        # 截图: 表单填写完成
        liuyi.screenshot(path=str(OUTPUT_DIR / f"03_form_filled_{dt.datetime.now().strftime('%H%M%S')}.png"))

        log("[STEP 7] 点击「确认」按钮")
        click_confirm_js = """
        () => {
            // 找弹窗里的「确认」按钮
            const dialog = document.querySelector('.el-dialog__wrapper:not([style*="display: none"]) .el-dialog')
                        || document.querySelector('.el-dialog');
            if (dialog) {
                const buttons = dialog.querySelectorAll('button');
                for (let btn of buttons) {
                    const text = (btn.textContent || '').trim();
                    if (text === '确认' || text === '确 认' || text === '确定' || text === '确 定') {
                        btn.click();
                        return {ok: true, text: text};
                    }
                }
            }
            // fallback: 全局找
            const buttons = document.querySelectorAll('button');
            for (let btn of buttons) {
                const text = (btn.textContent || '').trim();
                if (text === '确认' || text === '确 认') {
                    btn.click();
                    return {ok: true, text: text, fallback: true};
                }
            }
            return {ok: false};
        }
        """
        r = liuyi.evaluate(click_confirm_js)
        log(f"  -> {r}")
        liuyi.wait_for_timeout(3000)

        # 截图: 提交后
        liuyi.screenshot(path=str(OUTPUT_DIR / f"04_after_confirm_{dt.datetime.now().strftime('%H%M%S')}.png"))

        log("[OK] 标签数据同步配置已完成")
        log(f"  截图保存在: {OUTPUT_DIR}")

        # 保存最新登录态
        try:
            ctx.storage_state(path=str(auth_path))
        except Exception:
            pass

        liuyi.wait_for_timeout(3000)
        browser.close()


if __name__ == "__main__":
    main()
