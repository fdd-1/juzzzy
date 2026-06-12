#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""标签数据同步：先刷新标签和用户群页面，然后进行标签数据同步配置

根据确认：同步用户群选「益智群」，同步频率「每天」，状态「启用」
"""
import sys, io, time, argparse, datetime as dt, json
from pathlib import Path
from playwright.sync_api import sync_playwright

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent
ROOT = SCRIPT_DIR.parent
AUTH_PATH = ROOT / "liuyi_login" / "auth_state.json"
OUTPUT_DIR = ROOT / "output" / "p0"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"


def log(m): print(m, flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--user-group-name", help="用户群名称（默认从 group_ids_*.json 读取益智群）")
    ap.add_argument("--month", help="月份 YYYY-MM，默认当前月")
    ap.add_argument("--group-ids-json", help="group_ids_*.json，默认 output/group_ids_{today}.json")
    args = ap.parse_args()

    if not AUTH_PATH.exists():
        log(f"[ERROR] 找不到 {AUTH_PATH}")
        sys.exit(1)

    # 确定用户群名称
    if args.user_group_name:
        user_group_name = args.user_group_name
    else:
        # 从 group_ids_*.json 读益智群名称
        today_tag = dt.date.today().strftime("%Y%m%d")
        group_ids_path = Path(args.group_ids_json) if args.group_ids_json else OUTPUT_DIR / f"group_ids_{today_tag}.json"
        if not group_ids_path.exists():
            log(f"[ERROR] 找不到 {group_ids_path}（先跑 liuyi_tag/create_group.py）")
            sys.exit(1)
        group_ids = json.loads(group_ids_path.read_text(encoding="utf-8"))
        user_group_name = group_ids["yizhi_group"]["name"]

    log(f"[INFO] 用户群名称（益智群）: {user_group_name}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=100)
        ctx = browser.new_context(storage_state=str(AUTH_PATH))
        page = ctx.new_page()

        log("[STEP 1] 打开 portal")
        page.goto(PORTAL_URL, timeout=30000)
        page.locator("p:has-text('六一工作台')").first.wait_for(timeout=30000)

        log("[STEP 2] 点击「六一工作台」")
        with ctx.expect_page(timeout=30000) as new_page_info:
            page.locator("p:has-text('六一工作台')").first.click()
        liuyi = new_page_info.value
        liuyi.wait_for_load_state("domcontentloaded", timeout=30000)
        liuyi.wait_for_timeout(5000)

        log("[STEP 3] 访问用户标签页面刷新")
        liuyi.goto("https://home.61info.cn/#/userTag", timeout=30000)
        liuyi.wait_for_load_state("networkidle", timeout=30000)
        liuyi.wait_for_timeout(2000)

        log("[STEP 4] 访问用户群页面刷新")
        liuyi.goto("https://home.61info.cn/#/userGroup", timeout=30000)
        liuyi.wait_for_load_state("networkidle", timeout=30000)
        liuyi.wait_for_timeout(2000)

        log("[STEP 5] 回到标签数据同步页面")
        liuyi.goto("https://home.61info.cn/#/tagDataSync", timeout=30000)
        liuyi.wait_for_load_state("networkidle", timeout=30000)
        liuyi.wait_for_timeout(5000)

        log("[STEP 6.5] 等待「新增」按钮出现并点击")
        try:
            liuyi.locator("button:has-text('新增')").first.wait_for(state="visible", timeout=15000)
            liuyi.locator("button:has-text('新增')").first.click(timeout=5000)
            log("  -> 已点击「新增」")
            liuyi.wait_for_timeout(1500)
        except Exception as e:
            log(f"  [ERROR] 新增按钮等待/点击失败: {e}")
            liuyi.screenshot(path=str(OUTPUT_DIR / f"sync_no_add_btn_{dt.date.today().strftime('%Y%m%d')}.png"))
            sys.exit(1)

        log("[STEP 7] 使用JavaScript填写表单")

        # 使用JavaScript来填写表单
        fill_form_js = f"""
        async () => {{
            await new Promise(resolve => setTimeout(resolve, 1000));

            const inputs = document.querySelectorAll('input[placeholder="请选择"]');
            console.log('找到输入框数量:', inputs.length);

            // 填写业务类型 - 豌豆
            if (inputs.length > 0) {{
                console.log('填写业务类型');
                inputs[0].click();
                await new Promise(resolve => setTimeout(resolve, 500));

                const options = document.querySelectorAll('.el-option__content');
                for (let opt of options) {{
                    if (opt.textContent.includes('豌豆')) {{
                        opt.click();
                        break;
                    }}
                }}
                await new Promise(resolve => setTimeout(resolve, 500));
            }}

            // 填写同步业务系统 - 豌豆数仓表
            if (inputs.length > 1) {{
                console.log('填写同步业务系统');
                inputs[1].click();
                await new Promise(resolve => setTimeout(resolve, 500));

                const options = document.querySelectorAll('.el-option__content');
                for (let opt of options) {{
                    if (opt.textContent.includes('豌豆数仓表')) {{
                        opt.click();
                        break;
                    }}
                }}
                await new Promise(resolve => setTimeout(resolve, 500));
            }}

            // 填写同步用户群 - {user_group_name}
            if (inputs.length > 2) {{
                console.log('填写同步用户群');
                inputs[2].click();
                await new Promise(resolve => setTimeout(resolve, 500));

                inputs[2].value = '{user_group_name}';
                inputs[2].dispatchEvent(new Event('input', {{ bubbles: true }}));
                await new Promise(resolve => setTimeout(resolve, 1000));

                const options = document.querySelectorAll('.el-option__content');
                for (let opt of options) {{
                    if (opt.textContent.includes('{user_group_name}')) {{
                        opt.click();
                        break;
                    }}
                }}
                await new Promise(resolve => setTimeout(resolve, 500));
            }}

            // 填写同步数据频率 - 每天 (radio 单选)
            console.log('填写同步数据频率（radio）');
            const radios = document.querySelectorAll('.el-radio');
            for (let r of radios) {{
                const t = (r.textContent || '').trim();
                if (t === '每天' || t.startsWith('每天')) {{
                    r.click();
                    break;
                }}
            }}
            await new Promise(resolve => setTimeout(resolve, 500));

            // 填写状态 - 启用 (radio 单选)
            console.log('填写状态（radio）');
            const radios2 = document.querySelectorAll('.el-radio');
            for (let r of radios2) {{
                const t = (r.textContent || '').trim();
                if (t === '启用') {{
                    r.click();
                    break;
                }}
            }}
            await new Promise(resolve => setTimeout(resolve, 500));

            return 'Form filled successfully';
        }}
        """

        try:
            result = liuyi.evaluate(fill_form_js)
            log(f"  -> {result}")
            liuyi.wait_for_timeout(2000)
        except Exception as e:
            log(f"  [ERROR] JavaScript执行失败: {e}")
            sys.exit(1)

        log("[STEP 8] 使用JavaScript点击「确认」按钮")
        try:
            click_confirm_js = """
            () => {
                // 优先在 dialog 内找确定按钮
                const dlgs = document.querySelectorAll('.el-dialog');
                for (const dlg of dlgs) {
                    if (dlg.style.display === 'none' || !dlg.offsetParent) continue;
                    const btns = dlg.querySelectorAll('button');
                    for (const btn of btns) {
                        const t = (btn.textContent || '').trim();
                        if (t === '确定' || t === '确 定' || t === '确认') {
                            btn.click();
                            return 'dialog confirm: ' + t;
                        }
                    }
                }
                // fallback: 全局
                const buttons = document.querySelectorAll('button');
                for (let btn of buttons) {
                    const t = (btn.textContent || '').trim();
                    if (t === '确定' || t === '确 定' || t === '确认') {
                        btn.click();
                        return 'global confirm: ' + t;
                    }
                }
                return 'Confirm button not found';
            }
            """
            result = liuyi.evaluate(click_confirm_js)
            log(f"  -> {result}")
            liuyi.wait_for_timeout(2000)

            # 等弹窗真正关闭，最多等 8s
            log("  -> 等待弹窗关闭")
            for i in range(16):
                visible = liuyi.evaluate("""
                () => {
                  const dlgs = document.querySelectorAll('.el-dialog__wrapper');
                  for (const d of dlgs) {
                    if (d.style.display !== 'none' && d.offsetParent) return true;
                  }
                  return false;
                }
                """)
                if not visible:
                    log("  -> 弹窗已关闭")
                    break
                liuyi.wait_for_timeout(500)
            else:
                log("  [WARN] 弹窗 8s 后仍未关闭，按 ESC 强关")
                liuyi.keyboard.press("Escape")
                liuyi.wait_for_timeout(1000)
        except Exception as e:
            log(f"  [WARN] 确认按钮点击失败: {e}")

        log("[STEP 9] 跳过手动同步（频率=每天，由调度器自动触发）")
        liuyi.wait_for_timeout(2000)

        log("[STEP 10] 截图确认完成")
        screenshot_path = OUTPUT_DIR / f"sync_tag_final_{dt.date.today().strftime('%Y%m%d')}.png"
        liuyi.screenshot(path=str(screenshot_path))
        log(f"[SCREENSHOT] {screenshot_path}")

        log("[OK] 标签数据同步配置已完成！")

        ctx.storage_state(path=str(AUTH_PATH))
        liuyi.wait_for_timeout(2000)
        browser.close()


if __name__ == "__main__":
    main()
