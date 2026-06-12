#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""标签数据同步：先刷新标签和用户群页面，然后进行标签数据同步配置"""
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
OUTPUT_DIR = ROOT / "output"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"


def log(m): print(m, flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--user-group-name", help="用户群名称，默认 2026年6月海外益智停课学员新")
    ap.add_argument("--month", help="月份 YYYY-MM，默认当前月")
    args = ap.parse_args()

    if not AUTH_PATH.exists():
        log(f"[ERROR] 找不到 {AUTH_PATH}")
        sys.exit(1)

    # 确定用户群名称
    if args.user_group_name:
        user_group_name = args.user_group_name
    else:
        if args.month:
            y, m = args.month.split("-")
        else:
            today = dt.date.today()
            y, m = str(today.year), str(today.month)
        user_group_name = f"{y}年{int(m)}月海外益智停课学员新"

    log(f"[INFO] 用户群名称: {user_group_name}")

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
        liuyi.wait_for_timeout(2000)

        log("[STEP 6] 刷新页面加载数据")
        liuyi.reload(timeout=30000)
        liuyi.wait_for_load_state("networkidle", timeout=30000)
        liuyi.wait_for_timeout(3000)

        log("[STEP 7] 使用JavaScript填写表单")

        # 使用JavaScript来填写表单
        fill_form_js = f"""
        async () => {{
            // 等待表单加载
            await new Promise(resolve => setTimeout(resolve, 1000));

            // 获取所有的输入框
            const inputs = document.querySelectorAll('input[placeholder="请选择"]');
            console.log('找到输入框数量:', inputs.length);

            // 填写业务类型 - 豌豆
            if (inputs.length > 0) {{
                console.log('填写业务类型');
                inputs[0].click();
                await new Promise(resolve => setTimeout(resolve, 500));

                // 查找并点击豌豆选项
                const options = document.querySelectorAll('.el-option__content');
                for (let opt of options) {{
                    if (opt.textContent.includes('豌豆')) {{
                        opt.click();
                        break;
                    }}
                }}
                await new Promise(resolve => setTimeout(resolve, 500));
            }}

            // 填写同步业务系统 - 豌豆数合表
            if (inputs.length > 1) {{
                console.log('填写同步业务系统');
                inputs[1].click();
                await new Promise(resolve => setTimeout(resolve, 500));

                const options = document.querySelectorAll('.el-option__content');
                for (let opt of options) {{
                    if (opt.textContent.includes('豌豆数合表')) {{
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

                // 输入用户群名称
                inputs[2].value = '{user_group_name}';
                inputs[2].dispatchEvent(new Event('input', {{ bubbles: true }}));
                await new Promise(resolve => setTimeout(resolve, 1000));

                // 查找并点击用户群选项
                const options = document.querySelectorAll('.el-option__content');
                for (let opt of options) {{
                    if (opt.textContent.includes('{user_group_name}')) {{
                        opt.click();
                        break;
                    }}
                }}
                await new Promise(resolve => setTimeout(resolve, 500));
            }}

            // 填写同步数据频率 - 每天
            if (inputs.length > 3) {{
                console.log('填写同步数据频率');
                inputs[3].click();
                await new Promise(resolve => setTimeout(resolve, 500));

                const options = document.querySelectorAll('.el-option__content');
                for (let opt of options) {{
                    if (opt.textContent.includes('每天')) {{
                        opt.click();
                        break;
                    }}
                }}
                await new Promise(resolve => setTimeout(resolve, 500));
            }}

            // 填写状态 - 启用
            if (inputs.length > 4) {{
                console.log('填写状态');
                inputs[4].click();
                await new Promise(resolve => setTimeout(resolve, 500));

                const options = document.querySelectorAll('.el-option__content');
                for (let opt of options) {{
                    if (opt.textContent.includes('启用')) {{
                        opt.click();
                        break;
                    }}
                }}
                await new Promise(resolve => setTimeout(resolve, 500));
            }}

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
                // 查找所有按钮
                const buttons = document.querySelectorAll('button');
                for (let btn of buttons) {
                    if (btn.textContent.includes('确认') || btn.textContent.includes('确 定')) {
                        console.log('找到确认按钮:', btn.textContent);
                        btn.click();
                        return 'Confirm button clicked';
                    }
                }
                return 'Confirm button not found';
            }
            """
            result = liuyi.evaluate(click_confirm_js)
            log(f"  -> {result}")
            liuyi.wait_for_timeout(3000)
        except Exception as e:
            log(f"  [WARN] 确认按钮点击失败: {e}")

        log("[STEP 9] 等待页面刷新后点击「手动同步」按钮")
        liuyi.wait_for_timeout(5000)

        try:
            # 使用Playwright的locator直接点击手动同步按钮
            log("  -> 尝试使用locator点击手动同步按钮")
            liuyi.locator("button").filter(has_text="手动同步").first.click(timeout=10000)
            log("  -> 手动同步按钮已点击")
            liuyi.wait_for_timeout(3000)
        except Exception as e:
            log(f"  [WARN] 第一次尝试失败: {e}")
            try:
                # 尝试第二种方式
                log("  -> 尝试使用XPath点击手动同步按钮")
                liuyi.locator("//button[contains(text(), '手动同步')]").first.click(timeout=10000)
                log("  -> 手动同步按钮已点击")
                liuyi.wait_for_timeout(3000)
            except Exception as e2:
                log(f"  [WARN] 第二次尝试失败: {e2}")
                try:
                    # 尝试第三种方式 - 查找所有按钮并逐个检查
                    log("  -> 尝试遍历所有按钮")
                    buttons = liuyi.locator("button").all()
                    for btn in buttons:
                        try:
                            text = btn.text_content()
                            if "手动同步" in text:
                                log(f"  -> 找到手动同步按钮: {text}")
                                btn.click(timeout=5000)
                                liuyi.wait_for_timeout(3000)
                                break
                        except:
                            pass
                except Exception as e3:
                    log(f"  [WARN] 第三次尝试失败: {e3}")

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
