"""快速诊断：用 Playwright 打开 BI 报表，列出所有 frame 的 alias / input 控件结构。"""
import time
from pathlib import Path
from playwright.sync_api import sync_playwright

BI_URL = "https://bi.61info.cn/smartbi/vision/index.jsp"
USERNAME = "68448"
PASSWORD = "12345678"
REPORT_PATH = ["海外直播业务线", "海外后端", "思维-后端", "服务", "语义分析_服务", "海外思维服务SOP执行情况"]


def main():
    with sync_playwright() as pw:
        browser = pw.chromium.launch(
            channel="msedge",
            executable_path=r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
            headless=False,
        )
        ctx = browser.new_context(accept_downloads=True)
        page = ctx.new_page()

        # 登录
        page.goto(BI_URL, wait_until="networkidle", timeout=30000)
        time.sleep(3)
        page.locator("input.item-textinput").first.fill(USERNAME)
        page.locator("input.item-textinput").nth(1).fill(PASSWORD)
        page.locator("input.item-submit").click()
        page.wait_for_load_state("networkidle")
        time.sleep(5)

        # 进入分析展现
        page.locator("span[bofid='Analysis']").first.click()
        time.sleep(3)

        # 打开报表
        for i, name in enumerate(REPORT_PATH):
            is_last = (i == len(REPORT_PATH) - 1)
            page.evaluate(f"""
                () => {{
                    for (const a of document.querySelectorAll('a')) {{
                        if (a.textContent.trim() === '{name}') {{
                            a.dispatchEvent(new MouseEvent('dblclick', {{ bubbles: true, cancelable: true }}));
                            return true;
                        }}
                    }}
                    return false;
                }}
            """)
            time.sleep(15 if is_last else 3)

        # 等额外 10s 让 iframe 加载
        time.sleep(10)

        print(f"\n=== 诊断 frame 结构 ===")
        print(f"context.pages: {len(ctx.pages)}")
        # 切到最新 page
        if len(ctx.pages) > 1:
            page = ctx.pages[-1]
            page.wait_for_load_state("networkidle")
            time.sleep(3)
            print(f"切换到最新 page")

        frames = page.frames
        print(f"page.frames count: {len(frames)}")
        for idx, f in enumerate(frames):
            try:
                url = f.url
                summary = f.evaluate("""
                    () => ({
                        title: document.title,
                        bodyChildren: document.body ? document.body.children.length : 0,
                        aliasSpanCount: document.querySelectorAll('span.aliasSpan').length,
                        aliasSpanSamples: Array.from(document.querySelectorAll('span.aliasSpan')).slice(0,15).map(s => s.textContent.trim()),
                        inputCount: document.querySelectorAll('input').length,
                        comboBtnCount: document.querySelectorAll('input.combobox-button').length,
                        iframeCount: document.querySelectorAll('iframe').length
                    })
                """)
                print(f"\n  frame[{idx}] url={url[:120]}")
                for k, v in summary.items():
                    print(f"    {k}: {v}")
            except Exception as e:
                print(f"  frame[{idx}] EXC: {e}")

        # 关闭
        time.sleep(5)
        browser.close()


if __name__ == "__main__":
    main()
