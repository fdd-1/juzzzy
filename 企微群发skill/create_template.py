#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""六一工作台 → 企微话术模板管理 → 豌豆素质 Tab → 批量新建话术模板（UI 自动化）

输入：
  --modules-json  解析好的模块文件（parse_docx.py 输出）
  --date          模板名后缀，默认今天 MMDD
  --dry-run       只打开页面、跑到点「新建」前停下，不真提交
  --only          只建第 N 个模块（从 1 开始，多个用逗号 1,3）

依赖：
  - 同目录 browser_profile/（持久化登录态，首次跑会扫码登录）
  - 模板内容里的附件路径必须存在
"""
import sys, io, os, json, time, argparse, csv, datetime as dt
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SCRIPT_DIR = Path(__file__).resolve().parent
# 永久浏览器 profile 目录（扫码一次永久有效）
BROWSER_PROFILE = SCRIPT_DIR / "browser_profile"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

DEFAULT_SUBJECT = "豌豆益智"
DEFAULT_FUNC = "企微群发任务配置"
DEFAULT_TYPE = "公共话术"
DEFAULT_TEAM_TAB = "豌豆素质"
TEMPLATE_NAME_LIMIT = 20


def log(m): print(m, flush=True)


def make_template_name(module_name: str, date_str: str) -> str:
    """模板名 = 模块名 + MMDD，超 20 字截前 (20 - len(date)) 字 + 月日。"""
    suffix = date_str
    if len(module_name) + len(suffix) <= TEMPLATE_NAME_LIMIT:
        return module_name + suffix
    keep = TEMPLATE_NAME_LIMIT - len(suffix)
    return module_name[:keep] + suffix


def write_log_row(csv_path: Path, row: dict, fieldnames: list):
    new = not csv_path.exists()
    with csv_path.open("a", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=fieldnames)
        if new:
            w.writeheader()
        w.writerow(row)


def goto_template_page(ctx, portal_page):
    """从 portal 跳到「六一工作台」新页面，再点左侧「企微话术模板管理」菜单。"""
    log("[STEP] 进 portal")
    portal_page.goto(PORTAL_URL, timeout=30000)
    portal_page.wait_for_load_state("domcontentloaded")
    portal_page.wait_for_timeout(3000)

    # 检查是否跳到了登录页
    if "login.61info.cn" in portal_page.url:
        log("[INFO] 首次使用，需要扫码登录...")
        log("       请用钉钉扫码登录六一，扫一次后续不再需要")
        deadline = time.time() + 300
        logged_in = False
        while time.time() < deadline:
            cur = portal_page.url
            # 多重判断：URL 离开 login 域 / 出现 portal 内容 / 出现系统选择
            if "login.61info.cn" not in cur:
                logged_in = True
                break
            try:
                if (portal_page.locator("p:has-text('六一工作台')").count() > 0
                    or portal_page.locator("text=选择进入系统").count() > 0
                    or portal_page.locator("text=平台中心").count() > 0):
                    logged_in = True
                    break
            except Exception:
                pass
            portal_page.wait_for_timeout(2000)
        if not logged_in:
            log("[ERROR] 5 分钟内没登录成功")
            raise PWTimeout("登录超时")
        log("[INFO] 登录成功！后续不需要再扫码了。")
        portal_page.wait_for_timeout(2000)
        # 如果在选择系统页，点「六一工作台」
        if portal_page.locator("p:has-text('六一工作台')").count() == 0:
            portal_page.goto(PORTAL_URL, timeout=30000)
            portal_page.wait_for_load_state("domcontentloaded")
            portal_page.wait_for_timeout(2000)

    try:
        portal_page.locator("p:has-text('六一工作台')").first.wait_for(timeout=20000)
    except PWTimeout:
        shot = SCRIPT_DIR / "exports" / dt.datetime.now().strftime("%Y%m%d") / "portal_timeout.png"
        shot.parent.mkdir(parents=True, exist_ok=True)
        portal_page.screenshot(path=str(shot), full_page=True)
        log(f"[ERROR] portal 没渲染瓷砖。截图: {shot}")
        log(f"        当前 URL: {portal_page.url}")
        raise

    log("[STEP] 点开「六一工作台」瓷砖")
    with ctx.expect_page(timeout=30000) as new_page_info:
        portal_page.locator("p:has-text('六一工作台')").first.click()
    liuyi = new_page_info.value
    liuyi.wait_for_load_state("domcontentloaded", timeout=30000)
    liuyi.wait_for_timeout(3000)

    log("[STEP] 点左侧菜单「企微话术模板管理」")
    # 菜单文字 span 可能 hidden（在折叠分组里 / 屏幕外），用 evaluate 直接点 LI / A 祖先
    try:
        liuyi.evaluate(
            """() => {
                const all = Array.from(document.querySelectorAll('span, a, li'));
                const target = all.find(el => (el.textContent || '').trim() === '企微话术模板管理');
                if (!target) throw new Error('没找到「企微话术模板管理」节点');
                const clickable = target.closest('li, a, [role=menuitem]') || target;
                clickable.scrollIntoView({block: 'center'});
                clickable.click();
            }"""
        )
    except Exception as e:
        log(f"[WARN] evaluate click 失败 {e}，回退到 locator.click(force=True)")
        liuyi.locator("text=企微话术模板管理").first.click(force=True)
    liuyi.wait_for_timeout(2500)

    log("[STEP] 切换到「豌豆素质」Tab")
    try:
        page_tab = liuyi.locator(f"text={DEFAULT_TEAM_TAB}").first
        page_tab.wait_for(state="visible", timeout=15000)
        page_tab.click()
    except PWTimeout:
        # 可能菜单没生效，重试点击
        liuyi.evaluate(
            """() => {
                const all = Array.from(document.querySelectorAll('span, a, li'));
                const target = all.find(el => (el.textContent || '').trim() === '企微话术模板管理');
                if (target) (target.closest('li, a, [role=menuitem]') || target).click();
            }"""
        )
        liuyi.wait_for_timeout(3000)
        liuyi.locator(f"text={DEFAULT_TEAM_TAB}").first.click()
    liuyi.wait_for_timeout(1500)
    return liuyi


def open_create_dialog(page):
    log("  · 点「新建话术模板」")
    page.locator("button:has-text('新建话术模板')").first.click()
    # 弹窗标题
    page.locator("text=创建话术模板").first.wait_for(timeout=10000)
    page.wait_for_timeout(500)


def select_option(page, select_index: int, option_text: str):
    """点开第 N 个 .el-select，等 popper 可见，再点选项。"""
    selects = page.locator(".el-dialog__body .el-select")
    target = selects.nth(select_index)
    target.scroll_into_view_if_needed()
    target.click()
    page.wait_for_timeout(400)
    # popper 在 body 下，找到最近一个 visible dropdown
    dropdown = page.locator(".el-select-dropdown").locator(
        "xpath=self::*[not(contains(@style,'display: none'))]"
    ).last
    item = dropdown.locator(f".el-select-dropdown__item:has-text('{option_text}')").first
    item.wait_for(state="visible", timeout=10000)
    item.click()
    page.wait_for_timeout(500)


def fill_form(page, template_name: str, attachments: list, texts: list):
    log(f"  · 填模板名: {template_name}")
    name_input = page.locator(".el-dialog__body input[placeholder='输入模板名称']").first
    name_input.fill("")
    name_input.fill(template_name)

    log(f"  · 选科目: {DEFAULT_SUBJECT}")
    select_option(page, 0, DEFAULT_SUBJECT)

    log(f"  · 选可调用功能: {DEFAULT_FUNC}")
    select_option(page, 1, DEFAULT_FUNC)

    log(f"  · 选话术类型: {DEFAULT_TYPE}")
    select_option(page, 2, DEFAULT_TYPE)

    # 加文字
    for t in texts:
        log(f"  · 加文字（{len(t)} 字）")
        page.locator(".el-dialog__body button:has-text('文字')").first.click()
        page.wait_for_timeout(1500)
        # 「添加文字」子弹窗弹出，里面有 contenteditable div
        editor = page.locator(".emoji-wysiwyg-editor, [contenteditable='true']").last
        editor.wait_for(state="visible", timeout=8000)

        # 输入并验证 editor 真有内容（之前 S7/S8 翻车就是这步：保存了但 editor 是空的）
        typed_ok = False
        for attempt in range(1, 4):
            editor.click()
            page.wait_for_timeout(200)
            editor.evaluate("el => el.focus()")
            page.wait_for_timeout(200)
            # 清掉残留再输（再次 click 可能定位到别的位置）
            editor.evaluate("el => { el.innerText = ''; }")
            page.wait_for_timeout(100)
            editor.click()
            page.wait_for_timeout(200)
            page.keyboard.type(t, delay=8)
            page.wait_for_timeout(600)
            content = editor.evaluate("el => (el.innerText || el.textContent || '').trim()")
            if content:
                log(f"  · 第 {attempt} 次输入成功，editor 当前 {len(content)} 字")
                typed_ok = True
                break
            log(f"  [WARN] 第 {attempt} 次输入后 editor 为空，重试")
        if not typed_ok:
            page.keyboard.press("Escape")
            page.wait_for_timeout(800)
            raise RuntimeError("文字输入失败：editor 三次后仍为空")

        # 点子弹窗（最上层 dialog）里的「保存」按钮
        save_btn = page.locator(".el-dialog__wrapper").filter(
            has=page.locator(".emoji-wysiwyg-editor, [contenteditable='true']")
        ).last.locator("button:has-text('保存')").first
        save_btn.wait_for(state="visible", timeout=5000)
        save_btn.click()
        page.wait_for_timeout(2000)
        # 验证子弹窗已关闭（可见 dialog 数量应该回到 1）
        still_open = page.evaluate(
            """() => {
                const dlgs = Array.from(document.querySelectorAll('.el-dialog__wrapper'))
                  .filter(d => d.style.display !== 'none' && getComputedStyle(d).display !== 'none');
                return dlgs.length;
            }"""
        )
        if still_open > 1:
            log(f"  [WARN] 保存后子弹窗仍在（{still_open} 个 dialog），文字可能没保存成功")
        else:
            log(f"  · 文字保存成功")

    # 加附件：先点「+ 文件」按钮让组件初始化 input，再上传
    for att in attachments:
        att_path = att["path"]
        att_type = att["type"]
        if att_type == "image":
            btn_label = "图片"
        elif att_type == "video":
            btn_label = "视频"
        else:
            btn_label = "文件"
        log(f"  · 上传 {btn_label}: {Path(att_path).name}")
        # 点「+ 文件」按钮，用 expect_file_chooser 捕获原生文件选择器
        try:
            with page.expect_file_chooser(timeout=5000) as fc_info:
                page.locator(f".el-dialog__body button:has-text('{btn_label}')").first.click()
            fc = fc_info.value
            fc.set_files(att_path)
        except Exception:
            # 兜底：如果 file_chooser 没捕获到，直接给 input 设文件
            log(f"  · [兜底] 直接设 input[type=file]")
            file_input = page.locator(".el-dialog__body input[type='file']").last
            file_input.set_input_files(att_path)
        page.wait_for_timeout(3000)
        # 如果弹了原生对话框，按 Escape 关掉
        page.keyboard.press("Escape")
        page.wait_for_timeout(1000)


def submit_form(page) -> bool:
    log("  · 提交（点弹窗「确定」）")
    # 用 Playwright locator 真实点击 footer 里的「确 定」按钮
    try:
        ok_btn = page.locator(".el-dialog__footer button:has-text('确定'), .el-dialog__footer button:has-text('确 定')").last
        ok_btn.wait_for(state="visible", timeout=5000)
        ok_btn.click()
    except PWTimeout:
        log("  [WARN] footer 确定按钮不可见，尝试 force click")
        page.locator(".el-dialog__footer button:has-text('确')").last.click(force=True)
    # 等弹窗消失
    page.wait_for_timeout(3000)
    still_open = page.evaluate(
        """() => {
            const dlgs = Array.from(document.querySelectorAll('.el-dialog__wrapper'))
              .filter(d => d.style.display !== 'none' && getComputedStyle(d).display !== 'none');
            return dlgs.length;
        }"""
    )
    if still_open > 0:
        # 可能有 success toast 但弹窗动画没完，再等一下
        page.wait_for_timeout(3000)
        still_open = page.evaluate(
            """() => {
                const dlgs = Array.from(document.querySelectorAll('.el-dialog__wrapper'))
                  .filter(d => d.style.display !== 'none' && getComputedStyle(d).display !== 'none');
                return dlgs.length;
            }"""
        )
    return still_open == 0


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--modules-json", required=True)
    ap.add_argument("--date", default=None, help="模板名后缀，默认 MMDD")
    ap.add_argument("--dry-run", action="store_true", help="只打开页面跑到「新建」前停下")
    ap.add_argument("--only", default=None, help="只跑第 N 个模块，多个用逗号，例如 1 或 1,3")
    ap.add_argument("--no-headless", action="store_true", help="可见浏览器（默认就是可见）")
    ap.add_argument("--keep-open", action="store_true", help="跑完留住浏览器方便检查")
    args = ap.parse_args()

    modules_path = Path(args.modules_json).resolve()
    if not modules_path.exists():
        log(f"[ERROR] 找不到 {modules_path}")
        sys.exit(1)
    modules = json.loads(modules_path.read_text(encoding="utf-8"))
    if not modules:
        log("[ERROR] modules 为空")
        sys.exit(1)

    if args.only:
        idx = [int(x.strip()) for x in args.only.split(",") if x.strip()]
        modules = [m for i, m in enumerate(modules, start=1) if i in idx]
        log(f"[INFO] --only 过滤后剩 {len(modules)} 个模块")

    date_str = args.date or dt.datetime.now().strftime("%m%d")

    # 计划摘要
    log("=" * 60)
    log("[计划]")
    for i, m in enumerate(modules, start=1):
        tn = make_template_name(m["name"], date_str)
        log(f"  [{i}] {tn}  | 文字 {len(m['texts'])} 段 | 附件 {len(m['attachments'])} 个")
        for a in m["attachments"]:
            ok = "✓" if Path(a["path"]).exists() else "✗"
            log(f"      {ok} {a['type']}: {Path(a['path']).name}")
    log("=" * 60)

    if args.dry_run:
        log("[dry-run] 不打开浏览器，只打印计划。去掉 --dry-run 才会真跑。")
        return

    # 校验所有附件存在
    bad = [a["path"] for m in modules for a in m["attachments"] if not Path(a["path"]).exists()]
    if bad:
        log(f"[ERROR] 以下附件不存在:")
        for b in bad:
            log(f"  - {b}")
        sys.exit(2)

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = SCRIPT_DIR / "exports" / dt.datetime.now().strftime("%Y%m%d")
    out_dir.mkdir(parents=True, exist_ok=True)
    csv_path = out_dir / f"templates_log_{stamp}.csv"
    shot_dir = out_dir / "screenshots"
    shot_dir.mkdir(exist_ok=True)
    fields = ["index", "module", "template_name", "status", "error"]

    with sync_playwright() as p:
        # 用 persistent context：浏览器 profile 永久保存登录态，不再需要反复扫码
        BROWSER_PROFILE.mkdir(parents=True, exist_ok=True)
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(BROWSER_PROFILE),
            headless=False, channel="chrome", slow_mo=60,
        )
        portal = ctx.new_page()

        try:
            page = goto_template_page(ctx, portal)
        except Exception as e:
            log(f"[ERROR] 打开企微话术模板管理页失败: {e}")
            portal.screenshot(path=str(shot_dir / f"open_fail_{stamp}.png"))
            ctx.close()
            sys.exit(3)

        for i, m in enumerate(modules, start=1):
            tn = make_template_name(m["name"], date_str)
            log(f"\n[{i}/{len(modules)}] 建模板「{tn}」")
            try:
                open_create_dialog(page)
                fill_form(page, tn, m["attachments"], m["texts"])
                ok = submit_form(page)
                if ok:
                    log(f"  ✓ 提交成功")
                    write_log_row(csv_path, {
                        "index": i, "module": m["name"], "template_name": tn,
                        "status": "OK", "error": "",
                    }, fields)
                else:
                    page.screenshot(path=str(shot_dir / f"submit_fail_{i}_{stamp}.png"))
                    write_log_row(csv_path, {
                        "index": i, "module": m["name"], "template_name": tn,
                        "status": "FAIL", "error": "弹窗未消失",
                    }, fields)
            except Exception as e:
                log(f"  ✗ {e}")
                try:
                    page.screenshot(path=str(shot_dir / f"err_{i}_{stamp}.png"))
                except Exception:
                    pass
                write_log_row(csv_path, {
                    "index": i, "module": m["name"], "template_name": tn,
                    "status": "FAIL", "error": str(e),
                }, fields)
                # 试着关掉可能残留的弹窗
                try:
                    page.keyboard.press("Escape")
                    page.wait_for_timeout(800)
                except Exception:
                    pass
            page.wait_for_timeout(1500)

        log(f"\n[OK] 全部完成，日志: {csv_path}")
        if args.keep_open:
            log("[STAY] --keep-open，浏览器保留 5 分钟")
            page.wait_for_timeout(300_000)
        ctx.close()


if __name__ == "__main__":
    main()
