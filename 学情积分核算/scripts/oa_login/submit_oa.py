#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""自动提交 OA 豌豆币添加申请。

策略：
  1. 复用 auth_state.json 进入 portal
  2. 尝试多种方式点击「豌豆币添加申请」入口（含 iframe / scroll / force）
  3. 任一成功路径触发后等待表单出现；都失败则提示手动点击
  4. 表单出现后：填写「需要添加的豌豆币/魔力币数总共」+ 上传附件
  5. 暂停 8 秒供用户肉眼复核，然后点击「流程处理 / 提交」

附件路径解析顺序（可被 --attachment 覆盖）：
  1. 命令行 --attachment <path>
  2. 命令行 --output-dir <dir>，在该目录下找最新的「发放豌豆币文档填写模板*.xlsx」
  3. 默认：03_output/ 下最新的「学情积分发放明细」目录里的最新模板文件
"""
import argparse
import io
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent
AUTH_PATH = SCRIPT_DIR / "auth_state.json"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

BASE_DIR = Path(r"c:\Users\fengjianyi\Desktop\学情积分核算")
OUTPUT_ROOT = BASE_DIR / "03_output"
ENTRY_TEXT = "豌豆币添加申请"


def find_latest_period_dir() -> Path | None:
    """找 03_output/ 下最新的「学情积分发放明细」期次目录。"""
    if not OUTPUT_ROOT.exists():
        return None
    folders = [f for f in OUTPUT_ROOT.iterdir() if f.is_dir() and "学情积分发放明细" in f.name]
    if not folders:
        return None
    return max(folders, key=lambda p: p.stat().st_mtime)


def find_latest_template(period_dir: Path) -> Path | None:
    """在期次目录里找最新的「发放豌豆币文档填写模板*.xlsx」。
    优先带日期后缀的，没有就用裸名。"""
    if not period_dir or not period_dir.exists():
        return None
    candidates = sorted(
        period_dir.glob("发放豌豆币文档填写模板*.xlsx"),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return candidates[0] if candidates else None


def resolve_attachment(args) -> Path:
    if args.attachment:
        p = Path(args.attachment)
        if not p.exists():
            print(f"[ERROR] --attachment 不存在: {p}", flush=True)
            sys.exit(1)
        return p
    if args.output_dir:
        period_dir = Path(args.output_dir)
    else:
        period_dir = find_latest_period_dir()
        if not period_dir:
            print(f"[ERROR] 在 {OUTPUT_ROOT} 下找不到「学情积分发放明细」期次目录", flush=True)
            sys.exit(1)
    template = find_latest_template(period_dir)
    if not template:
        print(f"[ERROR] 在 {period_dir} 下找不到「发放豌豆币文档填写模板*.xlsx」", flush=True)
        sys.exit(1)
    print(f"[INFO] 期次目录: {period_dir}", flush=True)
    print(f"[INFO] 附件: {template.name}", flush=True)
    return template


_arg_parser = argparse.ArgumentParser(add_help=False)
_arg_parser.add_argument("--output-dir", help="期次目录，会在其中找最新的发放模板")
_arg_parser.add_argument("--attachment", help="直接指定附件路径，覆盖自动探测")
_args, _ = _arg_parser.parse_known_args()

ATTACHMENT = resolve_attachment(_args)


# 从附件 Excel 读取总积分数
def get_total_amount():
    """从发放模板 Excel 的第3列（添加的豌豆币数量）求和"""
    import openpyxl
    wb = openpyxl.load_workbook(ATTACHMENT, read_only=True)
    ws = wb.active
    total = 0
    for row in ws.iter_rows(min_row=2, values_only=True):  # 跳过表头
        if row and len(row) > 2 and row[2]:  # 第3列是数量
            try:
                total += int(row[2])
            except (ValueError, TypeError):
                pass
    return total

TOTAL_AMOUNT = get_total_amount()
print(f"[INFO] 从附件读取总积分数: {TOTAL_AMOUNT}", flush=True)

# 表单字段值（按表单从上到下顺序填；金额、附件单独处理）
FIELD_VALUES = {
    "申请类型": "豌豆商城豌豆币添加",
    "申请原因": "海外益智学情积分（预习课后习题消课）",
    "科目": "益智",
    "是否合同赠送": "否",
    "虚拟币类型": "豌豆币",
    "活动批准审批单号": "无",
    "积分成本归属部门": "欢乐童年_海外直播业务线_海外业务中心_海外业务运营处_海外教学服务运营组",
    "用户ID": "详情见附件",
}

if not AUTH_PATH.exists():
    print(f"[ERROR] 找不到 {AUTH_PATH}，请先完成登录。", flush=True)
    sys.exit(1)

if not ATTACHMENT.exists():
    print(f"[ERROR] 找不到附件: {ATTACHMENT}", flush=True)
    sys.exit(1)


def shot(page, name):
    p = SCRIPT_DIR / name
    try:
        page.screenshot(path=str(p), full_page=True)
        print(f"  [SHOT] {p.name}", flush=True)
    except Exception as e:
        print(f"  [SHOT-FAIL] {name}: {e}", flush=True)


def try_click_entry(page):
    """多策略尝试点击「豌豆币添加申请」入口。

    返回 (success: bool, where: str)
    """
    # 1) 主页面 — 精确文本
    candidates = [
        ("page exact-text", lambda: page.get_by_text(ENTRY_TEXT, exact=True).first),
        ("page partial-text", lambda: page.get_by_text(ENTRY_TEXT).first),
        ("page link role", lambda: page.get_by_role("link", name=ENTRY_TEXT).first),
        ("page button role", lambda: page.get_by_role("button", name=ENTRY_TEXT).first),
    ]
    for label, getter in candidates:
        try:
            loc = getter()
            if loc.count() == 0:
                continue
            print(f"  [TRY] {label} count={loc.count()}", flush=True)
            try:
                loc.scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass
            try:
                loc.click(timeout=3000)
                print(f"  [OK] clicked via {label}", flush=True)
                return True, label
            except Exception as e:
                print(f"  [SKIP] {label} normal click failed: {e}", flush=True)
                # 强制点击
                try:
                    loc.click(timeout=3000, force=True)
                    print(f"  [OK] clicked via {label} (force)", flush=True)
                    return True, f"{label}+force"
                except Exception as e2:
                    print(f"  [SKIP] {label} force click failed: {e2}", flush=True)
        except Exception as e:
            print(f"  [SKIP] {label} setup error: {e}", flush=True)

    # 2) 遍历 iframe
    for i, frame in enumerate(page.frames):
        if frame == page.main_frame:
            continue
        try:
            loc = frame.get_by_text(ENTRY_TEXT, exact=True).first
            if loc.count() > 0:
                print(f"  [TRY] iframe[{i}] url={frame.url[:60]} count={loc.count()}", flush=True)
                try:
                    loc.scroll_into_view_if_needed(timeout=3000)
                except Exception:
                    pass
                try:
                    loc.click(timeout=3000)
                    return True, f"iframe[{i}]"
                except Exception:
                    try:
                        loc.click(timeout=3000, force=True)
                        return True, f"iframe[{i}]+force"
                    except Exception as e:
                        print(f"  [SKIP] iframe[{i}] click failed: {e}", flush=True)
        except Exception:
            continue

    # 3) DOM 搜索 + JS 点击
    try:
        clicked = page.evaluate(
            """(text) => {
                const all = Array.from(document.querySelectorAll('a, span, div, p, li, button'));
                const target = all.find(el => el.textContent && el.textContent.trim() === text);
                if (target) { target.click(); return true; }
                return false;
            }""",
            ENTRY_TEXT,
        )
        if clicked:
            return True, "js evaluate"
    except Exception as e:
        print(f"  [SKIP] js evaluate failed: {e}", flush=True)

    return False, "none"


def wait_for_form(page, ctx, timeout_seconds=300):
    """等待表单出现：可能在新 tab，也可能是当前页跳转 / iframe 注入。

    判定规则：页面上能找到「需要添加」+「豌豆币」相关字样；或检测到新 page。
    """
    deadline = time.time() + timeout_seconds
    start_pages = set(p.url for p in ctx.pages)
    print(f"  [WAIT] 等待表单出现，最多 {timeout_seconds}s", flush=True)
    last_url = None
    while time.time() < deadline:
        # 检查新 tab
        for p in ctx.pages:
            if p.url not in start_pages:
                try:
                    p.wait_for_load_state("domcontentloaded", timeout=3000)
                except Exception:
                    pass
                txt_count = p.locator(":text('需要添加'), :text('豌豆币添加申请')").count()
                if txt_count > 0:
                    print(f"  [HIT] 新 tab 含表单: {p.url[:80]}", flush=True)
                    return p
        # 当前 page 是否已变成表单页
        try:
            cur = page.url
            if cur != last_url:
                print(f"  [URL] {cur[:80]}", flush=True)
                last_url = cur
            txt_count = page.locator(":text('需要添加')").count()
            if txt_count > 0:
                print(f"  [HIT] 当前 page 含表单字段", flush=True)
                return page
            for fr in page.frames:
                if fr == page.main_frame:
                    continue
                try:
                    if fr.locator(":text('需要添加')").count() > 0:
                        print(f"  [HIT] iframe 含表单: {fr.url[:80]}", flush=True)
                        return page
                except Exception:
                    pass
        except Exception as e:
            print(f"  [POLL-ERR] {e}", flush=True)
        time.sleep(1.5)
    return None


def find_in_page_or_frames(page, locator_fn):
    """在主 page 与所有 iframe 中尝试同一个 locator 工厂，返回第一个 count>0 的 locator。"""
    try:
        loc = locator_fn(page)
        if loc.count() > 0:
            return loc, page
    except Exception:
        pass
    for fr in page.frames:
        if fr == page.main_frame:
            continue
        try:
            loc = locator_fn(fr)
            if loc.count() > 0:
                return loc, fr
        except Exception:
            continue
    return None, None


with sync_playwright() as p:
    browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=120)
    # 先尝试用已保存的登录态
    if AUTH_PATH.exists():
        ctx = browser.new_context(storage_state=str(AUTH_PATH))
    else:
        ctx = browser.new_context()
    page = ctx.new_page()

    print(f"[STEP 1] 打开 portal: {PORTAL_URL}", flush=True)
    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    page.wait_for_timeout(3000)
    print(f"  [URL] {page.url}", flush=True)

    # 截图确认当前状态
    shot(page, "submit_step0_state.png")

    # 判断是否已登录：检测页面内容而非 URL
    is_logged_in = (
        page.locator("text=选择进入系统").count() > 0
        or page.locator("text=OA系统").count() > 0
        or page.locator("text=CMS系统").count() > 0
    )

    # 如果未登录，等待用户扫码（在当前浏览器会话中）
    if not is_logged_in:
        print("=" * 60, flush=True)
        print("[ACTION] 未登录，请在右侧用钉钉扫码登录（或左侧账号密码 + 短信）", flush=True)
        print("[POLL]   登录成功后脚本会自动检测并继续，最多等 10 分钟", flush=True)
        print("=" * 60, flush=True)

        deadline = time.time() + 600
        while time.time() < deadline:
            try:
                has_system_select = page.locator("text=选择进入系统").count() > 0
                has_oa_tile = page.locator("text=OA系统").count() > 0
                has_portal_content = page.locator("text=CMS系统").count() > 0
                if has_system_select or has_oa_tile or has_portal_content:
                    print(f"[LOGIN OK] 检测到登录成功", flush=True)
                    is_logged_in = True
                    # 保存登录态供下次使用
                    ctx.storage_state(path=str(AUTH_PATH))
                    print(f"[SAVE] 登录态已保存到: {AUTH_PATH}", flush=True)
                    break
            except Exception:
                pass
            page.wait_for_timeout(2000)

        if not is_logged_in:
            print("[ERROR] 10 分钟内未检测到登录", flush=True)
            browser.close()
            sys.exit(1)

    # 进入 OA
    if page.locator("text=OA系统").count() > 0:
        print("[STEP 2] 点击 OA系统 瓷砖", flush=True)
        try:
            with ctx.expect_page(timeout=15000) as new_page_info:
                page.locator("text=OA系统").first.click()
            oa_page = new_page_info.value
        except PWTimeout:
            oa_page = page
        oa_page.wait_for_load_state("domcontentloaded", timeout=30000)
        oa_page.wait_for_timeout(3000)
    else:
        print("[STEP 2] 未见系统选择页，认为已在 OA 内", flush=True)
        oa_page = page

    print(f"  [OA URL] {oa_page.url}", flush=True)
    shot(oa_page, "submit_step1_home.png")

    # 等待 portal 页面的 iframe 充分加载（portal 有很多异步加载的 iframe）
    print("[STEP 2.5] 等待 portal iframe 加载完成（10秒）", flush=True)
    oa_page.wait_for_timeout(10000)

    print(f"[STEP 3] 尝试自动点击「{ENTRY_TEXT}」入口", flush=True)
    ok, where = try_click_entry(oa_page)
    if ok:
        print(f"  [AUTO-CLICK OK] via {where}", flush=True)
    else:
        print("  [AUTO-CLICK FAIL] 等你手动点击「豌豆币添加申请」（最多 5 分钟）...", flush=True)

    print("[STEP 4] 等待表单出现", flush=True)
    form_page = wait_for_form(oa_page, ctx, timeout_seconds=300)
    if form_page is None:
        print("[ERROR] 等不到表单页面", flush=True)
        shot(oa_page, "submit_err_no_form.png")
        browser.close()
        sys.exit(1)

    form_page.wait_for_timeout(5000)
    shot(form_page, "submit_step2_form.png")

    # 调试：在所有 frame 中搜索含"积分成本"或"归属"的元素
    print("[DEBUG] 搜索含「积分成本」的元素:", flush=True)
    for i, fr in enumerate(form_page.frames):
        try:
            info = fr.evaluate("""() => {
                const all = Array.from(document.querySelectorAll('*'));
                const matches = all.filter(el => {
                    const t = (el.textContent || '').trim();
                    return t.includes('积分成本') || t.includes('归属部门');
                }).filter(el => el.children.length < 3);  // 叶子节点
                return matches.slice(0, 5).map(el => ({
                    tag: el.tagName,
                    text: el.textContent.trim().substring(0, 60),
                    cls: el.className,
                    id: el.id
                }));
            }""")
            if info and len(info) > 0:
                print(f"  [FRAME {i}] url={fr.url[:60]}", flush=True)
                for item in info:
                    print(f"    {item}", flush=True)
        except Exception:
            pass

    # 调试：搜索"点击上传"按钮所在的 frame
    print("[DEBUG] 搜索「点击上传」按钮:", flush=True)
    for i, fr in enumerate(form_page.frames):
        try:
            info = fr.evaluate("""() => {
                const all = Array.from(document.querySelectorAll('*'));
                const matches = all.filter(el => {
                    const t = (el.textContent || '').trim();
                    return t === '点击上传';
                });
                return matches.map(el => ({
                    tag: el.tagName,
                    id: el.id,
                    cls: el.className,
                    visible: el.offsetParent !== null,
                    display: getComputedStyle(el).display,
                    parent_visible: el.parentElement ? el.parentElement.offsetParent !== null : false
                }));
            }""")
            if info and len(info) > 0:
                print(f"  [FRAME {i}] url={fr.url[:60]}", flush=True)
                for item in info:
                    print(f"    {item}", flush=True)
        except Exception:
            pass

    # ──────────────────────────────────────────────────────────────────────
    # STEP 4.5 — 通用字段填写（按 label 文本就近找输入框 / radio / dropdown）
    # ──────────────────────────────────────────────────────────────────────
    print("[STEP 4.5] 填写其他必填字段", flush=True)
    fill_js = r"""
    (args) => {
        const [label, value] = args;
        // 在所有 scope 内查找含此 label 的元素（td/label/span），定位其字段容器，再写值
        function visible(el) { return el && el.offsetParent !== null; }
        function txt(el) { return ((el && el.textContent) || '').trim(); }

        // 1) 先把 label 文本节点找出来
        const allTextNodes = Array.from(document.querySelectorAll('td, label, span, div'))
            .filter(el => visible(el) && txt(el) === label);
        if (allTextNodes.length === 0) {
            // 模糊匹配：label 中包含 "*" 或前后有空格
            const fuzzy = Array.from(document.querySelectorAll('td, label, span, div'))
                .filter(el => visible(el) && txt(el).replace(/[\s*：:]/g, '') === label.replace(/[\s*：:]/g, ''));
            if (fuzzy.length === 0) return {ok: false, reason: 'label-not-found'};
            allTextNodes.push(...fuzzy);
        }

        // 2) 对每个 label，找它后面的字段容器
        // 关键：表单是双列布局（一个 <tr> 里有 4 个 td：label1, value1, label2, value2）
        // 必须只在"该 label 所在 td 的 nextElementSibling"里找控件，
        // 千万不能扩散到整个 <tr>，否则会拿到隔壁字段的输入框 / radio。
        for (const lab of allTextNodes) {
            const containers = [];
            // 首选：label td 的下一个 td（即该字段自己的值单元格）
            const labelTd = lab.closest('td');
            if (labelTd && labelTd.nextElementSibling) {
                containers.push(labelTd.nextElementSibling);
            }
            // 备用：lab 自身的 .lui-formfield-value / .field-value / td.fieldValue
            const v = (labelTd || lab).querySelector
                ? (labelTd || lab).querySelector('.lui-formfield-value, .field-value, td.fieldValue')
                : null;
            if (v) containers.push(v);
            // 兜底：lab 的直接父元素（非 TR / 非 TBODY，避免扩散到隔壁列）
            const par = lab.parentElement;
            if (par && par.tagName !== 'TR' && par.tagName !== 'TBODY' && par.tagName !== 'TABLE') {
                containers.push(par);
            }

            for (const c of containers) {
                if (!c) continue;
                // a) radio: 在容器里找 label 文本 == value 的 radio 按钮
                const radios = Array.from(c.querySelectorAll('input[type=radio]'));
                if (radios.length > 0) {
                    for (const r of radios) {
                        // 找它的可读 label：parentElement 的文本
                        const labelText = txt(r.parentElement) || r.value || '';
                        if (labelText === value || labelText.includes(value)) {
                            r.click();
                            return {ok: true, kind: 'radio', via: labelText};
                        }
                    }
                    // 也找 lui-radio 风格（div + 文本）
                    const luiRadios = Array.from(c.querySelectorAll('.lui-radio, .lui_radio, label'));
                    for (const lr of luiRadios) {
                        if (txt(lr) === value || txt(lr).includes(value)) {
                            lr.click();
                            return {ok: true, kind: 'lui-radio', via: txt(lr)};
                        }
                    }
                }
                // b) checkbox 同理（«是否含同课包» 实际可能是 radio）
                // c) select 下拉
                const sel = c.querySelector('select');
                if (sel) {
                    const opts = Array.from(sel.options);
                    const m = opts.find(o => o.text === value || o.text.includes(value) || o.value === value);
                    if (m) {
                        sel.value = m.value;
                        sel.dispatchEvent(new Event('change', {bubbles: true}));
                        return {ok: true, kind: 'select', via: m.text};
                    }
                }
                // d) 文本输入 / textarea
                const inp = c.querySelector('input[type=text], input:not([type]), textarea, input[type=number]');
                if (inp) {
                    // 跳过隐藏 / readonly
                    if (!visible(inp)) continue;
                    inp.focus();
                    // 尝试原生赋值 + 触发 input/change
                    const proto = inp.tagName === 'TEXTAREA' ? HTMLTextAreaElement.prototype : HTMLInputElement.prototype;
                    const setter = Object.getOwnPropertyDescriptor(proto, 'value').set;
                    setter.call(inp, value);
                    inp.dispatchEvent(new Event('input', {bubbles: true}));
                    inp.dispatchEvent(new Event('change', {bubbles: true}));
                    inp.blur();
                    return {ok: true, kind: 'input', via: inp.name || inp.id || ''};
                }
            }
        }
        return {ok: false, reason: 'no-suitable-control'};
    }
    """

    def fill_field(label, value):
        """在 form_page 主页和所有 iframe 里尝试填一个字段"""
        scopes = [form_page] + [fr for fr in form_page.frames if fr != form_page.main_frame]
        for sc in scopes:
            try:
                res = sc.evaluate(fill_js, [label, value])
                if res and res.get("ok"):
                    return True, res
            except Exception as e:
                last_err = str(e)
                continue
        return False, None

    def fill_inputselect_field(label, value, search=None, retries=2):
        """K2 inputselectsgl autocomplete 控件专用：输入文字 → 等候选 → 点 highlighted LI。

        基于真实录制：用户先点 .inputselectsgl 占位 DIV，控件展开为 input.mf_input.mp_input，
        输入触发 ol#mp_xxx_list 出现 li.mp_item，第一项加 mp_highlighted；mousedown LI 即选中。

        - value: 期望最终选中的完整文本（用于 chip 校验）
        - search: 搜索关键字；省略时使用 value 本身。完整路径含下划线 K2 不分词，
          需要给一个能让目标候选出现的短词。
        """
        search_term = search or value
        scopes = [form_page] + [fr for fr in form_page.frames if fr != form_page.main_frame]
        for sc in scopes:
            try:
                # 1) 找到 label 对应值单元格里的 inputselect 占位 DIV，标记并点开
                located = sc.evaluate(r"""(label) => {
                    function visible(el) { return el && el.offsetParent !== null; }
                    function txt(el) { return ((el && el.textContent) || '').trim(); }
                    const norm = s => (s || '').replace(/[\s*：: ]/g, '');
                    const target = norm(label);
                    const labs = Array.from(document.querySelectorAll('td, label, span, div'))
                        .filter(el => visible(el) && norm(txt(el)) === target);
                    for (const lab of labs) {
                        const labelTd = lab.closest('td');
                        const valueCell = labelTd ? labelTd.nextElementSibling : null;
                        if (!valueCell) continue;
                        const sgl = valueCell.querySelector('.inputselectsgl');
                        if (sgl) {
                            sgl.setAttribute('data-claude-target', '1');
                            return {ok: true, kind: 'inputselectsgl'};
                        }
                        const inp = valueCell.querySelector('input.mp_input, input.mf_input');
                        if (inp) {
                            inp.setAttribute('data-claude-target', '1');
                            return {ok: true, kind: 'mp_input'};
                        }
                    }
                    return {ok: false, reason: 'no-control'};
                }""", label)
                if not (located and located.get('ok')):
                    continue

                container = sc.locator("[data-claude-target='1']").first
                container.click()
                form_page.wait_for_timeout(500)

                # 2) 找展开后的真实输入框，再次 click 确保聚焦
                inp_loc = sc.locator("input.mp_input:visible, input.mf_input:visible").first
                if inp_loc.count() == 0:
                    inp_loc = container.locator("input").first
                inp_loc.click()
                form_page.wait_for_timeout(300)

                # 3) 用 press_sequentially 直接对 input 输入；这样焦点不会跑掉
                try:
                    inp_loc.press_sequentially(search_term, delay=60)
                except Exception:
                    # 旧版 API 兜底
                    sc.keyboard.type(search_term, delay=60)

                # 3.1) K2 autocomplete 监听 keyup/input 事件，press_sequentially 可能被合并；
                # 强制 dispatch keyup + input 事件让候选下拉刷新
                try:
                    sc.evaluate(r"""() => {
                        const inp = document.querySelector("input.mp_input[data-claude-target], input.mf_input[data-claude-target]")
                            || document.querySelector("input.mp_input, input.mf_input");
                        if (!inp) return {ok: false};
                        // K2 mp/mf widget 监听这几个事件
                        ['input', 'keyup', 'keydown', 'change'].forEach(ev => {
                            inp.dispatchEvent(new Event(ev, {bubbles: true}));
                        });
                        // 显式触发 KeyboardEvent keyup 以模拟真实键盘
                        inp.dispatchEvent(new KeyboardEvent('keyup', {bubbles: true, key: 'a', keyCode: 65}));
                        return {ok: true, val: inp.value};
                    }""")
                except Exception:
                    pass
                # 等 ajax 候选回来
                form_page.wait_for_timeout(2500)

                # 诊断：input 当前值 + 候选 ol 数量（仅在 ol_count=0 时打印）
                try:
                    diag = sc.evaluate(r"""() => {
                        function visible(el) { return el && el.offsetParent !== null; }
                        const inps = Array.from(document.querySelectorAll('input.mp_input, input.mf_input')).filter(visible);
                        const ols = Array.from(document.querySelectorAll('ol.mp_list')).filter(visible);
                        const lis = ols.flatMap(ol => Array.from(ol.querySelectorAll('li')));
                        return {
                            inp_count: inps.length,
                            inp_values: inps.map(i => (i.value||'').slice(0, 60)),
                            ol_count: ols.length,
                            li_count: lis.length,
                            highlighted: document.querySelector('li.mp_highlighted') ? (document.querySelector('li.mp_highlighted').textContent||'').slice(0,60) : null,
                        };
                    }""")
                except Exception as e:
                    diag = None
                    print(f"  [DIAG-ERR] {e}", flush=True)

                # 3.2) 候选还没出来时，退格补字让 K2 重新触发搜索
                if (not diag) or diag.get('ol_count', 0) == 0:
                    print("  [RETRY] 候选未出现，退格 1 字符重输", flush=True)
                    try:
                        inp_loc.press("End")
                        inp_loc.press("Backspace")
                        form_page.wait_for_timeout(500)
                        # 重新输入最后一个字符
                        last_char = search_term[-1]
                        inp_loc.press_sequentially(last_char, delay=80)
                        form_page.wait_for_timeout(2000)
                        diag2 = sc.evaluate(r"""() => {
                            function visible(el) { return el && el.offsetParent !== null; }
                            const ols = Array.from(document.querySelectorAll('ol.mp_list')).filter(visible);
                            return {ol_count: ols.length, highlighted: document.querySelector('li.mp_highlighted') ? (document.querySelector('li.mp_highlighted').textContent||'').slice(0,60) : null};
                        }""")
                        print(f"  [RETRY-DIAG] {diag2}", flush=True)
                    except Exception as e:
                        print(f"  [RETRY-ERR] {e}", flush=True)

                # 3.5) 首选：直接按 Enter，K2 autocomplete 会选中 highlighted 项
                inp_loc.press("Enter")
                form_page.wait_for_timeout(800)

                # 验证 chip 是否已经出现 — 如果出现就跳过 LI 兜底
                chip_ok = sc.evaluate(r"""(args) => {
                    const value = args[0];
                    const search = args[1];
                    function visible(el) { return el && el.offsetParent !== null; }
                    function txt(el) { return ((el && el.textContent) || '').trim(); }
                    const chips = Array.from(document.querySelectorAll('ol.mf_list li.mf_item, .inputselectsgl li.mf_item')).filter(visible);
                    return chips.some(c => txt(c).includes(value) || txt(c).includes(search));
                }""", [value, search_term])
                if chip_ok:
                    return True, {'kind': 'inputselect-enter', 'picked': search_term}

                # 4) 兜底：去 document 全局找候选 ol，点 highlighted/匹配 LI
                pick = sc.evaluate(r"""(args) => {
                    const value = args[0];
                    const search = args[1];
                    function visible(el) { return el && el.offsetParent !== null; }
                    function txt(el) { return ((el && el.textContent) || '').trim(); }
                    const ols = Array.from(document.querySelectorAll('ol.mp_list')).filter(visible);
                    for (const ol of ols) {
                        let li = ol.querySelector('li.mp_highlighted');
                        if (!li && ol.id) {
                            li = document.getElementById(ol.id.replace('_list', '_highlighted'));
                        }
                        if (!li) {
                            const items = Array.from(ol.querySelectorAll('li.mp_item, li.mp_selectable')).filter(visible);
                            li = items.find(x => txt(x).includes(value)) || items.find(x => txt(x).includes(search)) || items[0];
                        }
                        if (li) {
                            li.setAttribute('data-claude-pick', '1');
                            return {ok: true, count: ol.children.length, text: txt(li).slice(0, 80)};
                        }
                    }
                    return {ok: false, reason: 'no-candidate', ol_count: ols.length};
                }""", [value, search_term])
                if not (pick and pick.get('ok')):
                    print(f"  [WARN] {label}: 候选未出现 ({pick})", flush=True)
                    if retries > 0:
                        sc.keyboard.press("Escape")
                        form_page.wait_for_timeout(500)
                        return fill_inputselect_field(label, value, search=search_term, retries=retries - 1)
                    return False, pick

                # 5) mousedown + click 候选 LI（录制显示用户先 mousedown 才生效）
                li_loc = sc.locator("[data-claude-pick='1']").first
                li_loc.dispatch_event("mousedown")
                form_page.wait_for_timeout(200)
                li_loc.dispatch_event("click")
                form_page.wait_for_timeout(800)

                # 6) 验证：chip 出现在 ol.mf_list 里（已选中状态）
                ok = sc.evaluate(r"""(args) => {
                    const value = args[0];
                    const search = args[1];
                    function visible(el) { return el && el.offsetParent !== null; }
                    function txt(el) { return ((el && el.textContent) || '').trim(); }
                    const chips = Array.from(document.querySelectorAll('ol.mf_list li.mf_item')).filter(visible);
                    return chips.some(c => txt(c).includes(value) || txt(c).includes(search));
                }""", [value, search_term])
                return (True, {'kind': 'inputselect', 'picked': pick.get('text', '')}) if ok else (False, {'reason': 'chip-not-found'})
            except Exception as e:
                print(f"  [WARN] inputselect 异常 scope={sc.url[:40] if hasattr(sc,'url') else 'main'}: {e}", flush=True)
                continue
        return False, {'reason': 'no-scope-matched'}

    INPUTSELECT_LABELS = {"积分成本归属部门"}
    INPUTSELECT_SEARCH = {
        "积分成本归属部门": "海外教学服务运营组",  # 录制确认：此关键词能让目标候选唯一命中
    }

    # 按表单从上到下顺序填（部门已并入 FIELD_VALUES）
    # 注意：金额和「需要添加的豌豆币/魔力币数总共」由 STEP 5 单独填，不在这里
    for lab, val in FIELD_VALUES.items():
        if lab in INPUTSELECT_LABELS:
            ok, info = fill_inputselect_field(lab, val, search=INPUTSELECT_SEARCH.get(lab))
            kind = (info or {}).get('kind', 'inputselect')
            if ok:
                picked = (info or {}).get('picked', '')
                print(f"  [OK] {lab} = {val}  ({kind}) picked={picked[:40]}", flush=True)
            else:
                print(f"  [WARN] 未能自动填写「{lab}」({info})，请在 STEP 7 暂停期间手动填", flush=True)
        else:
            ok, info = fill_field(lab, val)
            if ok:
                print(f"  [OK] {lab} = {val}  ({info.get('kind')})", flush=True)
            else:
                print(f"  [WARN] 未能自动填写「{lab}」，请在 STEP 7 暂停期间手动填", flush=True)
        form_page.wait_for_timeout(400)

    print(f"[STEP 5] 填写「需要添加的豌豆币/魔力币数总共」= {TOTAL_AMOUNT}", flush=True)
    # 该 OA 使用 K2/泛微风格 xform，输入框带 subject 属性 — 精确匹配
    amount_loc, amount_scope = find_in_page_or_frames(
        form_page,
        lambda s: s.locator("input[subject*='需要添加豌豆币'][subject*='数量总和']").first,
    )
    if amount_loc is None:
        amount_loc, amount_scope = find_in_page_or_frames(
            form_page,
            lambda s: s.get_by_role("textbox", name="需要添加豌豆币/魔力币数量总和").first,
        )
    if amount_loc is None:
        amount_loc, amount_scope = find_in_page_or_frames(
            form_page,
            lambda s: s.locator("input[title*='需要添加豌豆币']").first,
        )
    if amount_loc is None:
        print("[ERROR] 找不到金额输入框，请手动填写", flush=True)
    else:
        try:
            amount_loc.click()
            amount_loc.fill("")
            amount_loc.fill(str(TOTAL_AMOUNT))
            print(f"  [OK] 已填写 {TOTAL_AMOUNT}", flush=True)
        except Exception as e:
            print(f"  [WARN] 填写失败: {e}", flush=True)

    form_page.wait_for_timeout(800)

    print(f"[STEP 6] 上传附件: {ATTACHMENT.name}", flush=True)
    # 策略：页面里多个 id=upload_<fdId>_div_buttom 的"点击上传"，每个都配一个
    # input[type=file]（通常 id 包含同样的 fdId）。先用 JS 找到当前可见按钮，
    # 再按 fdId 前缀配对到对应 input[type=file]，对它直接 set_input_files —
    # 既不会因父级 visible 失败，又能精确选到「附件」那一栏。
    upload_done = False
    for scope in [form_page] + [fr for fr in form_page.frames if fr != form_page.main_frame]:
        try:
            pair = scope.evaluate("""() => {
                const btns = Array.from(document.querySelectorAll('[id*="upload_"][id*="_div_buttom"]'));
                const visible = btns.filter(el => {
                    const t = (el.textContent || '').trim();
                    return t === '点击上传' && el.offsetParent !== null;
                });
                if (visible.length === 0) return null;
                const btn = visible[0];
                const m = btn.id.match(/^upload_(.+?)_div_buttom$/);
                const fdId = m ? m[1] : null;
                const inputs = Array.from(document.querySelectorAll('input[type="file"]'));
                let matched = null;
                if (fdId) {
                    matched = inputs.find(i =>
                        (i.id && i.id.includes(fdId)) ||
                        (i.name && i.name.includes(fdId))
                    );
                }
                if (!matched) {
                    // 兜底：从按钮往上找最近的、含 input[type=file] 的容器
                    let p = btn;
                    while (p && p !== document.body) {
                        const f = p.querySelector('input[type="file"]');
                        if (f) { matched = f; break; }
                        p = p.parentElement;
                    }
                }
                if (!matched) return {btnId: btn.id, fdId: fdId, inputId: null};
                // 确保 input 不会被 display:none 挡住 set_input_files
                // (Playwright 对 hidden file input 默认可以设置文件)
                return {
                    btnId: btn.id,
                    fdId: fdId,
                    inputId: matched.id || null,
                    inputName: matched.name || null
                };
            }""")
            if not pair:
                continue
            print(f"  [FOUND] 上传按钮 id={pair.get('btnId')} fdId={pair.get('fdId')} "
                  f"-> input id={pair.get('inputId')} name={pair.get('inputName')}", flush=True)

            input_id = pair.get("inputId")
            input_name = pair.get("inputName")
            file_loc = None
            if input_id:
                file_loc = scope.locator(f"input[type=file]#{input_id}").first
            elif input_name:
                file_loc = scope.locator(f"input[type=file][name='{input_name}']").first

            if file_loc is None:
                print(f"  [WARN] 配对的 input[type=file] 没有可用 id/name，跳过本 frame", flush=True)
                continue

            try:
                file_loc.set_input_files(str(ATTACHMENT))
                print(f"  [OK] 已对配对 input 设置文件: {ATTACHMENT.name}", flush=True)
                upload_done = True
                form_page.wait_for_timeout(4000)
                break
            except Exception as e:
                print(f"  [WARN] 配对 input set_input_files 失败: {e}", flush=True)

            # 退路：JS 派发 click 到该 input，捕获 file_chooser
            try:
                print(f"  [TRY] JS dispatch click 到配对 input + file_chooser 拦截", flush=True)
                with form_page.expect_file_chooser(timeout=8000) as fc_info:
                    scope.evaluate(
                        """(sel) => {
                            const el = document.querySelector(sel);
                            if (el) el.click();
                        }""",
                        f"input[type=file]#{input_id}" if input_id
                        else f"input[type=file][name='{input_name}']",
                    )
                file_chooser = fc_info.value
                file_chooser.set_files(str(ATTACHMENT))
                print(f"  [OK] 已通过 JS click + file_chooser 上传 {ATTACHMENT.name}", flush=True)
                upload_done = True
                form_page.wait_for_timeout(4000)
                break
            except Exception as e:
                print(f"  [WARN] JS click + file_chooser 失败: {e}", flush=True)
        except Exception as e:
            print(f"  [WARN] 当前 frame 上传逻辑异常: {e}", flush=True)
            continue

    if not upload_done:
        # 最终兜底：随便取一个 input[type=file]（可能选错列，仅当上面全失败再用）
        file_loc, _ = find_in_page_or_frames(
            form_page,
            lambda s: s.locator("input[type=file]").first,
        )
        if file_loc is None:
            print("[ERROR] 找不到文件上传控件，请手动上传", flush=True)
        else:
            try:
                file_loc.set_input_files(str(ATTACHMENT))
                print(f"  [OK] 已通过通用 input[type=file] 上传 {ATTACHMENT.name}", flush=True)
                upload_done = True
                form_page.wait_for_timeout(4000)
            except Exception as e:
                print(f"  [WARN] 通用 input[type=file] 上传失败: {e}", flush=True)

    if upload_done:
        # 校验：附件名应出现在页面上
        verify_loc, _ = find_in_page_or_frames(
            form_page,
            lambda s: s.locator(f":text('{ATTACHMENT.stem}')").first,
        )
        if verify_loc is not None:
            print(f"  [VERIFY] 附件名已出现在页面上", flush=True)
        else:
            print(f"  [VERIFY-WARN] 未在页面上看到附件名，请手动确认", flush=True)

    # 关键：上传完成后，K2/泛微 xform 通常会留一个 lui_dialog（带 mask）等用户点「确定/上传」
    # 如果不关掉它，后面点「流程处理」时 mask 会拦截事件
    print("[STEP 6.5] 检查/关闭上传后残留的对话框（确定/上传/完成）", flush=True)
    try:
        dlg_info = form_page.evaluate(
            """() => {
                const masks = Array.from(document.querySelectorAll('.lui_dialog_mask, .lui_dialog'));
                const visible = masks.filter(el => el.offsetParent !== null);
                if (visible.length === 0) return {count: 0, clicked: null};
                // 在可见对话框里找 确定/上传/完成/确认 按钮
                const wanted = ['确定', '上传', '完成', '确认', 'OK', 'Ok'];
                const dlgs = Array.from(document.querySelectorAll('.lui_dialog'))
                    .filter(el => el.offsetParent !== null);
                for (const dlg of dlgs) {
                    const btns = Array.from(dlg.querySelectorAll('a, button, div, span'));
                    for (const b of btns) {
                        const t = (b.textContent || '').trim();
                        if (wanted.includes(t) && b.offsetParent !== null) {
                            b.click();
                            return {count: visible.length, clicked: t};
                        }
                    }
                }
                return {count: visible.length, clicked: null};
            }"""
        )
        print(f"  [DLG] 可见对话框/遮罩元素数={dlg_info.get('count')}, 点击={dlg_info.get('clicked')}", flush=True)
        if dlg_info.get('clicked'):
            form_page.wait_for_timeout(2500)
        # 若仍有遮罩，尝试 ESC 关闭
        residual = form_page.evaluate(
            """() => Array.from(document.querySelectorAll('.lui_dialog_mask'))
                .filter(el => el.offsetParent !== null).length"""
        )
        if residual > 0:
            print(f"  [DLG] 仍有 {residual} 个 mask，按 ESC 尝试关闭", flush=True)
            form_page.keyboard.press("Escape")
            form_page.wait_for_timeout(800)
    except Exception as e:
        print(f"  [DLG-WARN] 检查对话框失败: {e}", flush=True)

    shot(form_page, "submit_step3_filled.png")

    print("[STEP 7] 暂停 15 秒供肉眼复核（确认金额 + 附件）...", flush=True)
    form_page.wait_for_timeout(15000)

    # 再次扫描遮罩 — 用户可能在 15s 复核期间打开了别的弹窗
    try:
        residual = form_page.evaluate(
            """() => Array.from(document.querySelectorAll('.lui_dialog_mask'))
                .filter(el => el.offsetParent !== null).length"""
        )
        if residual > 0:
            print(f"  [DLG] 提交前发现 {residual} 个遮罩，尝试 JS 隐藏", flush=True)
            form_page.evaluate(
                """() => {
                    document.querySelectorAll('.lui_dialog_mask').forEach(el => {
                        if (el.offsetParent !== null) el.style.display = 'none';
                    });
                }"""
            )
    except Exception:
        pass

    print("[STEP 8] 点击「流程处理 / 提交」", flush=True)
    # K2/泛微 xform：工具栏按钮多为 <a> / <div> + 文本「流程处理」「提交」「发送」
    submit_loc, _ = find_in_page_or_frames(
        form_page,
        lambda s: s.get_by_role("button", name="流程处理").first,
    )
    if submit_loc is None:
        submit_loc, _ = find_in_page_or_frames(
            form_page,
            lambda s: s.get_by_role("button", name="提交").first,
        )
    if submit_loc is None:
        submit_loc, _ = find_in_page_or_frames(
            form_page,
            lambda s: s.locator(
                "a:has-text('流程处理'), a:has-text('提交'), a:has-text('发送'), "
                "div.lui_toolbar_btn:has-text('流程处理'), div.lui_toolbar_btn:has-text('提交'), "
                "button:has-text('流程处理'), button:has-text('提交'), input[type=submit]"
            ).first,
        )
    if submit_loc is None:
        # JS 兜底：先清掉所有遮罩，再找文本为「流程处理」且可点击的元素
        try:
            clicked = form_page.evaluate(
                """() => {
                    // 清除一切可能拦截事件的 mask
                    document.querySelectorAll('.lui_dialog_mask, .ui-widget-overlay').forEach(el => {
                        el.style.display = 'none';
                        el.style.pointerEvents = 'none';
                    });
                    const wanted = ['流程处理', '提交', '发送'];
                    const all = Array.from(document.querySelectorAll('a, button, div, span'));
                    for (const el of all) {
                        const t = (el.textContent || '').trim();
                        if (wanted.includes(t) && el.offsetParent !== null) {
                            el.click();
                            return t;
                        }
                    }
                    return null;
                }"""
            )
            if clicked:
                print(f"  [OK] JS 点击成功：{clicked}", flush=True)
                form_page.wait_for_timeout(5000)
                shot(form_page, "submit_step4_done.png")
                print(f"  [URL] {form_page.url}", flush=True)
                print("[OK] 流程结束。浏览器保持 60 秒便于确认。", flush=True)
                form_page.wait_for_timeout(60000)
                browser.close()
                sys.exit(0)
        except Exception as e:
            print(f"  [WARN] JS 兜底失败: {e}", flush=True)
        print("[ERROR] 找不到提交按钮，请手动点击", flush=True)
    else:
        click_done = False
        try:
            submit_loc.click(timeout=5000)
            print("  [OK] 已点击提交按钮", flush=True)
            click_done = True
        except Exception as e:
            print(f"  [WARN] 普通点击失败: {e}", flush=True)
            # mask 拦截 → 先 JS 隐藏 mask 再 force click
            try:
                form_page.evaluate(
                    """() => {
                        document.querySelectorAll('.lui_dialog_mask, .ui-widget-overlay').forEach(el => {
                            el.style.display = 'none';
                            el.style.pointerEvents = 'none';
                        });
                    }"""
                )
                submit_loc.click(timeout=5000, force=True)
                print("  [OK] 已点击提交按钮（force after mask clear）", flush=True)
                click_done = True
            except Exception as e2:
                print(f"  [WARN] force 点击仍失败: {e2}", flush=True)
        if click_done:
            form_page.wait_for_timeout(5000)
            # 提交后通常会再弹一个「下一步处理人 / 审批意见」对话框 → 自动点「发送/确定」
            try:
                followup = form_page.evaluate(
                    """() => {
                        const wanted = ['发送', '确定', '提交', '同意', '确认'];
                        const dlgs = Array.from(document.querySelectorAll('.lui_dialog'))
                            .filter(el => el.offsetParent !== null);
                        for (const dlg of dlgs) {
                            const btns = Array.from(dlg.querySelectorAll('a, button, div, span'));
                            for (const b of btns) {
                                const t = (b.textContent || '').trim();
                                if (wanted.includes(t) && b.offsetParent !== null) {
                                    b.click();
                                    return t;
                                }
                            }
                        }
                        return null;
                    }"""
                )
                if followup:
                    print(f"  [FOLLOWUP] 已点击后续对话框按钮：{followup}", flush=True)
                    form_page.wait_for_timeout(5000)
                else:
                    print(f"  [FOLLOWUP] 未发现需要确认的后续对话框", flush=True)
            except Exception as e:
                print(f"  [FOLLOWUP-WARN] {e}", flush=True)

    shot(form_page, "submit_step4_done.png")
    print(f"  [URL] {form_page.url}", flush=True)
    print("[OK] 流程结束。浏览器保持 60 秒便于确认。", flush=True)
    form_page.wait_for_timeout(60000)
    browser.close()
