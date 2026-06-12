#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""六一工作台 → 企微群发任务配置 → 豌豆素质 Tab → 新建群发任务（UI 自动化）

输入：
  --template-name  对应已建好的话术模板名（任务名也用同名）
  --user-group     选用户群名，模糊搜索后选第一条匹配
  --start-min      开始时间 = 现在 + N 分钟（默认 22，必须 >= 21）
  --end-min        结束时间 = 现在 + N 分钟（默认 28800 = 20 天）
  --teams          执行团队，逗号分隔；不填按用户群名自动推：
                     名字带「港澳/香港/澳门」→ 港澳1组,港澳2组,港澳组
                     名字带「亚欧/美澳」     → 美澳1组,美澳2组,美澳3组,美澳4组,美澳5组
  --no-submit      停在「预览」按钮前不点（默认开启，安全）
  --click-preview  跑完点一下「预览」（用于人工确认完整表单），不点「确定」
  --keep-open      跑完留住浏览器
"""
import sys, io, time, argparse, datetime as dt
from pathlib import Path
from playwright.sync_api import sync_playwright, TimeoutError as PWTimeout

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

SCRIPT_DIR = Path(__file__).resolve().parent
BROWSER_PROFILE = SCRIPT_DIR / "browser_profile"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

DEFAULT_TEAM_TAB = "豌豆素质"
DEFAULT_SUBJECT = "豌豆益智"
DEFAULT_TASK_TYPE = "续费"
DEFAULT_PRIORITY = "高级群发"
DEFAULT_REPLY_FILTER = "过滤"
DEFAULT_EXEC_MODE = "定时执行"
DEFAULT_SEND_MODE = "手动群发"


def log(m): print(m, flush=True)


def infer_teams(user_group: str) -> list:
    g = user_group
    if any(k in g for k in ["港澳", "香港", "澳门"]):
        return ["港澳1组", "港澳2组", "港澳组"]
    if any(k in g for k in ["亚欧", "美澳", "欧美"]):
        return ["美澳1组", "美澳2组", "美澳3组", "美澳4组", "美澳5组"]
    return []


def goto_task_page(ctx, portal_page):
    """portal → 六一工作台 → 企微群发任务配置 → 豌豆素质 tab"""
    log("[STEP] 进 portal")
    portal_page.goto(PORTAL_URL, timeout=30000)
    portal_page.wait_for_load_state("domcontentloaded")
    portal_page.wait_for_timeout(3000)

    if "login.61info.cn" in portal_page.url:
        log("[INFO] 需要扫码登录...")
        deadline = time.time() + 300
        while time.time() < deadline:
            if "login.61info.cn" not in portal_page.url:
                break
            portal_page.wait_for_timeout(2000)

    log("[STEP] 点开「六一工作台」瓷砖")
    with ctx.expect_page(timeout=30000) as new_page_info:
        portal_page.locator("p:has-text('六一工作台')").first.click()
    liuyi = new_page_info.value
    liuyi.wait_for_load_state("domcontentloaded", timeout=30000)
    liuyi.wait_for_timeout(3000)

    log("[STEP] 点左侧菜单「企微群发任务配置」")
    try:
        liuyi.evaluate(
            """() => {
                const all = Array.from(document.querySelectorAll('span, a, li'));
                const t = all.find(el => (el.textContent || '').trim() === '企微群发任务配置');
                if (!t) throw new Error('没找到「企微群发任务配置」节点');
                const c = t.closest('li, a, [role=menuitem]') || t;
                c.scrollIntoView({block:'center'}); c.click();
            }"""
        )
    except Exception as e:
        log(f"[WARN] evaluate click 失败 {e}，回退 force click")
        liuyi.locator("text=企微群发任务配置").first.click(force=True)
    liuyi.wait_for_timeout(2500)

    log(f"[STEP] 切换到「{DEFAULT_TEAM_TAB}」Tab")
    tab = liuyi.locator(f"text={DEFAULT_TEAM_TAB}").first
    tab.wait_for(state="visible", timeout=15000)
    tab.click()
    liuyi.wait_for_timeout(1500)
    return liuyi


def open_create_dialog(page):
    log("  · 点「新建任务」")
    page.locator("button:has-text('新建任务')").first.click()
    page.wait_for_timeout(2500)
    deadline = time.time() + 15
    while time.time() < deadline:
        n = page.evaluate(
            """() => Array.from(document.querySelectorAll('.el-dialog__wrapper'))
                  .filter(d => d.style.display !== 'none' && getComputedStyle(d).display !== 'none').length"""
        )
        if n > 0:
            log(f"  · dialog 已弹出（{n} 个）")
            page.wait_for_timeout(1000)
            # 调试：dump 所有 form-item 的 label 文字 + 内含组件类型
            structure = page.evaluate(
                """() => {
                    const wrappers = Array.from(document.querySelectorAll('.el-dialog__wrapper'))
                      .filter(d => d.style.display !== 'none' && getComputedStyle(d).display !== 'none');
                    const out = [];
                    for (const w of wrappers) {
                        const items = Array.from(w.querySelectorAll('.el-form-item'));
                        for (const it of items) {
                            const lbl = it.querySelector('label, .el-form-item__label');
                            const lblTxt = lbl ? (lbl.textContent || '').trim() : '(无 label)';
                            const types = [];
                            if (it.querySelector('.el-select')) types.push('el-select');
                            if (it.querySelector('.el-cascader')) types.push('el-cascader');
                            if (it.querySelector('.el-date-editor')) types.push('el-date');
                            if (it.querySelector('.el-radio-group')) types.push('el-radio');
                            if (it.querySelector('input[type=text]') || it.querySelector('.el-input__inner')) types.push('input');
                            out.push({label: lblTxt, types: types.join(',') || '?'});
                        }
                    }
                    return out;
                }"""
            )
            log("  [DEBUG] dialog 表单结构:")
            for s in structure:
                log(f"    - 「{s['label']}」 → {s['types']}")
            return
        page.wait_for_timeout(500)
    raise RuntimeError("点完新建任务后没有弹出 dialog（可能跳到了别的页面）")


def select_dropdown_in_dialog(page, select_index: int, option_text: str):
    """点开当前可见 dialog 里第 N 个 .el-select，选 option_text。"""
    selects = page.locator(".el-dialog__body:visible .el-select")
    target = selects.nth(select_index)
    target.scroll_into_view_if_needed()
    target.click()
    page.wait_for_timeout(400)
    dropdown = page.locator(".el-select-dropdown").locator(
        "xpath=self::*[not(contains(@style,'display: none'))]"
    ).last
    item = dropdown.locator(f".el-select-dropdown__item:has-text('{option_text}')").first
    item.wait_for(state="visible", timeout=10000)
    item.click()
    page.wait_for_timeout(500)


def _find_form_item(page, label_text: str):
    """在可见 dialog 里找 .el-form-item，其 .el-form-item__label 文本含 label_text。"""
    return page.locator(".el-dialog__wrapper:visible .el-form-item").filter(
        has=page.locator(".el-form-item__label", has_text=label_text)
    ).first


def select_by_label(page, label_text: str, option_text: str):
    item = _find_form_item(page, label_text)
    sel = item.locator(".el-select").first
    sel.scroll_into_view_if_needed()
    sel.click()
    page.wait_for_timeout(400)
    dropdown = page.locator(".el-select-dropdown").locator(
        "xpath=self::*[not(contains(@style,'display: none'))]"
    ).last
    opt = dropdown.locator(f".el-select-dropdown__item:has-text('{option_text}')").first
    opt.wait_for(state="visible", timeout=10000)
    opt.click()
    page.wait_for_timeout(500)


def fill_input_by_label(page, label_text: str, value: str):
    item = _find_form_item(page, label_text)
    inp = item.locator("input").first
    inp.fill("")
    inp.fill(value)
    page.wait_for_timeout(300)


def search_and_pick(page, label_text: str, query: str):
    item = _find_form_item(page, label_text)
    sel = item.locator(".el-select").first
    sel.scroll_into_view_if_needed()
    sel.click()
    page.wait_for_timeout(400)
    inp = item.locator(".el-select input").first
    inp.fill(query)
    page.wait_for_timeout(1500)
    dropdown = page.locator(".el-select-dropdown").locator(
        "xpath=self::*[not(contains(@style,'display: none'))]"
    ).last
    opt = dropdown.locator(f".el-select-dropdown__item:has-text('{query}')").first
    try:
        opt.wait_for(state="visible", timeout=8000)
        opt.click()
    except PWTimeout:
        opt = dropdown.locator(".el-select-dropdown__item:not(.is-disabled)").first
        opt.click()
    page.wait_for_timeout(500)


def pick_teams(page, label_text: str, teams: list):
    if not teams:
        return
    item = _find_form_item(page, label_text)
    box = item.locator(".el-select, .el-cascader").first
    box.scroll_into_view_if_needed()
    box.click()
    page.wait_for_timeout(1200)

    # 调试：dump cascader 当前可见节点
    nodes_dump = page.evaluate(
        """() => {
            const poppers = Array.from(document.querySelectorAll('.el-cascader__dropdown, .el-cascader-panel, .el-popper'))
              .filter(d => d.offsetParent !== null);
            const out = [];
            for (const p of poppers) {
                const nodes = Array.from(p.querySelectorAll('.el-cascader-node, .el-tree-node, li'));
                for (const n of nodes) {
                    out.push((n.textContent || '').replace(/\\s+/g,'').slice(0, 40));
                }
            }
            return out.slice(0, 30);
        }"""
    )
    log(f"  [DEBUG] cascader 顶层节点: {nodes_dump}")

    # 尝试找 cascader 的搜索输入框（filterable=true 时存在）
    search_inp = item.locator("input.el-cascader__search-input, input[placeholder*='请输入']").first
    has_search = search_inp.count() > 0
    log(f"  [DEBUG] cascader 是否可搜索: {has_search}")

    for t in teams:
        log(f"    · 勾选: {t}")
        if has_search:
            try:
                search_inp.fill("")
                search_inp.fill(t)
                page.wait_for_timeout(1500)
                # 搜索结果一般在 .el-cascader__suggestion-panel 或 .el-cascader-panel 里
                ok = page.evaluate(
                    """(name) => {
                        const poppers = Array.from(document.querySelectorAll('.el-cascader__dropdown, .el-cascader-panel, .el-popper, .el-cascader__suggestion-panel'))
                          .filter(d => d.offsetParent !== null);
                        for (const p of poppers) {
                            // 搜索建议项一般是 li.el-cascader__suggestion-item
                            const sugs = Array.from(p.querySelectorAll('.el-cascader__suggestion-item, li, .el-cascader-node'));
                            const target = sugs.find(s => (s.textContent || '').includes(name));
                            if (target) {
                                target.scrollIntoView({block:'center'});
                                target.click();
                                return {ok: true, txt: (target.textContent || '').slice(0, 60)};
                            }
                        }
                        return {ok: false};
                    }""",
                    t,
                )
                log(f"      → 搜索点击: {ok}")
            except Exception as e:
                log(f"    [WARN] 搜索失败 {e}")
            page.wait_for_timeout(600)
        else:
            log(f"    [WARN] cascader 不支持搜索，跳过「{t}」（请人工补）")

    # 点 dialog body 关闭 popper（避免点 header 失败）
    try:
        item.locator("xpath=..").first.click()
    except Exception:
        pass
    page.keyboard.press("Escape")
    page.wait_for_timeout(500)


def fill_datetime_range(page, label_text: str, start_dt: dt.datetime, end_dt: dt.datetime):
    item = _find_form_item(page, label_text)
    inputs = item.locator("input")
    fmt = "%Y-%m-%d %H:%M:%S"
    s = start_dt.strftime(fmt)
    e = end_dt.strftime(fmt)
    log(f"  · 开始: {s}")
    inputs.nth(0).click()
    page.wait_for_timeout(300)
    inputs.nth(0).fill(s)
    page.keyboard.press("Enter")
    page.wait_for_timeout(600)
    log(f"  · 结束: {e}")
    inputs.nth(1).click()
    page.wait_for_timeout(300)
    inputs.nth(1).fill(e)
    page.keyboard.press("Enter")
    page.wait_for_timeout(600)
    # 关 picker popper：先点 popper 内的「确定」按钮
    page.evaluate(
        """() => {
            const poppers = Array.from(document.querySelectorAll('.el-picker-panel, .el-date-range-picker'))
              .filter(d => d.offsetParent !== null);
            for (const p of poppers) {
                const btn = Array.from(p.querySelectorAll('button')).find(b => {
                    const t = (b.textContent || '').replace(/\\s+/g,'').trim();
                    return t === '确定' || t === '确认';
                });
                if (btn) btn.click();
            }
        }"""
    )
    page.wait_for_timeout(500)
    # 再 dispatch 一个 mousedown 到 document.body 触发 element-ui 的「点外部关闭」逻辑
    page.evaluate(
        """() => {
            const ev = new MouseEvent('mousedown', {bubbles: true, cancelable: true, view: window});
            document.body.dispatchEvent(ev);
        }"""
    )
    page.wait_for_timeout(400)
    # 终极兜底：把所有还在的 picker popper 设成 display:none + 阻止 pointer events
    page.evaluate(
        """() => {
            document.querySelectorAll('.el-picker-panel, .el-date-range-picker').forEach(p => {
                p.style.display = 'none';
                p.style.pointerEvents = 'none';
            });
        }"""
    )
    page.wait_for_timeout(400)


def fill_form(page, args, start_dt, end_dt, exec_dt, teams):
    log(f"  · 发送方式: {DEFAULT_SEND_MODE}")
    select_by_label(page, "发送方式", DEFAULT_SEND_MODE)

    # 选完后小弹窗里点「确定」才会进大表单
    log("  · 点小弹窗「确定」进大表单")
    clicked = page.evaluate(
        """() => {
            const wrappers = Array.from(document.querySelectorAll('.el-dialog__wrapper'))
              .filter(d => d.style.display !== 'none' && getComputedStyle(d).display !== 'none');
            for (const w of wrappers) {
                const btn = Array.from(w.querySelectorAll('.el-dialog__footer button, .el-dialog button'))
                  .find(b => /^(确 ?定|确认|下一步)$/.test((b.textContent || '').replace(/\\s+/g,'')));
                if (btn) {
                    btn.scrollIntoView({block:'center'});
                    btn.click();
                    return true;
                }
            }
            return false;
        }"""
    )
    if not clicked:
        raise RuntimeError("没在可见 dialog 里找到「确定/确认」按钮")
    page.wait_for_timeout(2500)

    # 等大表单出现：找「选择科目」label
    log("  · 等大表单字段渲染...")
    deadline = time.time() + 15
    labels = []
    while time.time() < deadline:
        labels = page.evaluate(
            """() => {
                const items = Array.from(document.querySelectorAll('.el-form-item'))
                  .filter(it => it.offsetParent !== null);
                const out = [];
                for (const it of items) {
                    const lbl = it.querySelector('label, .el-form-item__label');
                    if (lbl) out.push((lbl.textContent || '').trim());
                }
                return out;
            }"""
        )
        if any("选择科目" in l for l in labels):
            log(f"  · 字段已展开（{len(labels)} 个 form-item）: {labels}")
            break
        page.wait_for_timeout(500)
    else:
        log(f"  [DEBUG] 当前可见 form-item: {labels}")
        raise RuntimeError("点完确定后大表单没展开")

    log(f"  · 选科目: {DEFAULT_SUBJECT}")
    select_by_label(page, "选择科目", DEFAULT_SUBJECT)

    log(f"  · 任务类型: {DEFAULT_TASK_TYPE}")
    select_by_label(page, "任务类型", DEFAULT_TASK_TYPE)

    log(f"  · 任务名称: {args.template_name}")
    fill_input_by_label(page, "任务名称", args.template_name)

    log(f"  · 优先使用: {DEFAULT_PRIORITY}")
    select_by_label(page, "优先使用", DEFAULT_PRIORITY)

    log(f"  · 选用户群: {args.user_group}")
    search_and_pick(page, "选用户群", args.user_group)

    log(f"  · 选话术: {args.template_name}")
    search_and_pick(page, "选择话术", args.template_name)

    log(f"  · 沟通中和未回复: {DEFAULT_REPLY_FILTER}")
    select_by_label(page, "沟通中和未回复", DEFAULT_REPLY_FILTER)

    log(f"  · 任务时效")
    # 实时重算：六一限制 start ≥ 现在+10min，end ≤ 现在+10080min（=7 天）
    # 给点 buffer：start = 现在+15min（5min 缓冲），end = 现在+10000min（80min 缓冲，~6天22小时）
    fresh_now = dt.datetime.now().replace(microsecond=0)
    fresh_start = fresh_now + dt.timedelta(minutes=15)
    fresh_end = fresh_now + dt.timedelta(minutes=10000)
    if start_dt < fresh_start or end_dt > fresh_end:
        log(f"  [INFO] 重算时效以满足六一约束:")
        log(f"         start: {start_dt} → {fresh_start}")
        log(f"         end:   {end_dt} → {fresh_end}")
        start_dt, end_dt = fresh_start, fresh_end
    fill_datetime_range(page, "任务时效", start_dt, end_dt)

    log(f"  · 执行团队: {teams}")
    pick_teams(page, "执行团队", teams)

    log(f"  · 执行方式: {DEFAULT_EXEC_MODE}")
    select_by_label(page, "执行方式", DEFAULT_EXEC_MODE)
    page.wait_for_timeout(1000)

    # 定时执行可能会多出一个「执行时间」字段，dump 一下当前 label 看看
    labels_after = page.evaluate(
        """() => {
            const items = Array.from(document.querySelectorAll('.el-form-item'))
              .filter(it => it.offsetParent !== null);
            return items.map(it => {
                const lbl = it.querySelector('label, .el-form-item__label');
                return lbl ? (lbl.textContent || '').trim() : '?';
            });
        }"""
    )
    log(f"  [DEBUG] 选完执行方式后字段: {labels_after}")
    # 如果出现「执行时间」，把它填上
    if any("执行时间" in l for l in labels_after):
        log(f"  · 检测到「执行时间」字段，填入 {exec_dt}")
        item = _find_form_item(page, "执行时间")
        inp = item.locator("input").first
        inp.click()
        page.wait_for_timeout(800)
        inp.fill(exec_dt.strftime("%Y-%m-%d %H:%M:%S"))
        page.wait_for_timeout(800)
        # 重新点 input 确保 popper 打开
        inp.click()
        page.wait_for_timeout(800)

        # dump popper 里所有按钮文字方便调试
        popper_btns = page.evaluate(
            """() => {
                const poppers = Array.from(document.querySelectorAll('.el-picker-panel, .el-time-panel, .el-popper, [x-placement]'))
                  .filter(d => d.offsetParent !== null);
                const out = [];
                for (const p of poppers) {
                    for (const b of p.querySelectorAll('button, a, span')) {
                        const t = (b.textContent || '').replace(/\\s+/g,'').trim();
                        if (t && t.length <= 6) out.push(`${b.tagName}:${t}`);
                    }
                }
                return out;
            }"""
        )
        log(f"  [DEBUG] 当前 popper 里的可点元素: {popper_btns}")

        confirmed = page.evaluate(
            """() => {
                const poppers = Array.from(document.querySelectorAll('.el-picker-panel, .el-time-panel, .el-popper, [x-placement]'))
                  .filter(d => d.offsetParent !== null);
                for (const p of poppers) {
                    // 同时找 button 和 a / span（element-ui 的 picker 确定可能是 a 标签）
                    const cand = Array.from(p.querySelectorAll('button, a, span'));
                    const btn = cand.find(b => {
                        const t = (b.textContent || '').replace(/\\s+/g,'').trim();
                        return t === '确定' || t === '确认' || t === '确 定';
                    });
                    if (btn) { btn.click(); return true; }
                }
                return false;
            }"""
        )
        if not confirmed:
            log("  [WARN] popper 里还是没找到「确定」，回车兜底")
            page.keyboard.press("Enter")
        page.wait_for_timeout(1000)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--template-name", required=True)
    ap.add_argument("--user-group", required=True)
    ap.add_argument("--start-min", type=int, default=15)
    ap.add_argument("--start-at", default=None,
                    help="任务时效开始时间 'YYYY-MM-DD HH:MM'。指定后忽略 --start-min")
    ap.add_argument("--end-days", type=float, default=6.0,
                    help="任务时效结束 = 现在 + N 天（默认 6；总跨度必须 ≤ 7 天）")
    ap.add_argument("--exec-at", default=None,
                    help="定时执行的执行时间 'YYYY-MM-DD HH:MM'。不填默认取开始时间")
    ap.add_argument("--teams", default=None, help="逗号分隔；不填按用户群名自动推")
    ap.add_argument("--click-preview", action="store_true", help="跑完点「预览」")
    ap.add_argument("--submit", action="store_true",
                    help="跑完点「预览」→ 等预览页 →点「确认」真正创建任务（默认不点）")
    ap.add_argument("--keep-open", action="store_true")
    args = ap.parse_args()

    if args.start_min < 10:
        log(f"[WARN] start-min={args.start_min} 太短，建议 ≥ 15")

    teams = [t.strip() for t in args.teams.split(",")] if args.teams else infer_teams(args.user_group)
    if not teams:
        log("[WARN] 没有执行团队，--teams 没传也没匹配到规则；脚本会跑但执行团队会留空")

    now = dt.datetime.now().replace(microsecond=0)
    if args.start_at:
        start_dt = dt.datetime.strptime(args.start_at, "%Y-%m-%d %H:%M")
    else:
        start_dt = now + dt.timedelta(minutes=args.start_min)
    end_dt = now + dt.timedelta(days=args.end_days)
    exec_dt = dt.datetime.strptime(args.exec_at, "%Y-%m-%d %H:%M") if args.exec_at else start_dt
    # 校验：六一限制总跨度 ≤ 7 天
    if (end_dt - start_dt).total_seconds() > 7 * 24 * 3600:
        log(f"[WARN] 结束 - 开始 = {(end_dt - start_dt)} > 7 天，可能被服务端拒绝")

    log("=" * 60)
    log(f"  任务名 / 话术: {args.template_name}")
    log(f"  用户群       : {args.user_group}")
    log(f"  任务时效开始 : {start_dt}")
    log(f"  任务时效结束 : {end_dt}")
    log(f"  定时执行时间 : {exec_dt}")
    log(f"  执行团队     : {teams}")
    if args.submit:
        action = "填表 → 预览 → 确认（真创建）"
    elif args.click_preview:
        action = "填表 → 点预览（不确认）"
    else:
        action = "填表后停下"
    log(f"  动作         : {action}")
    log("=" * 60)

    stamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
    out_dir = SCRIPT_DIR / "exports" / dt.datetime.now().strftime("%Y%m%d") / "screenshots"
    out_dir.mkdir(parents=True, exist_ok=True)

    with sync_playwright() as p:
        BROWSER_PROFILE.mkdir(parents=True, exist_ok=True)
        ctx = p.chromium.launch_persistent_context(
            user_data_dir=str(BROWSER_PROFILE),
            headless=False, channel="chrome", slow_mo=60,
        )
        portal = ctx.new_page()

        try:
            page = goto_task_page(ctx, portal)
        except Exception as e:
            log(f"[ERROR] 进群发任务配置页失败: {e}")
            portal.screenshot(path=str(out_dir / f"task_open_fail_{stamp}.png"))
            ctx.close()
            sys.exit(2)

        # 落地页截图
        page.screenshot(path=str(out_dir / f"task_landing_{stamp}.png"), full_page=True)
        log(f"[INFO] 落地页截图: task_landing_{stamp}.png")

        try:
            open_create_dialog(page)
            fill_form(page, args, start_dt, end_dt, exec_dt, teams)
            page.screenshot(path=str(out_dir / f"task_filled_{stamp}.png"), full_page=True)
            log(f"\n[OK] 表单已填，截图: task_filled_{stamp}.png")

            if args.click_preview or args.submit:
                log("  · 点「预览」")
                clicked = page.evaluate(
                    """() => {
                        const wrappers = Array.from(document.querySelectorAll('.el-dialog__wrapper'))
                          .filter(d => d.style.display !== 'none' && getComputedStyle(d).display !== 'none');
                        for (const w of wrappers) {
                            const btns = Array.from(w.querySelectorAll('button'));
                            const btn = btns.find(b => (b.textContent||'').replace(/\\s+/g,'') === '预览');
                            if (btn && btn.offsetParent !== null) {
                                btn.scrollIntoView({block:'center'});
                                btn.click();
                                return true;
                            }
                        }
                        return false;
                    }"""
                )
                if not clicked:
                    raise RuntimeError("找不到可见的「预览」按钮")
                page.wait_for_timeout(2500)
                page.screenshot(path=str(out_dir / f"task_preview_{stamp}.png"), full_page=True)
                log(f"[OK] 预览截图: task_preview_{stamp}.png")

            if args.submit:
                log("  · 预览页点「确认 / 确定」真创建任务")
                # 先 dump 预览页所有可见按钮
                btns_dump = page.evaluate(
                    """() => {
                        const wrappers = Array.from(document.querySelectorAll('.el-dialog__wrapper'))
                          .filter(d => d.style.display !== 'none' && getComputedStyle(d).display !== 'none');
                        const out = [];
                        for (const w of wrappers) {
                            for (const b of w.querySelectorAll('button')) {
                                if (b.offsetParent !== null) out.push((b.textContent||'').replace(/\\s+/g,'').trim());
                            }
                        }
                        return out;
                    }"""
                )
                log(f"  [DEBUG] 预览页可见按钮: {btns_dump}")
                clicked = page.evaluate(
                    """() => {
                        const wrappers = Array.from(document.querySelectorAll('.el-dialog__wrapper'))
                          .filter(d => d.style.display !== 'none' && getComputedStyle(d).display !== 'none');
                        for (const w of wrappers) {
                            const btns = Array.from(w.querySelectorAll('button')).filter(b => b.offsetParent !== null);
                            const btn = btns.find(b => /^(确认创建|确认|确定|确 定|提交|创建|保存|发布)$/.test((b.textContent||'').replace(/\\s+/g,'')));
                            if (btn) { btn.scrollIntoView({block:'center'}); btn.click(); return (btn.textContent||'').trim(); }
                        }
                        return null;
                    }"""
                )
                if clicked:
                    log(f"  · 已点: {clicked}")
                else:
                    log("  [WARN] 预览页找不到确认按钮")
                page.wait_for_timeout(3000)
                page.screenshot(path=str(out_dir / f"task_submitted_{stamp}.png"), full_page=True)
                log(f"[OK] 提交后截图: task_submitted_{stamp}.png")
            elif not args.click_preview:
                log("[INFO] --click-preview / --submit 都没开，停在表单上")
        except Exception as e:
            log(f"[ERROR] 填表失败: {e}")
            page.screenshot(path=str(out_dir / f"task_fail_{stamp}.png"), full_page=True)
            raise

        if args.keep_open:
            log("[STAY] --keep-open，浏览器保留 5 分钟")
            try:
                page.wait_for_timeout(300_000)
            except Exception:
                pass
        ctx.close()


if __name__ == "__main__":
    main()
