"""录制用户在 OA 表单页填写「积分成本归属部门」的真实交互。

用法：
    python scripts/oa_login/record_dept_flow.py [--keyword 海外]

流程：
  1. 复用 auth_state.json 打开 portal，自动点击「豌豆币添加申请」入口，等待表单 tab。
  2. 向 form_page 下所有 frame 注入事件监听器，实时记录：
       - click / mousedown / focus / blur / input / change / keydown
       - DOM mutations（新增/属性变化）
     每 0.5s 轮询一次把事件捞回来，写到 record_events.ndjson。
  3. 终端高亮包含 --keyword 的事件（默认「积分成本」）。
  4. 操作完后按 Ctrl+C 退出，会额外 dump 一个 record_events.json 汇总。

注意：脚本不会自己填字段，让你手动操作；目的是看清你的真实点击/输入路径。
"""
from __future__ import annotations

import argparse
import json
import os
import signal
import sys
import time
from pathlib import Path

from playwright.sync_api import sync_playwright

HERE = Path(__file__).parent
AUTH_STATE = HERE / "auth_state.json"
EVENTS_NDJSON = HERE / "record_events.ndjson"
EVENTS_JSON = HERE / "record_events.json"
STOP_FLAG = HERE / "record_events.STOP"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

RECORDER_JS = r"""
(() => {
  if (window.__claudeRecorder) return 'already-installed';
  const buf = [];
  window.__claudeRecorder = {
    drain() { const out = buf.splice(0, buf.length); return out; },
  };

  function elInfo(el) {
    if (!el || el.nodeType !== 1) return null;
    const rect = (() => { try { return el.getBoundingClientRect(); } catch (e) { return {}; } })();
    let txt = '';
    try { txt = (el.innerText || el.textContent || '').trim().slice(0, 120); } catch (e) {}
    const attrs = {};
    try {
      for (const a of el.attributes || []) attrs[a.name] = a.value;
    } catch (e) {}
    // 往上记录最多 4 层 ancestor 的 tag/id/class，方便复现
    const chain = [];
    let p = el;
    for (let i = 0; i < 4 && p; i++) {
      chain.push({
        tag: p.tagName,
        id: p.id || '',
        cls: (p.className && typeof p.className === 'string') ? p.className.slice(0, 80) : '',
      });
      p = p.parentElement;
    }
    return {
      tag: el.tagName,
      id: el.id || '',
      name: el.name || '',
      cls: (el.className && typeof el.className === 'string') ? el.className.slice(0, 120) : '',
      type: el.type || '',
      value: (el.value !== undefined) ? String(el.value).slice(0, 200) : '',
      text: txt,
      rect: { x: Math.round(rect.x || 0), y: Math.round(rect.y || 0), w: Math.round(rect.width || 0), h: Math.round(rect.height || 0) },
      attrs,
      chain,
    };
  }

  function push(kind, ev, extra) {
    try {
      const target = ev && ev.target ? elInfo(ev.target) : null;
      buf.push({
        t: Date.now(),
        kind,
        target,
        extra: extra || null,
      });
    } catch (e) {}
  }

  // 指针 / 键盘 / 输入事件
  ['mousedown', 'click', 'focus', 'blur', 'input', 'change', 'keydown'].forEach(k => {
    document.addEventListener(k, (ev) => {
      const extra = {};
      if (k === 'keydown') extra.key = ev.key;
      if (k === 'input' || k === 'change') {
        try { extra.value = String(ev.target.value).slice(0, 200); } catch (e) {}
      }
      push(k, ev, extra);
    }, true);
  });

  // DOM 变更监听
  try {
    const mo = new MutationObserver((muts) => {
      for (const m of muts) {
        if (m.type === 'childList') {
          for (const n of m.addedNodes) {
            if (n.nodeType === 1) {
              buf.push({
                t: Date.now(),
                kind: 'dom-add',
                target: elInfo(n),
                extra: null,
              });
            }
          }
        } else if (m.type === 'attributes') {
          buf.push({
            t: Date.now(),
            kind: 'dom-attr',
            target: elInfo(m.target),
            extra: { attr: m.attributeName },
          });
        }
      }
    });
    mo.observe(document.documentElement, { childList: true, subtree: true, attributes: true });
  } catch (e) {}

  return 'installed';
})();
"""


def install_recorder(scope, frame_label: str, log) -> None:
    try:
        result = scope.evaluate(RECORDER_JS)
        log(f"  [INSTALL] {frame_label} -> {result}")
    except Exception as e:
        log(f"  [INSTALL-ERR] {frame_label} -> {e}")


def install_in_page(page, log) -> None:
    """给 page 主文档 + 所有 frame 注入 recorder。"""
    install_recorder(page, f"page {page.url[:60]}", log)
    for fr in page.frames:
        if fr == page.main_frame:
            continue
        try:
            install_recorder(fr, f"  frame {fr.url[:60]}", log)
        except Exception as e:
            log(f"  [INSTALL-ERR] frame {fr.url[:60]} -> {e}")


def install_in_all_frames(page, log) -> None:
    install_in_page(page, log)


def drain_all_pages(context) -> list[dict]:
    """遍历 context 里所有 page 的所有 frame，drain 事件。"""
    out: list[dict] = []
    for pg in list(context.pages):
        try:
            scopes = [pg] + [fr for fr in pg.frames if fr != pg.main_frame]
        except Exception:
            continue
        for sc in scopes:
            try:
                chunk = sc.evaluate("() => (window.__claudeRecorder ? window.__claudeRecorder.drain() : [])")
                if chunk:
                    url = sc.url if hasattr(sc, "url") else ""
                    for ev in chunk:
                        ev["frame_url"] = url[:80]
                    out.extend(chunk)
            except Exception:
                continue
    return out


ENTRY_TEXT = "豌豆币添加申请"


def try_click_entry(page, log) -> tuple[bool, str]:
    """复用 submit_oa.py 的多策略点击「豌豆币添加申请」入口。"""
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
            log(f"  [TRY] {label} count={loc.count()}")
            try:
                loc.scroll_into_view_if_needed(timeout=3000)
            except Exception:
                pass
            try:
                loc.click(timeout=3000)
                log(f"  [OK] clicked via {label}")
                return True, label
            except Exception as e:
                log(f"  [SKIP] {label} normal click failed: {e}")
                try:
                    loc.click(timeout=3000, force=True)
                    return True, f"{label}+force"
                except Exception as e2:
                    log(f"  [SKIP] {label} force click failed: {e2}")
        except Exception as e:
            log(f"  [SKIP] {label} setup error: {e}")

    for i, frame in enumerate(page.frames):
        if frame == page.main_frame:
            continue
        try:
            loc = frame.get_by_text(ENTRY_TEXT, exact=True).first
            if loc.count() > 0:
                log(f"  [TRY] iframe[{i}] url={frame.url[:60]} count={loc.count()}")
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
                        log(f"  [SKIP] iframe[{i}] click failed: {e}")
        except Exception:
            continue

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
        log(f"  [SKIP] js evaluate failed: {e}")

    return False, "none"


def click_entry(page, log) -> bool:
    """先等 portal iframe 就绪，再多策略尝试点入口。"""
    log("[STEP] 等待 portal iframe 就绪 (10s)...")
    page.wait_for_timeout(10000)
    ok, where = try_click_entry(page, log)
    if ok:
        log(f"  [AUTO-CLICK OK] via {where}")
        return True
    log("  [WARN] 未能自动点开入口，请手动点开「豌豆币添加申请」")
    return False


def find_form_page(context, log, timeout_s: int = 300):
    log(f"[STEP] 等待表单 tab 出现，最多 {timeout_s}s")
    deadline = time.time() + timeout_s
    while time.time() < deadline:
        for pg in context.pages:
            try:
                u = pg.url
            except Exception:
                continue
            if "kmReviewMain.do" in u and "method=add" in u:
                log(f"  [HIT] 表单 tab: {u[:90]}")
                return pg
        time.sleep(1)
    return None


def main() -> None:
    ap = argparse.ArgumentParser()
    ap.add_argument("--keyword", default="积分成本", help="终端高亮匹配的关键字（默认 积分成本）")
    args = ap.parse_args()
    keyword = args.keyword

    if not AUTH_STATE.exists():
        print(f"[ERR] 缺少登录态: {AUTH_STATE}")
        sys.exit(2)

    # 清空旧文件
    EVENTS_NDJSON.write_text("", encoding="utf-8")
    if STOP_FLAG.exists():
        STOP_FLAG.unlink()

    def log(msg: str) -> None:
        print(msg, flush=True)

    log(f"[INFO] 录制启动，关键字 = {keyword!r}")
    log(f"[INFO] 实时事件流 -> {EVENTS_NDJSON}")
    log(f"[INFO] 操作完成后按 Ctrl+C 退出，会再 dump 一份 -> {EVENTS_JSON}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", args=["--start-maximized"])
        context = browser.new_context(storage_state=str(AUTH_STATE), no_viewport=True)
        page = context.new_page()
        log(f"[STEP] 打开 portal: {PORTAL_URL}")
        page.goto(PORTAL_URL, wait_until="domcontentloaded")
        page.wait_for_timeout(2000)

        # STEP 2: 如果停在系统选择页，点 OA系统 瓷砖
        try:
            if page.locator("text=OA系统").count() > 0:
                log("[STEP 2] 点击 OA系统 瓷砖")
                try:
                    with context.expect_page(timeout=15000) as new_page_info:
                        page.locator("text=OA系统").first.click()
                    oa_page = new_page_info.value
                except Exception:
                    oa_page = page
                try:
                    oa_page.wait_for_load_state("domcontentloaded", timeout=30000)
                except Exception:
                    pass
                oa_page.wait_for_timeout(3000)
            else:
                log("[STEP 2] 未见系统选择页，认为已在 OA 内")
                oa_page = page
        except Exception as e:
            log(f"[STEP 2] 异常 {e}，继续")
            oa_page = page

        log(f"  [OA URL] {oa_page.url}")
        click_entry(oa_page, log)

        # 不再依赖 URL 匹配 — 给 context 里所有 page 都注入 recorder，
        # 新 page 出现时自动注入；新 frame 出现时自动注入。
        log("[STEP] 等表单 tab/iframe 出现并注入 recorder（最多 60s）...")
        # 给已有 pages 都注入一次
        for pg in list(context.pages):
            try:
                pg.wait_for_load_state("domcontentloaded", timeout=3000)
            except Exception:
                pass
            install_in_page(pg, log)

        # context 级别监听：新 page / 新 frame
        def on_new_page(pg):
            log(f"[NEW-PAGE] {pg.url[:80]}")
            try:
                pg.wait_for_load_state("domcontentloaded", timeout=10000)
            except Exception:
                pass
            install_in_page(pg, log)
            try:
                pg.on("frameattached", lambda fr: (fr.wait_for_load_state("domcontentloaded", timeout=5000) if hasattr(fr, "wait_for_load_state") else None, install_recorder(fr, f"  new-frame {fr.url[:60]}", log)))
            except Exception:
                pass

        context.on("page", on_new_page)
        for pg in list(context.pages):
            try:
                pg.on("frameattached", lambda fr: install_recorder(fr, f"  new-frame {fr.url[:60]}", log))
            except Exception:
                pass

        # 等几秒让表单 tab 起来再做一次注入
        oa_page.wait_for_timeout(8000)
        for pg in list(context.pages):
            install_in_page(pg, log)

        log("=" * 70)
        log("现在请在浏览器里手动操作「积分成本归属部门」字段")
        log(f"高亮规则：事件 target 的 text/value/attrs 任一含 {keyword!r} 即标记 ★")
        log("操作完成后回到这里按 Ctrl+C 结束")
        log("=" * 70)

        all_events: list[dict] = []
        stop = {"flag": False}

        def handle_sigint(signum, frame):  # noqa: ARG001
            stop["flag"] = True

        signal.signal(signal.SIGINT, handle_sigint)

        ndjson_fp = EVENTS_NDJSON.open("a", encoding="utf-8")
        try:
            seq = 0
            last_reinstall = time.time()
            while not stop["flag"]:
                try:
                    chunk = drain_all_pages(context)
                except Exception as e:
                    log(f"[DRAIN-ERR] {e}")
                    chunk = []
                for ev in chunk:
                    seq += 1
                    ev["seq"] = seq
                    ndjson_fp.write(json.dumps(ev, ensure_ascii=False) + "\n")
                    all_events.append(ev)

                    blob = json.dumps(ev, ensure_ascii=False)
                    matched = keyword and (keyword in blob)
                    star = "★" if matched else " "
                    kind = ev.get("kind", "?")
                    tgt = ev.get("target") or {}
                    tag = tgt.get("tag", "")
                    val = tgt.get("value", "")
                    txt = (tgt.get("text") or "").replace("\n", " ")[:60]
                    extra = ev.get("extra") or {}
                    extra_str = ""
                    if kind == "keydown":
                        extra_str = f" key={extra.get('key')}"
                    elif kind in ("input", "change"):
                        extra_str = f" v={(extra.get('value') or '')[:40]!r}"
                    elif kind == "dom-attr":
                        extra_str = f" attr={extra.get('attr')}"
                    if matched or kind in ("click", "mousedown", "change", "keydown"):
                        log(f"{star} #{seq:04d} {kind:10s} <{tag} id={tgt.get('id','')} name={tgt.get('name','')} val={val[:30]!r} text={txt!r}>{extra_str}")
                ndjson_fp.flush()

                # 每 5s 给所有 page/frame 重新注入一次（幂等）
                if time.time() - last_reinstall > 5:
                    for pg in list(context.pages):
                        try:
                            install_in_page(pg, lambda *_: None)
                        except Exception:
                            pass
                    last_reinstall = time.time()
                time.sleep(0.5)
        finally:
            ndjson_fp.close()
            EVENTS_JSON.write_text(json.dumps(all_events, ensure_ascii=False, indent=2), encoding="utf-8")
            log(f"[OK] 共记录 {len(all_events)} 条事件 -> {EVENTS_JSON}")
            try:
                browser.close()
            except Exception:
                pass


if __name__ == "__main__":
    main()
