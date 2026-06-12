#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""平台中心 → 用户管理 → 批量用户查询：学员ID → 豌豆大账号

流程：
  1) 复用六一登录态进 portal
  2) 点「平台中心」瓷砖（新 tab）
  3) 左侧菜单展开「用户管理」→ 点「批量用户查询」
  4) 点「新建」
  5) 弹窗：
     - 导入类型：下拉选「豌豆用户」
     - 上传文件（upload_*.xlsx）
     - 导出类型：勾选「豌豆大账号」
     - 点「保存」
  6) 弹窗关闭后等 10s，刷新页面
  7) 通过 page.request POST queryExportRecords 拿列表，找到刚提交那条
     （以 importFileName 完全匹配 + 自己 optUserName 来锚定）
  8) 轮询 handleStatus=1（处理成功），最多 2 分钟
  9) 下载 exportFile 的 OSS xlsx → 保存到 dadou_mapping_*.xlsx
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
ROOT = SCRIPT_DIR.parent
AUTH_PATH = ROOT / "liuyi_login" / "auth_state.json"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"
QUERY_API = "https://gw-mg.61info.cn/bizcenter-usercenter/p/v1/user/queryExportRecords"


def log(msg):
    print(msg, flush=True)


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--upload", help="上传 xlsx 路径，默认 pingtai_query/upload_{today}.xlsx")
    ap.add_argument("--import-type", default="豌豆用户",
                    help="导入类型下拉选项（默认 豌豆用户）")
    ap.add_argument("--export-type", default="豌豆大账号",
                    help="导出类型勾选项（默认 豌豆大账号）")
    ap.add_argument("--poll-timeout", type=int, default=120,
                    help="轮询处理结果的最大秒数（默认 120）")
    args = ap.parse_args()

    today = dt.date.today().strftime("%Y%m%d")
    if args.upload:
        upload_path = Path(args.upload)
    else:
        # 找当天最新的 upload_*.xlsx
        candidates = sorted(SCRIPT_DIR.glob(f"upload_{today}*.xlsx"), key=lambda p: p.stat().st_mtime)
        if not candidates:
            log(f"[ERROR] 找不到 upload_{today}*.xlsx，请先跑 prepare_template.py")
            sys.exit(1)
        upload_path = candidates[-1]
    if not upload_path.exists():
        log(f"[ERROR] 找不到上传文件: {upload_path}")
        sys.exit(1)
    if not AUTH_PATH.exists():
        log(f"[ERROR] 找不到登录态: {AUTH_PATH}，先跑 liuyi_login/login_liuyi.py")
        sys.exit(1)

    upload_filename = upload_path.name
    log(f"[INFO] 上传文件: {upload_path} (filename={upload_filename})")
    log(f"[INFO] 导入类型: {args.import_type}, 导出类型: {args.export_type}")

    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, channel="chrome", slow_mo=80)
        ctx = browser.new_context(storage_state=str(AUTH_PATH))
        page = ctx.new_page()

        # ── Step 1: portal → 平台中心
        log(f"[STEP 1] 打开 portal: {PORTAL_URL}")
        page.goto(PORTAL_URL, timeout=30000)
        page.locator("p:has-text('平台中心')").first.wait_for(timeout=30000)
        log("[STEP 2] 点击「平台中心」瓷砖")
        with ctx.expect_page(timeout=30000) as new_page_info:
            page.locator("p:has-text('平台中心')").first.click()
        pt = new_page_info.value
        pt.wait_for_load_state("domcontentloaded", timeout=30000)
        pt.wait_for_timeout(3000)
        log(f"  [PT URL] {pt.url}")
        pt.screenshot(path=str(SCRIPT_DIR / "step1_pt_landing.png"))

        # ── Step 3: 左侧菜单 用户管理 → 批量用户查询
        log("[STEP 3] 左侧菜单进入「批量用户查询」")
        # 用户管理可能要先点开
        try:
            user_mgmt = pt.locator("text=用户管理").first
            user_mgmt.click(timeout=10000)
            pt.wait_for_timeout(800)
        except Exception as e:
            log(f"[WARN] 点击「用户管理」未成功（可能已展开）: {e}")
        try:
            batch_query = pt.locator("text=批量用户查询").first
            batch_query.click(timeout=10000)
        except Exception as e:
            log(f"[ERROR] 点击「批量用户查询」失败: {e}")
            pt.screenshot(path=str(SCRIPT_DIR / "err_no_batch_menu.png"))
            sys.exit(2)
        pt.wait_for_timeout(2000)
        pt.screenshot(path=str(SCRIPT_DIR / "step2_batch_query.png"))

        # ── Step 4: 点「新建」
        log("[STEP 4] 点击「新建」")
        try:
            pt.locator("button:has-text('新建')").first.click(timeout=10000)
        except Exception as e:
            log(f"[ERROR] 没找到「新建」按钮: {e}")
            pt.screenshot(path=str(SCRIPT_DIR / "err_no_new_btn.png"))
            sys.exit(3)
        pt.wait_for_timeout(1500)
        pt.screenshot(path=str(SCRIPT_DIR / "step3_modal_open.png"))
        # dump 弹窗 HTML 便于排查
        try:
            modal_html = pt.locator(".el-dialog").first.inner_html(timeout=5000)
            (SCRIPT_DIR / "modal_dump.html").write_text(modal_html, encoding="utf-8")
        except Exception:
            pass

        dialog = pt.locator(".el-dialog").first

        # ── Step 5: 填弹窗
        # 5.1 导入类型下拉（弹窗内第 1 个 .el-select 就是导入类型）
        log(f"[STEP 5.1] 导入类型 → 选「{args.import_type}」")
        try:
            import_select = dialog.locator(".el-select").first
            import_select.click(timeout=10000)
            pt.wait_for_timeout(800)
            # 选项浮层挂在 body 下，不在 dialog 内
            option = pt.locator(
                f".el-select-dropdown li:has-text('{args.import_type}')"
            ).first
            option.click(timeout=8000)
            pt.wait_for_timeout(500)
        except Exception as e:
            log(f"[ERROR] 导入类型下拉失败: {e}")
            pt.screenshot(path=str(SCRIPT_DIR / "err_import_type.png"))
            sys.exit(4)

        # 5.2 上传文件
        log(f"[STEP 5.2] 上传文件: {upload_path}")
        try:
            file_input = dialog.locator("input[type='file']").first
            file_input.set_input_files(str(upload_path), timeout=10000)
            pt.wait_for_timeout(1500)
        except Exception as e:
            log(f"[ERROR] 上传文件失败: {e}")
            pt.screenshot(path=str(SCRIPT_DIR / "err_upload.png"))
            sys.exit(5)

        # 5.3 导出类型勾选（弹窗内 .el-checkbox，文字匹配）
        log(f"[STEP 5.3] 导出类型 → 勾选「{args.export_type}」")
        try:
            checkbox = dialog.locator(
                f".el-checkbox:has-text('{args.export_type}')"
            ).first
            checkbox.click(timeout=10000)
            pt.wait_for_timeout(500)
        except Exception as e:
            log(f"[ERROR] 勾选导出类型失败: {e}")
            pt.screenshot(path=str(SCRIPT_DIR / "err_export_type.png"))
            sys.exit(6)

        pt.screenshot(path=str(SCRIPT_DIR / "step4_modal_filled.png"))

        # 5.4 保存
        log("[STEP 5.4] 点击「保存」")
        try:
            dialog.locator("button:has-text('保存')").first.click(timeout=10000)
        except Exception as e:
            log(f"[ERROR] 点击保存失败: {e}")
            pt.screenshot(path=str(SCRIPT_DIR / "err_save.png"))
            sys.exit(7)
        save_clicked_at = time.time()
        pt.wait_for_timeout(3000)
        pt.screenshot(path=str(SCRIPT_DIR / "step5_after_save.png"))

        # ── Step 6: 等 10s + 调接口轮询
        log("[STEP 6] 等 10s 后开始轮询 queryExportRecords 找刚提交那条")
        pt.wait_for_timeout(10000)

        # 用浏览器 fetch 调接口（自动从 localStorage/sessionStorage 找 JWT token 加 Authorization header）
        def call_query_api():
            return pt.evaluate("""
                async () => {
                  let token = '';
                  for (let i = 0; i < localStorage.length; i++) {
                    const k = localStorage.key(i);
                    const v = localStorage.getItem(k);
                    if (v && typeof v === 'string' && v.startsWith('eyJ')) { token = v; break; }
                  }
                  if (!token) {
                    for (let i = 0; i < sessionStorage.length; i++) {
                      const k = sessionStorage.key(i);
                      const v = sessionStorage.getItem(k);
                      if (v && typeof v === 'string' && v.startsWith('eyJ')) { token = v; break; }
                    }
                  }
                  try {
                    const r = await fetch('https://gw-mg.61info.cn/bizcenter-usercenter/p/v1/user/queryExportRecords', {
                      method: 'POST',
                      headers: {
                        'content-type': 'application/json',
                        'authorization': token
                      },
                      body: JSON.stringify({pageSize: 20, pageNum: 1}),
                      credentials: 'include'
                    });
                    const body = await r.json();
                    return { status: r.status, body, tokenFound: !!token };
                  } catch (e) {
                    return { status: -1, error: String(e), tokenFound: !!token };
                  }
                }
            """)

        target_record = None
        deadline = time.time() + args.poll_timeout
        attempt = 0
        while time.time() < deadline:
            attempt += 1
            try:
                result = call_query_api()
            except Exception as e:
                log(f"[POLL #{attempt}] evaluate 异常: {e}")
                time.sleep(3)
                continue

            body = result.get("body") or {}
            if attempt == 1:
                log(f"[POLL] tokenFound={result.get('tokenFound')}, status={result.get('status')}")

            if body.get("code") != 0:
                log(f"[POLL #{attempt}] 接口异常: {body}")
                time.sleep(3)
                continue

            items = body.get("data", {}).get("list", [])
            for it in items:
                if it.get("importFileName") == upload_filename:
                    target_record = it
                    break
            if target_record:
                status = target_record.get("handleStatus")
                rid = target_record.get("id")
                log(f"[POLL #{attempt}] 找到记录 id={rid}, handleStatus={status}, "
                    f"importType={target_record.get('importType')}, "
                    f"exportTypes={target_record.get('exportTypes')}")
                if status == 1:
                    break
                if status == 2:
                    log(f"[ERROR] 记录处理失败: {target_record}")
                    sys.exit(8)
            else:
                log(f"[POLL #{attempt}] 列表里还没找到 {upload_filename}，等 3s")
            time.sleep(3)

        if not target_record:
            log(f"[ERROR] {args.poll_timeout}s 内没在列表里找到刚提交的记录")
            sys.exit(9)
        if target_record.get("handleStatus") != 1:
            log(f"[ERROR] 记录最终状态非「处理成功」: {target_record}")
            sys.exit(10)

        # ── Step 7: 下载 exportFile（OSS 公开 URL，直接 GET 不需要 token）
        export_url = target_record.get("exportFile")
        log(f"[STEP 7] 下载导出文件: {export_url}")
        if not export_url:
            log("[ERROR] 记录没有 exportFile 字段")
            sys.exit(11)
        out = ROOT / "output" / f"dadou_mapping_{today}.xlsx"
        out.parent.mkdir(exist_ok=True)
        resp = pt.request.get(export_url)
        if resp.status != 200:
            log(f"[ERROR] 下载失败: {resp.status}")
            sys.exit(12)
        out.write_bytes(resp.body())
        log(f"[OK] 已下载 → {out} ({out.stat().st_size/1024:.1f} KB)")

        # 保存最新登录态（含平台中心 cookie）
        ctx.storage_state(path=str(AUTH_PATH))
        log(f"[OK] 登录态已更新: {AUTH_PATH}")

        # 输出关键信息
        log("=" * 60)
        log("[DONE] 学员ID → 豌豆大账号 映射已生成")
        log(f"  身份证号: {target_record.get('id')}")
        log(f"  导入文件: {target_record.get('importFileName')}")
        log(f"  导入类型: {target_record.get('importType')}")
        log(f"  导出类型: {target_record.get('exportTypes')}")
        log(f"  导入文件 URL: {target_record.get('importFile')}")
        log(f"  导出文件 URL: {export_url}")
        log(f"  本地输出: {out}")
        log("=" * 60)

        pt.wait_for_timeout(5000)
        browser.close()


if __name__ == "__main__":
    main()
