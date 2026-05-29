#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""平台中心 批量用户查询自动化：学员 ID → 豌豆大账号 ID

输入：filtered_*.xlsx（包含学员 ID 列）
输出：dadou_mapping_*.xlsx（学员 ID + 豌豆大账号 ID 映射表）

流程：
  1. 读 filtered_*.xlsx 的学员 ID 列
  2. 填充到模板 导入用户id.xlsx 的 A 列
  3. 复用登录态进平台中心 → 用户管理 → 批量用户查询
  4. 点「新建」→ 选「导入类型=豌豆用户」「导出类型=豌豆大账号」→ 上传文件 → 保存
  5. 轮询列表页，等「处理状态=处理成功」
  6. 点「导出」下载结果 → 保存到 output/dadou_mapping_*.xlsx
"""
import sys, io, time, shutil
from pathlib import Path
from playwright.sync_api import sync_playwright
import pandas as pd
import datetime as dt

if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")
    sys.stdout.reconfigure(line_buffering=True)
    sys.stderr.reconfigure(line_buffering=True)

SCRIPT_DIR = Path(__file__).parent
ROOT = SCRIPT_DIR.parent
AUTH_PATH = ROOT / "liuyi_login" / "auth_state.json"
OUTPUT_DIR = ROOT / "output"
TEMPLATE = ROOT / "导入用户id.xlsx"
PORTAL_URL = "https://dingding.61info.cn/sys/portal/page.jsp"

# 命令行参数：输入的 filtered xlsx
import argparse
p = argparse.ArgumentParser()
p.add_argument("--input", help="输入 filtered_*.xlsx 路径")
args = p.parse_args()

if not args.input:
    # 默认找 output/filtered_{today}.xlsx
    today_tag = dt.date.today().strftime("%Y%m%d")
    args.input = str(OUTPUT_DIR / f"filtered_{today_tag}.xlsx")

input_path = Path(args.input)
if not input_path.exists():
    print(f"[ERROR] 找不到输入文件: {input_path}", flush=True)
    sys.exit(2)

if not AUTH_PATH.exists():
    print(f"[ERROR] 找不到登录态: {AUTH_PATH}，请先跑 liuyi_login/login_liuyi.py", flush=True)
    sys.exit(2)

if not TEMPLATE.exists():
    print(f"[ERROR] 找不到模板: {TEMPLATE}", flush=True)
    sys.exit(2)

print(f"[STEP 1] 读取 {input_path}，提取学员 ID 列（第一列）", flush=True)
df_in = pd.read_excel(str(input_path))
student_ids = df_in.iloc[:, 0].dropna().astype(int).tolist()
print(f"  提取到 {len(student_ids)} 个学员 ID", flush=True)

print(f"[STEP 2] 复制模板 {TEMPLATE.name}，填充学员 ID", flush=True)
timestamp = dt.datetime.now().strftime("%Y%m%d_%H%M%S")
upload_file = SCRIPT_DIR / f"导入用户id_{timestamp}.xlsx"
shutil.copy2(str(TEMPLATE), str(upload_file))
df_tpl = pd.DataFrame({"导入用户id": student_ids})
df_tpl.to_excel(str(upload_file), index=False)
print(f"  已生成上传文件: {upload_file} ({len(student_ids)} 条)", flush=True)

print(f"[STEP 3] 启动浏览器，进平台中心", flush=True)
with sync_playwright() as pw:
    browser = pw.chromium.launch(headless=False, channel="chrome", slow_mo=100)
    ctx = browser.new_context(storage_state=str(AUTH_PATH))
    page = ctx.new_page()

    page.goto(PORTAL_URL, timeout=30000)
    page.wait_for_load_state("domcontentloaded")
    print("  等 portal 瓷砖渲染", flush=True)
    page.locator("p:has-text('平台中心')").first.wait_for(timeout=30000)

    print("  点击「平台中心」瓷砖", flush=True)
    with ctx.expect_page(timeout=30000) as new_page_info:
        page.locator("p:has-text('平台中心')").first.click()
    pt = new_page_info.value

    pt.wait_for_load_state("domcontentloaded", timeout=30000)
    pt.wait_for_timeout(3000)
    print(f"  平台中心 URL: {pt.url}", flush=True)

    print("[STEP 4] 展开「用户管理」→ 点「批量用户查询」", flush=True)
    pt.locator("text=用户管理").first.click()
    pt.wait_for_timeout(1000)
    pt.locator("text=批量用户查询").first.click()
    pt.wait_for_timeout(2000)

    print("[STEP 5] 点右上「新建」", flush=True)
    pt.locator("button:has-text('新建')").click()
    pt.wait_for_timeout(1500)

    print("[STEP 6] 填写弹窗表单", flush=True)
    # 导入类型下拉
    print("  选择「导入类型=豌豆用户」", flush=True)
    pt.locator("text=请选择导入类型").click()
    pt.wait_for_timeout(500)
    pt.locator("li:has-text('豌豆用户')").click()
    pt.wait_for_timeout(500)

    # 导出类型多选框
    print("  勾选「导出类型=豌豆大账号」", flush=True)
    pt.locator("label:has-text('豌豆大账号')").locator("input[type='checkbox']").check()
    pt.wait_for_timeout(500)

    # 上传文件
    print(f"  上传文件: {upload_file.name}", flush=True)
    with pt.expect_file_chooser() as fc_info:
        pt.locator("button:has-text('上传')").click()
    fc = fc_info.value
    fc.set_files(str(upload_file))
    pt.wait_for_timeout(2000)

    # 保存
    print("  点击「保存」", flush=True)
    pt.locator("button:has-text('保存')").click()
    pt.wait_for_timeout(3000)

    print("[STEP 7] 轮询列表页，等待「处理成功」（最多 5 分钟）", flush=True)
    deadline = time.time() + 300
    task_id = None
    while time.time() < deadline:
        # 刷新列表
        pt.reload()
        pt.wait_for_timeout(3000)

        # 找到刚才上传的文件名对应的行
        rows = pt.locator("table tbody tr").all()
        for row in rows:
            cells = row.locator("td").all()
            if len(cells) < 8:
                continue
            file_col = cells[2].inner_text()
            status_col = cells[6].inner_text()
            if upload_file.stem in file_col:
                print(f"  找到任务: {file_col} | 状态={status_col}", flush=True)
                if "处理成功" in status_col:
                    task_id = cells[0].inner_text()
                    print(f"  任务已完成，身份证={task_id}", flush=True)
                    # 点导出
                    export_btn = cells[7].locator("button:has-text('导出')")
                    with pt.expect_download() as dl_info:
                        export_btn.click()
                    dl = dl_info.value
                    result_file = OUTPUT_DIR / f"dadou_mapping_{dt.date.today().strftime('%Y%m%d')}.xlsx"
                    dl.save_as(str(result_file))
                    print(f"[OK] 导出文件已保存: {result_file}", flush=True)
                    browser.close()
                    sys.exit(0)
                else:
                    print(f"  状态尚未完成，10s 后重试", flush=True)
                    break
        time.sleep(10)

    print("[ERROR] 5 分钟内未等到「处理成功」", flush=True)
    pt.screenshot(path=str(SCRIPT_DIR / "timeout.png"), full_page=True)
    browser.close()
    sys.exit(4)
