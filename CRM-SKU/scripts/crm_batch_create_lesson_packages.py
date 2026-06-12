#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
CRM 课时包批量创建脚本
基于 Playwright 自动化 Element UI 表单
"""

import sys
import io
import json
import time
import csv
import re
from pathlib import Path
from datetime import datetime
import openpyxl
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeout

# Windows 编码处理
if sys.platform == "win32":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

# 导入工具函数
sys.path.insert(0, str(Path(__file__).parent))
from utils.element_ui import (
    fill_text, fill_number, fill_filterable_dropdown,
    fill_cascade_dropdown, fill_multi_select,
    click_button_with_text, read_select_value
)
from utils.auth import get_credentials
from utils.precheck import precheck_lesson_excel, print_report, verify_in_list

# 配置
SCRIPT_DIR = Path(__file__).parent
CONFIG_FILE = SCRIPT_DIR / "config.local.json"
AUTH_STATE_FILE = SCRIPT_DIR / "auth_state.json"
LOG_DIR = SCRIPT_DIR / "logs"
LOG_DIR.mkdir(exist_ok=True)

# Excel 字段映射
HEADER_MAP = {
    "课时包名称": "name",
    "课包类型": "package_type",
    "课包分类": "package_category",
    "有效期": "valid_months",
    "补课次数": "makeup_chances",
    "普通课时": "normal_lessons",
    "赠送课时": "gift_lessons",
    "原价": "original_price",
    "优惠价": "discount_price",
    "试学期": "trial_days",
    "打卡次数": "checkin_count",
    "停课次数": "suspend_count",
    "适用课类": "applicable_classes",
    "赠送礼品": "gift_items",
}

# 课包类型父级映射
TYPE_PARENT = {
    "中课包": "常规正课",
    "年课包": "常规正课",
    "年课包pro": "常规正课",
    "季课包": "常规正课",
    "两年课包": "常规正课",
    "其他类型课包": "常规正课",
    "短期包": "常规正课",
}

# Excel 名称 -> CRM 名称（执行前校验也用这张表）
TYPE_NAME_MAPPING = {
    "年课包pro": "年课包",
    "两年包": "两年课包",
    "其他": "其他类型课包",
}


def load_config():
    """加载配置"""
    if not CONFIG_FILE.exists():
        raise FileNotFoundError(f"配置文件不存在: {CONFIG_FILE}")
    with open(CONFIG_FILE, encoding="utf-8") as f:
        return json.load(f)


def parse_excel(excel_path):
    """解析 Excel 配置（横向 KV 结构）"""
    wb = openpyxl.load_workbook(excel_path)
    ws = wb.active

    rows = []
    for raw in ws.iter_rows(values_only=True):
        if not raw or raw[0] is None:
            continue
        if str(raw[0]).strip() != "课时包名称":
            continue

        pkg = {}
        for i in range(0, len(raw), 2):
            label = raw[i]
            value = raw[i + 1] if i + 1 < len(raw) else None
            if label is None:
                break
            field = HEADER_MAP.get(str(label).strip())
            if field:
                if field == "package_type":
                    # 补充父级
                    leaf = str(value).strip() if value else ""
                    leaf = TYPE_NAME_MAPPING.get(leaf, leaf)
                    parent = TYPE_PARENT.get(leaf, "常规正课")
                    pkg[field] = [parent, leaf]
                elif field == "applicable_classes":
                    # 分割适用课类
                    if value:
                        classes = re.split(r'[、，,]', str(value))
                        pkg[field] = [c.strip() for c in classes if c.strip()]
                else:
                    pkg[field] = value

        if pkg:
            rows.append(pkg)

    return rows


def collect_messages(page):
    """收集所有可见的 toast/错误消息"""
    messages = []
    for sel in [".el-message", ".el-notification", ".el-form-item__error", ".el-message-box"]:
        for el in page.locator(sel).all():
            try:
                if el.is_visible():
                    txt = (el.inner_text() or "").strip()
                    if txt:
                        messages.append(f"{sel}: {txt}")
            except Exception:
                pass
    return messages


def submit_and_verify(page, timeout=15000):
    """提交表单并验证结果"""
    # 点击确定按钮
    try:
        click_button_with_text(page, "确定")
    except Exception as e:
        return False, f"点击确定按钮失败: {e}"

    # 轮询收集消息（15秒）
    captured = []
    deadline = page.evaluate("Date.now()") + timeout
    success = False

    while page.evaluate("Date.now()") < deadline:
        for m in collect_messages(page):
            if m not in captured:
                captured.append(m)
                print(f"  [MSG] {m}")

        # 检查成功关键词
        success_keywords = [
            "保存成功", "创建成功", "操作成功", "新增成功", "添加成功",
            ".el-message: OK", ".el-message: ok",
            ".el-message: Success", ".el-message: success",
        ]
        if any(kw in m for m in captured for kw in success_keywords):
            success = True
            break

        # 检查失败关键词
        fail_keywords = ["失败", "错误", "不能为空", "必填", "已存在"]
        if any(kw in m for m in captured for kw in fail_keywords):
            break

        page.wait_for_timeout(200)

    detail = " | ".join(captured) if captured else "（未抓到 toast）"
    return success, detail


def close_dialog(page):
    """关闭弹窗（处理脏状态）"""
    if page.locator(".el-dialog:visible").count() > 0:
        try:
            click_button_with_text(page, "取消")
        except Exception:
            page.keyboard.press("Escape")
        page.wait_for_timeout(300)

        if page.locator(".el-dialog:visible").count() > 0:
            try:
                click_button_with_text(page, "确定")
            except Exception:
                pass
            page.wait_for_timeout(300)


def is_name_exists(page, name):
    """检查课时包名称是否已存在"""
    try:
        # 在搜索框输入名称
        search_input = page.locator("input[placeholder*='课时包']").first
        search_input.fill(name)
        page.wait_for_timeout(300)

        # 点击查询按钮
        click_button_with_text(page, "查询")
        page.wait_for_timeout(800)

        # 检查表格中是否有该名称
        exists = page.locator(f".el-table tr").filter(has_text=name).first.is_visible()
        return exists
    except Exception:
        return False


def fill_form(page, data):
    """填充课时包表单"""
    print(f"  填充表单: {data.get('name', 'Unknown')}")

    # 1. 课时包名称
    fill_text(page, "课时包名称", data["name"])

    # 2. 课包类型（二级级联）
    fill_cascade_dropdown(page, "课包类型", data["package_type"])
    page.wait_for_timeout(500)

    # 3. 课包分类
    fill_filterable_dropdown(page, "课包分类", data["package_category"])
    page.wait_for_timeout(500)

    # 校验课包类型是否被清空（已知坑）
    expected_leaf = data["package_type"][-1]
    if expected_leaf not in read_select_value(page, "课包类型"):
        print(f"  [WARN] 课包类型被清空，重新选择")
        fill_cascade_dropdown(page, "课包类型", data["package_type"])
        page.wait_for_timeout(500)

    # 4. 数字字段
    fill_number(page, "有效期", data["valid_months"])
    fill_number(page, "补课次数", data["makeup_chances"])
    fill_number(page, "普通课时", data["normal_lessons"])
    fill_number(page, "赠送课时", data["gift_lessons"])
    fill_number(page, "原价", data["original_price"])
    fill_number(page, "优惠价", data["discount_price"])
    fill_number(page, "试学期", data["trial_days"])

    # 5. 动态字段（选完类型后才出现）
    if "checkin_count" in data:
        fill_number(page, "打卡次数", data["checkin_count"])
    if "suspend_count" in data:
        fill_number(page, "停课次数", data["suspend_count"])

    # 6. 适用课类（多选）
    if "applicable_classes" in data and data["applicable_classes"]:
        fill_multi_select(page, "适用课类", data["applicable_classes"])


def create_one_package(page, data, skip_existing=False):
    """创建一个课时包"""
    name = data.get("name", "Unknown")
    print(f"\n处理: {name}")

    try:
        # 检查是否已存在
        if skip_existing and is_name_exists(page, name):
            print(f"  ✓ 已存在，跳过")
            return "SKIP", "已存在"

        # 等待页面加载完成（避免使用 networkidle，参考 crm-lesson-package 已知坑）
        page.wait_for_timeout(2000)

        # 点击添加课时包按钮（多种选择器尝试）
        button_clicked = False
        for selector in [
            "button:has-text('添加课时包')",
            ".el-button:has-text('添加课时包')",
            "button:has-text('添加')",
            ".el-button--primary:has-text('添加')"
        ]:
            try:
                page.locator(selector).first.click(timeout=5000)
                button_clicked = True
                print(f"  ✓ 点击按钮成功: {selector}")
                break
            except Exception:
                continue

        if not button_clicked:
            raise ValueError("找不到添加课时包按钮")

        page.wait_for_timeout(1500)  # 增加等待时间

        # 填充表单
        fill_form(page, data)

        # 提交并验证
        success, detail = submit_and_verify(page)

        if success:
            print(f"  ✓ 创建成功")
            return "OK", detail
        else:
            print(f"  ✗ 创建失败: {detail}")
            # 截图
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            safe_name = re.sub(r'[^\w\-]', '_', name)[:50]
            screenshot_path = LOG_DIR / f"submit-fail-{safe_name}-{timestamp}.png"
            page.screenshot(path=str(screenshot_path))
            return "FAIL", detail

    except Exception as e:
        print(f"  ✗ 异常: {e}")
        return "FAIL", str(e)
    finally:
        # 清理弹窗
        close_dialog(page)


def batch_create(excel_path, start=1, limit=None, skip_existing=False, use_password=False,
                 strict_precheck=True):
    """批量创建课时包"""
    # 执行前校验
    ok, errors, warnings = precheck_lesson_excel(excel_path, TYPE_NAME_MAPPING, TYPE_PARENT)
    print_report("执行前校验", ok, errors, warnings)
    if not ok:
        if strict_precheck:
            print("校验未通过，已中止。修复后重跑，或加 --no-precheck 跳过（不推荐）。")
            return
        print("校验未通过但已选择跳过，继续执行。")

    print(f"解析 Excel: {excel_path}")
    packages = parse_excel(excel_path)
    print(f"共找到 {len(packages)} 条配置\n")

    if not packages:
        print("没有找到有效的课时包配置")
        return

    # 加载配置
    config = load_config()
    login_url = config["crm"]["login_url"]
    class_package_url = config["crm"]["class_package_url"]

    # 启动浏览器
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False)

        # 如果使用密码登录，不加载旧的登录态
        if use_password:
            context = browser.new_context()
            page = context.new_page()

            print("使用账号密码登录...")
            page.goto(login_url, timeout=30000)
            page.wait_for_timeout(2000)

            try:
                # 凭据：env > config.username > 交互式输入（密码不再从 config 读）
                username, password = get_credentials(config.get("auth"))

                page.locator("input[placeholder*='手机号'], input[placeholder*='账号']").first.fill(username)
                page.wait_for_timeout(500)
                page.locator("input[type='password'], input[placeholder*='密码']").first.fill(password)
                page.wait_for_timeout(500)
                page.locator("button:has-text('登录'), button:has-text('登 录')").first.click()
                page.wait_for_timeout(8000)

                # 等待登录跳转完成
                try:
                    page.wait_for_url(lambda url: "login" not in url.lower(), timeout=10000)
                except Exception:
                    pass

                print("✓ 登录完成\n")

                # 保存登录态
                context.storage_state(path=str(AUTH_STATE_FILE))
            except Exception as e:
                print(f"密码登录失败: {e}")
                return
        else:
            # 加载登录态
            if AUTH_STATE_FILE.exists():
                context = browser.new_context(storage_state=str(AUTH_STATE_FILE))
            else:
                context = browser.new_context()
            page = context.new_page()

        page.goto(class_package_url, timeout=30000)
        page.wait_for_timeout(3000)

        # 检查是否需要登录（多种检测方式）
        needs_login = False

        # 方式1：检查 URL
        if "login" in page.url.lower() or "passport" in page.url.lower():
            needs_login = True

        # 方式2：检查页面内容
        if not needs_login:
            try:
                # 如果能找到二维码，说明需要登录
                if page.locator(".qrcode, [class*='qr'], [class*='QR']").count() > 0:
                    needs_login = True
            except Exception:
                pass

        # 方式3：检查是否有"添加课时包"按钮
        if not needs_login:
            try:
                if page.locator("button:has-text('添加课时包')").count() == 0:
                    needs_login = True
            except Exception:
                pass

        if needs_login and not use_password:
            print("检测到需要登录，请在浏览器中扫码登录...")
            print("等待登录完成...")

            # 等待跳转到课时包管理页面
            try:
                page.wait_for_url("**/ClassPackageManage", timeout=120000)
            except Exception:
                # 如果 URL 没变，等待"添加课时包"按钮出现
                page.wait_for_selector("button:has-text('添加课时包')", timeout=120000)

            # 保存登录态
            context.storage_state(path=str(AUTH_STATE_FILE))
            print("✓ 登录成功，已保存登录态\n")
            page.wait_for_timeout(3000)

        # 准备日志
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        log_file = LOG_DIR / f"batch-{timestamp}.csv"

        with open(log_file, "w", encoding="utf-8-sig", newline="") as f:
            writer = csv.writer(f)
            writer.writerow(["序号", "名称", "结果", "详情", "搜索校验"])

            # 处理每条配置
            for i, pkg in enumerate(packages, 1):
                if i < start:
                    continue
                if limit and i >= start + limit:
                    break

                # 刷新列表页
                page.goto(class_package_url)
                page.wait_for_timeout(2000)

                # 创建课时包
                result, detail = create_one_package(page, pkg, skip_existing)

                # 执行后验证：列表页搜索确认课包已创建
                verify_detail = ""
                if result == "OK":
                    page.goto(class_package_url)
                    page.wait_for_timeout(1500)
                    found, verify_detail = verify_in_list(
                        page, pkg.get("name", ""),
                        search_input_selector="input[placeholder*='课时包']",
                    )
                    if not found:
                        result = "OK_BUT_NOT_FOUND"
                        print(f"  [WARN] 创建提示成功但列表搜不到：{pkg.get('name')} ({verify_detail})")

                # 记录日志
                writer.writerow([i, pkg.get("name", "Unknown"), result, detail, verify_detail])
                f.flush()

        print(f"\n批量创建完成，日志: {log_file}")

        # 保持浏览器打开（仅在交互式环境）
        try:
            input("\n按 Enter 关闭浏览器...")
        except EOFError:
            print("非交互式环境，5秒后自动关闭浏览器...")
            page.wait_for_timeout(5000)
        browser.close()


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="CRM 课时包批量创建")
    parser.add_argument("--xlsx", required=True, help="Excel 配置文件路径")
    parser.add_argument("--start", type=int, default=1, help="从第 N 条开始（1-based）")
    parser.add_argument("--limit", type=int, help="最多处理 N 条")
    parser.add_argument("--skip-existing", action="store_true", help="跳过已存在的课时包")
    parser.add_argument("--use-password", action="store_true", help="使用账号密码登录")
    parser.add_argument("--no-precheck", action="store_true",
                        help="跳过执行前校验（不推荐，仅在确认无误时使用）")

    args = parser.parse_args()

    batch_create(args.xlsx, args.start, args.limit, args.skip_existing, args.use_password,
                 strict_precheck=not args.no_precheck)
