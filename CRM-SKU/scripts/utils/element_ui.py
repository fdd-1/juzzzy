#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""Element UI 表单填充工具函数"""

import time


def fill_text(page, label_text, value, timeout=10000):
    """填充文本框"""
    form_item = page.locator(f".el-form-item:has-text('{label_text}')").first
    input_field = form_item.locator("input").first
    input_field.fill(str(value), timeout=timeout)
    page.wait_for_timeout(200)


def fill_number(page, label_text, value, timeout=10000):
    """填充数字输入框"""
    fill_text(page, label_text, str(value), timeout)


def fill_filterable_dropdown(page, label_text, value, timeout=10000):
    """填充可搜索下拉框"""
    form_item = page.locator(f".el-form-item:has-text('{label_text}')").first
    input_locator = form_item.locator("input").first

    # 点击触发下拉
    input_locator.click(timeout=timeout)
    page.wait_for_timeout(500)

    # 检查是否 readonly
    is_readonly = input_locator.get_attribute("readonly") is not None

    if not is_readonly:
        # 可编辑，清空并输入
        page.keyboard.press("Control+A")
        page.keyboard.press("Delete")
        page.keyboard.type(value, delay=50)
        page.wait_for_timeout(800)  # 增加等待时间

    # 在下拉面板中选择
    page.wait_for_timeout(500)  # 等待下拉面板出现

    for opt in page.locator(".el-select-dropdown:visible .el-select-dropdown__item").all():
        if value in opt.inner_text().strip():
            opt.click()
            page.wait_for_timeout(300)
            return

    raise ValueError(f"找不到选项: {value}")


def fill_cascade_dropdown(page, label_text, values, timeout=10000):
    """
    填充二级级联下拉（CRM 的课包类型用的是 el-cascader）
    values: [parent, leaf]
    """
    if not isinstance(values, list) or len(values) != 2:
        raise ValueError(f"级联下拉需要 [parent, leaf] 格式，收到: {values}")

    parent, leaf = values
    form_item = page.locator(f".el-form-item:has-text('{label_text}')").first

    # el-cascader 需要点击整个容器，不是 input
    cascader = form_item.locator(".el-cascader").first
    cascader.click(timeout=timeout)
    page.wait_for_timeout(1000)

    # 选择一级
    found_parent = False
    nodes = page.locator(".el-cascader-node").all()

    print(f"  [DEBUG] 找到 {len(nodes)} 个一级节点")

    for node in nodes:
        try:
            if not node.is_visible():
                continue
            text = node.inner_text().strip()
            print(f"  [DEBUG] 一级节点: {text}")
            if parent in text:
                print(f"  [DEBUG] 匹配到一级: {text}")
                node.click()
                page.wait_for_timeout(800)
                found_parent = True
                break
        except Exception as e:
            print(f"  [DEBUG] 一级节点异常: {e}")
            continue

    if not found_parent:
        raise ValueError(f"级联下拉 '{label_text}' 第 1 级找不到选项: {parent}")

    # 选择二级（点击一级后，二级会出现）
    found_leaf = False
    page.wait_for_timeout(500)
    nodes = page.locator(".el-cascader-node").all()

    print(f"  [DEBUG] 找到 {len(nodes)} 个二级节点")

    for node in nodes:
        try:
            if not node.is_visible():
                continue
            text = node.inner_text().strip()
            print(f"  [DEBUG] 二级节点: {text}")
            if leaf in text:
                print(f"  [DEBUG] 匹配到二级: {text}")
                node.click()
                page.wait_for_timeout(600)
                found_leaf = True
                break
        except Exception as e:
            print(f"  [DEBUG] 二级节点异常: {e}")
            continue

    if not found_leaf:
        raise ValueError(f"级联下拉 '{label_text}' 第 2 级找不到选项: {leaf}")


def fill_multi_select(page, label_text, values, timeout=10000):
    """
    填充多选下拉
    values: ["选项1", "选项2", ...]
    """
    form_item = page.locator(f".el-form-item:has-text('{label_text}')").first
    input_locator = form_item.locator("input").first

    # 点击触发下拉
    input_locator.click(timeout=timeout)
    page.wait_for_timeout(300)

    # 逐个选择
    for v in values:
        opt = page.locator(".el-select-dropdown:visible .el-select-dropdown__item").filter(has_text=v).first
        if opt.count() > 0:
            opt.click()
            page.wait_for_timeout(150)
        else:
            print(f"  [WARN] 多选下拉 '{label_text}' 找不到选项: {v}")

    # 按 Escape 收起
    page.keyboard.press("Escape")
    page.wait_for_timeout(200)


def click_button_with_text(page, text, timeout=10000):
    """点击按钮（处理中文 2 字按钮的 letter-spacing）"""
    candidates = [text]
    if len(text) == 2:
        candidates.append(f"{text[0]} {text[1]}")

    for t in candidates:
        for sel in [f"button:has-text('{t}')", f".el-button:has-text('{t}')"]:
            try:
                page.locator(sel).first.click(timeout=timeout)
                return
            except Exception:
                continue

    raise ValueError(f"找不到按钮: {text}")


def read_select_value(page, label_text):
    """读取下拉框当前值"""
    try:
        return page.locator(
            f".el-form-item:has-text('{label_text}') input"
        ).first.input_value() or ""
    except Exception:
        return ""
