# -*- coding: utf-8 -*-
"""SKU复盘自动化 - 分析模块"""

import sys
from collections import defaultdict

sys.stdout.reconfigure(encoding='utf-8')

import openpyxl
from extract_data import find_header_row, get_headers


def aggregate_by_node(matched_orders):
    """按【人群×SKU节点】聚合（节点=升舱/早鸟/其余，来自正式池池子节点2）
    返回包含'综合'汇总（同节点跨人群+同人群跨节点+全量综合）"""
    agg = defaultdict(lambda: {
        "count": 0, "total_amt": 0,
        "total_hrs_no": 0, "total_hrs_w": 0
    })

    for order in matched_orders:
        node = order.get("sku_node", "其余")
        cohort = order["cohort"]
        # (cohort, node)
        agg[(cohort, node)]["count"] += 1
        agg[(cohort, node)]["total_amt"] += order["amount"]
        agg[(cohort, node)]["total_hrs_no"] += order["hours_no_integ"]
        agg[(cohort, node)]["total_hrs_w"] += order["hours_with_integ"]
        # (综合, node) 跨人群同节点
        agg[("综合", node)]["count"] += 1
        agg[("综合", node)]["total_amt"] += order["amount"]
        agg[("综合", node)]["total_hrs_no"] += order["hours_no_integ"]
        agg[("综合", node)]["total_hrs_w"] += order["hours_with_integ"]
        # (cohort, 综合) 同人群跨节点
        agg[(cohort, "综合")]["count"] += 1
        agg[(cohort, "综合")]["total_amt"] += order["amount"]
        agg[(cohort, "综合")]["total_hrs_no"] += order["hours_no_integ"]
        agg[(cohort, "综合")]["total_hrs_w"] += order["hours_with_integ"]
        # (综合, 综合) 全量
        agg[("综合", "综合")]["count"] += 1
        agg[("综合", "综合")]["total_amt"] += order["amount"]
        agg[("综合", "综合")]["total_hrs_no"] += order["hours_no_integ"]
        agg[("综合", "综合")]["total_hrs_w"] += order["hours_with_integ"]

    results = {}
    for (cohort, node), data in agg.items():
        asp = data["total_amt"] / data["count"] if data["count"] else 0
        price_w = data["total_amt"] / data["total_hrs_w"] if data["total_hrs_w"] else 0
        price_no = data["total_amt"] / data["total_hrs_no"] if data["total_hrs_no"] else 0
        results[(cohort, node)] = {
            "count": data["count"],
            "total_amt": round(data["total_amt"], 2),
            "total_hrs_no": round(data["total_hrs_no"], 2),
            "total_hrs_w": round(data["total_hrs_w"], 2),
            "asp": round(asp, 2),
            "price_with_integ": round(price_w, 2),
            "price_no_integ": round(price_no, 2),
        }
    return results


def aggregate_by_category(matched_orders):
    """按【人群×类别】聚合计算指标"""
    agg = defaultdict(lambda: {
        "count": 0, "total_amt": 0,
        "total_hrs_no": 0, "total_hrs_w": 0
    })

    for order in matched_orders:
        key = (order["cohort"], order["category"])
        agg[key]["count"] += 1
        agg[key]["total_amt"] += order["amount"]
        agg[key]["total_hrs_no"] += order["hours_no_integ"]
        agg[key]["total_hrs_w"] += order["hours_with_integ"]

    results = {}
    for (cohort, category), data in agg.items():
        asp = data["total_amt"] / data["count"] if data["count"] else 0
        price_w = data["total_amt"] / data["total_hrs_w"] if data["total_hrs_w"] else 0
        price_no = data["total_amt"] / data["total_hrs_no"] if data["total_hrs_no"] else 0
        results[(cohort, category)] = {
            "count": data["count"],
            "total_amt": data["total_amt"],
            "asp": round(asp, 2),
            "price_with_integ": round(price_w, 2),
            "price_no_integ": round(price_no, 2),
        }

    return results


def aggregate_by_package(matched_orders):
    """按【人群×具体套餐】聚合计算指标和占比"""
    agg = defaultdict(lambda: {
        "count": 0, "total_amt": 0, "category": "", "sku_node": "",
        "total_hrs_no": 0, "total_hrs_w": 0
    })
    cohort_totals = defaultdict(int)

    for order in matched_orders:
        key = (order["cohort"], order["package"])
        agg[key]["count"] += 1
        agg[key]["total_amt"] += order["amount"]
        agg[key]["total_hrs_no"] += order["hours_no_integ"]
        agg[key]["total_hrs_w"] += order["hours_with_integ"]
        agg[key]["category"] = order.get("category", "")
        agg[key]["sku_node"] = order.get("sku_node", "")
        cohort_totals[order["cohort"]] += 1

    results = []
    for (cohort, package), data in agg.items():
        asp = data["total_amt"] / data["count"] if data["count"] else 0
        price_w = data["total_amt"] / data["total_hrs_w"] if data["total_hrs_w"] else 0
        price_no = data["total_amt"] / data["total_hrs_no"] if data["total_hrs_no"] else 0
        ratio = data["count"] / cohort_totals[cohort] if cohort_totals[cohort] else 0

        results.append({
            "cohort": cohort,
            "package": package,
            "category": data["category"],
            "sku_node": data["sku_node"],
            "count": data["count"],
            "total_amt": round(data["total_amt"], 2),
            "asp": round(asp, 2),
            "price_with_integ": round(price_w, 2),
            "price_no_integ": round(price_no, 2),
            "ratio": round(ratio, 4),
        })

    results.sort(key=lambda x: (x["cohort"], -x["count"]))
    return results, cohort_totals


def extract_budget_from_sku(sku_file):
    """从SKU测算文档(Sheet: SKU测算&实际对比)提取预算指标。

    返回结构：
    {
        ('含积分', cohort, node): {'asp': ..., 'price': ...},
        ('不含积分', cohort, node): {...}
    }
    cohort: 一续/多续/综合, node: 升舱/早鸟/其余/综合
    """
    import openpyxl

    wb = openpyxl.load_workbook(sku_file, data_only=True)
    target_sheet = None
    for name in wb.sheetnames:
        if "测算" in name and "对比" in name:
            target_sheet = name
            break
    if not target_sheet:
        target_sheet = wb.sheetnames[0]

    ws = wb[target_sheet]
    print(f"  SKU测算Sheet: {target_sheet}")

    budget = {}

    # 含积分块: 行3-8
    # R3: 节点 | 一续(实际) | _ | 一续(测算) | _ | 多续(实际) | _ | 多续(测算) | _ | 综合(实际) | _ | 综合(测算)
    # R4: ASP/单课时价 重复
    # R5-8: 升舱/早鸟/其余/综合
    # 列对应：测算列 → 一续测算 C4-C5, 多续测算 C8-C9, 综合测算 C12-C13
    node_rows_with = {5: "升舱", 6: "早鸟", 7: "其余", 8: "综合"}
    for row, node in node_rows_with.items():
        cell = ws.cell(row, 1).value
        if not cell or node not in str(cell):
            continue
        # 一续测算
        a = ws.cell(row, 4).value
        p = ws.cell(row, 5).value
        if isinstance(a, (int, float)) and a > 0:
            budget[("含积分", "一续", node)] = {
                "asp": round(float(a), 2),
                "price": round(float(p), 2) if isinstance(p, (int, float)) else 0,
            }
        # 多续测算
        a = ws.cell(row, 8).value
        p = ws.cell(row, 9).value
        if isinstance(a, (int, float)) and a > 0:
            budget[("含积分", "多续", node)] = {
                "asp": round(float(a), 2),
                "price": round(float(p), 2) if isinstance(p, (int, float)) else 0,
            }
        # 综合测算
        a = ws.cell(row, 12).value
        p = ws.cell(row, 13).value
        if isinstance(a, (int, float)) and a > 0:
            budget[("含积分", "综合", node)] = {
                "asp": round(float(a), 2),
                "price": round(float(p), 2) if isinstance(p, (int, float)) else 0,
            }

    # 不含积分块: 行11-16，结构同上
    node_rows_no = {13: "升舱", 14: "早鸟", 15: "其余", 16: "综合"}
    for row, node in node_rows_no.items():
        cell = ws.cell(row, 1).value
        if not cell or node not in str(cell):
            continue
        a = ws.cell(row, 4).value
        p = ws.cell(row, 5).value
        if isinstance(a, (int, float)) and a > 0:
            budget[("不含积分", "一续", node)] = {
                "asp": round(float(a), 2),
                "price": round(float(p), 2) if isinstance(p, (int, float)) else 0,
            }
        a = ws.cell(row, 8).value
        p = ws.cell(row, 9).value
        if isinstance(a, (int, float)) and a > 0:
            budget[("不含积分", "多续", node)] = {
                "asp": round(float(a), 2),
                "price": round(float(p), 2) if isinstance(p, (int, float)) else 0,
            }
        a = ws.cell(row, 12).value
        p = ws.cell(row, 13).value
        if isinstance(a, (int, float)) and a > 0:
            budget[("不含积分", "综合", node)] = {
                "asp": round(float(a), 2),
                "price": round(float(p), 2) if isinstance(p, (int, float)) else 0,
            }

    wb.close()
    print(f"  预算数据: {len(budget)} 个组合")
    return budget


def extract_package_budget_ratio(sku_file):
    """从SKU测算文档提取套餐占比预算（行21开始，列23-34）
    返回 {(cohort, node, package_name): budget_ratio}
    """
    import openpyxl
    wb = openpyxl.load_workbook(sku_file, data_only=True)
    ws = wb['SKU测算&实际对比'] if 'SKU测算&实际对比' in wb.sheetnames else wb.active

    # 列23-34: 一续升舱测/实, 多续升舱测/实, 一续早鸟测/实, 多续早鸟测/实, 一续其余测/实, 多续其余测/实
    col_map = {
        23: ('一续', '升舱'), 24: ('一续', '升舱'),
        25: ('多续', '升舱'), 26: ('多续', '升舱'),
        27: ('一续', '早鸟'), 28: ('一续', '早鸟'),
        29: ('多续', '早鸟'), 30: ('多续', '早鸟'),
        31: ('一续', '其余'), 32: ('一续', '其余'),
        33: ('多续', '其余'), 34: ('多续', '其余'),
    }
    budget_col = [23, 25, 27, 29, 31, 33]  # 测算列

    ratios = {}
    for row in range(24, ws.max_row + 1):
        pkg_name = ws.cell(row, 1).value
        if not pkg_name or not isinstance(pkg_name, str):
            continue
        if '【VIPTHINK】' not in pkg_name:
            continue
        for col in budget_col:
            val = ws.cell(row, col).value
            if isinstance(val, (int, float)) and val > 0:
                cohort, node = col_map[col]
                ratios[(cohort, node, pkg_name)] = float(val)

    wb.close()
    print(f"  套餐占比预算: {len(ratios)} 个")
    return ratios
