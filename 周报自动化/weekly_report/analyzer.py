"""
数据分析模块 - 从表格数据中提取指标、计算环比、生成结论文本
"""
from typing import Optional


def find_total_row(records: list[dict], key_col: str) -> Optional[dict]:
    """找到汇总行（通常包含'总计'或'整体'）"""
    for r in records:
        val = str(r.get(key_col, "")).strip()
        if val in ("总计", "汇总", "整体", "海外团队", "海外教学服务部"):
            return r
    return records[0] if records else None


def find_group_rows(records: list[dict], key_col: str) -> list[dict]:
    """找到各组数据行（排除汇总行）"""
    skip = {"总计", "汇总", "整体", "海外团队", "海外教学服务部", "", "None"}
    return [r for r in records if str(r.get(key_col, "")).strip() not in skip]


def get_metric_value(record: dict, field_name: str) -> Optional[float]:
    """从记录中提取指标值，模糊匹配字段名"""
    for key, val in record.items():
        if field_name in key and val is not None:
            try:
                return float(val)
            except (ValueError, TypeError):
                pass
    return None


def analyze_anomalies(
    records: list[dict],
    key_col: str,
    metric_field: str,
    threshold: float,
    direction: str = "below",
) -> list[dict]:
    """找出低于/高于阈值的异常组"""
    anomalies = []
    for r in records:
        group = str(r.get(key_col, "")).strip()
        if not group or group in ("总计", "汇总", "整体"):
            continue
        val = get_metric_value(r, metric_field)
        if val is None:
            continue
        if direction == "below" and val < threshold:
            anomalies.append({"group": group, "value": val, "gap": threshold - val})
        elif direction == "above" and val > threshold:
            anomalies.append({"group": group, "value": val, "gap": val - threshold})
    return sorted(anomalies, key=lambda x: x["gap"], reverse=True)


def compare_with_previous(
    current_records: list[dict],
    previous_records: list[dict],
    key_col: str,
    metric_field: str,
) -> dict:
    """计算本周 vs 上周的环比"""
    curr_total = find_total_row(current_records, key_col)
    prev_total = find_total_row(previous_records, key_col)
    if not curr_total or not prev_total:
        return {"current": None, "previous": None, "change": None}
    curr_val = get_metric_value(curr_total, metric_field)
    prev_val = get_metric_value(prev_total, metric_field)
    change = None
    if curr_val is not None and prev_val is not None and prev_val != 0:
        change = (curr_val - prev_val) / prev_val
    return {"current": curr_val, "previous": prev_val, "change": change}


def generate_service_tracking_conclusion(
    current_data: list[dict],
    previous_data: list[dict],
    config: dict,
) -> str:
    """生成 4.1 服务指标跟进结论"""
    key_col = config["key_col"]
    total = find_total_row(current_data, key_col)
    if not total:
        return "数据缺失，无法生成结论"

    lines = []
    lines.append("跟进：")

    follow_rate = get_metric_value(total, "跟进")
    timely_rate = get_metric_value(total, "及时跟进")
    new_students = get_metric_value(total, "新生数")

    if follow_rate and new_students and new_students > 0:
        follow_pct = follow_rate / new_students if follow_rate > 1 else follow_rate
    else:
        follow_pct = follow_rate

    if timely_rate and new_students and new_students > 0:
        timely_pct = timely_rate / new_students if timely_rate > 1 else timely_rate
    else:
        timely_pct = timely_rate

    target = 0.95
    target_status = "达本月目标（95%）" if (timely_pct or 0) >= target else "未达本月目标（95%）"
    lines.append(f"首通：总体跟进率{_fmt_pct(follow_pct)}；及时跟进率{_fmt_pct(timely_pct)}，{target_status}")

    groups = find_group_rows(current_data, key_col)
    anomalies = analyze_anomalies(groups, key_col, "及时跟进", 0.90, "below")
    for a in anomalies[:3]:
        lines.append(f"{a['group']}及时跟进率仅{_fmt_pct(a['value'])}，需注意")

    bind_rate = get_metric_value(total, "企微绑定")
    if bind_rate is not None and new_students and new_students > 0:
        bind_pct = bind_rate / new_students if bind_rate > 1 else bind_rate
        comp = compare_with_previous(current_data, previous_data, key_col, "企微绑定")
        change_text = ""
        if comp["change"] is not None:
            direction = "上升" if comp["change"] > 0 else "下降"
            change_text = f"较上周{direction}{abs(comp['change'])*100:.0f}%"
        lines.append(f"——企微绑定率整体为{_fmt_pct(bind_pct)}，{change_text}")

    return "\n".join(lines)


def _fmt_pct(value, digits=2) -> str:
    if value is None:
        return "N/A"
    if isinstance(value, (int, float)):
        if value <= 1:
            return f"{value * 100:.{digits}f}%"
        return f"{value:.{digits}f}%"
    return str(value)
