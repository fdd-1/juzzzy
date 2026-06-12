"""清理飞书表格 + 应用色阶

针对每个表格：
1. 删除底部空行（数据末尾之后的空行）
2. 对指定列应用色阶（绿→黄→红渐变）

色阶规则（参考 4.1）：
- 越高越好 → 高值绿色，低值红色
- 越低越好（如秒挂占比） → 高值红色，低值绿色
"""
from __future__ import annotations
import sys
import json
import subprocess
import shutil
import time
from pathlib import Path

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, ValueError):
        pass

LARK_CLI = shutil.which("lark-cli") or "lark-cli"


def lark_cli(cmd_args: list, timeout: int = 60) -> dict:
    cmd = [LARK_CLI] + cmd_args
    result = subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        shell=(sys.platform == "win32"), timeout=timeout
    )
    if result.returncode == 10:
        result = subprocess.run(
            cmd + ["--yes"], capture_output=True, text=True,
            encoding="utf-8", errors="replace",
            shell=(sys.platform == "win32"), timeout=timeout
        )
    if result.returncode != 0:
        return {"ok": False, "stderr": result.stderr}
    try:
        return json.loads(result.stdout)
    except Exception:
        return {"ok": False, "raw": result.stdout}


def col_letter(n: int) -> str:
    s = ""
    while n > 0:
        n -= 1
        s = chr(65 + n % 26) + s
        n //= 26
    return s


def lerp_color(ratio: float, reverse: bool = False) -> str:
    """ratio 0.0-1.0 渐变色。reverse=True 时反向（越低越好）"""
    if reverse:
        ratio = 1.0 - ratio

    if ratio >= 0.5:
        t = (ratio - 0.5) * 2
        r = int(255 * (1 - t) + 99 * t)
        g = int(235 * (1 - t) + 190 * t)
        b = int(132 * (1 - t) + 123 * t)
    else:
        t = ratio * 2
        r = int(248 * (1 - t) + 255 * t)
        g = int(105 * (1 - t) + 235 * t)
        b = int(107 * (1 - t) + 132 * t)
    return f"#{r:02X}{g:02X}{b:02X}"


def parse_pct_or_num(val):
    """解析百分比字符串或数字"""
    if val is None or val == "":
        return None
    s = str(val).strip()
    if s.endswith("%"):
        try:
            return float(s.rstrip("%")) / 100
        except ValueError:
            return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def get_sheet_data(spreadsheet_token: str, sheet_id: str, n_rows: int, n_cols: int) -> list:
    """读取表格数据"""
    rng = f"{sheet_id}!A1:{col_letter(n_cols)}{n_rows}"
    resp = lark_cli([
        "sheets", "+read",
        "--spreadsheet-token", spreadsheet_token,
        "--range", rng,
    ])
    if not resp.get("ok"):
        return []
    return resp.get("data", {}).get("valueRange", {}).get("values", [])


def get_sheet_info(spreadsheet_token: str) -> dict:
    """获取 sheet 信息（行数、列数）"""
    resp = lark_cli([
        "sheets", "+info",
        "--spreadsheet-token", spreadsheet_token,
    ])
    if not resp.get("ok"):
        return {}
    sheets = resp.get("data", {}).get("sheets", {}).get("sheets", [])
    if not sheets:
        return {}
    sheet = sheets[0]
    return {
        "sheet_id": sheet.get("sheet_id"),
        "row_count": sheet.get("grid_properties", {}).get("row_count", 200),
        "column_count": sheet.get("grid_properties", {}).get("column_count", 20),
    }


def find_data_end_row(data: list) -> int:
    """找到数据结束行（最后一个非空行的行号，1-based）"""
    for i in range(len(data) - 1, -1, -1):
        row = data[i]
        if any(v is not None and str(v).strip() != "" for v in row):
            return i + 1  # 1-based
    return len(data)


def delete_empty_rows(spreadsheet_token: str, sheet_id: str, total_rows: int):
    """删除数据末尾的空行（保留有数据的行）"""
    info = get_sheet_info(spreadsheet_token)
    if not info:
        print(f"  ⚠ 无法获取 sheet 信息")
        return

    n_cols = info.get("column_count", 20)
    data = get_sheet_data(spreadsheet_token, sheet_id, total_rows, n_cols)

    if not data:
        return

    last_data_row = find_data_end_row(data)
    n_empty = total_rows - last_data_row

    if n_empty <= 0:
        print(f"  ✓ 无需删除空行")
        return

    print(f"  数据行: 1-{last_data_row}, 空行: {n_empty} 行")

    # 删除空行（从 last_data_row+1 开始的所有行）
    resp = lark_cli([
        "sheets", "+delete-dimension",
        "--spreadsheet-token", spreadsheet_token,
        "--sheet-id", sheet_id,
        "--dimension", "ROWS",
        "--start-index", str(last_data_row + 1),
        "--end-index", str(total_rows),  # inclusive
    ])

    if resp.get("ok"):
        print(f"  ✓ 已删除 {n_empty} 行空行")
    else:
        print(f"  ✗ 删除空行失败: {resp.get('stderr', '')[:200]}")


def apply_color_scale_to_column(spreadsheet_token: str, sheet_id: str, col_idx: int,
                                  values: list, start_row: int = 2, reverse: bool = False) -> int:
    """对单列应用色阶"""
    nums = []
    for i, v in enumerate(values):
        n = parse_pct_or_num(v)
        nums.append((i, n))

    valid = [(i, n) for i, n in nums if n is not None]
    if len(valid) < 2:
        return 0

    vs = [n for _, n in valid]
    min_v = min(vs)
    max_v = max(vs)
    if max_v == min_v:
        return 0

    col = col_letter(col_idx + 1)
    ops = []
    for i, n in valid:
        ratio = (n - min_v) / (max_v - min_v)
        color = lerp_color(ratio, reverse=reverse)
        row = start_row + i
        ops.append({
            "ranges": [f"{sheet_id}!{col}{row}:{col}{row}"],
            "style": {"backColor": color}
        })

    # 分批应用
    sent = 0
    for i in range(0, len(ops), 30):
        batch = ops[i:i + 30]
        payload = json.dumps(batch, ensure_ascii=False)
        if len(payload) > 7500:
            # 进一步切分
            for j in range(0, len(batch), 10):
                small = batch[j:j+10]
                resp = lark_cli([
                    "sheets", "+batch-set-style",
                    "--spreadsheet-token", spreadsheet_token,
                    "--data", json.dumps(small, ensure_ascii=False)
                ])
                if resp.get("ok"):
                    sent += len(small)
        else:
            resp = lark_cli([
                "sheets", "+batch-set-style",
                "--spreadsheet-token", spreadsheet_token,
                "--data", payload
            ])
            if resp.get("ok"):
                sent += len(batch)
        time.sleep(0.2)

    return sent


def apply_color_scales(spreadsheet_token: str, sheet_id: str, color_specs: list):
    """对多列应用色阶
    color_specs: [{"col_keyword": "外呼跟进率", "reverse": False}, ...]
    """
    info = get_sheet_info(spreadsheet_token)
    if not info:
        return
    n_cols = info["column_count"]
    n_rows = info["row_count"]

    data = get_sheet_data(spreadsheet_token, sheet_id, n_rows, n_cols)
    if not data:
        return

    # 第1行是表头
    headers = data[0] if data else []

    for spec in color_specs:
        keyword = spec["col_keyword"]
        reverse = spec.get("reverse", False)
        match_type = spec.get("match", "contains")  # contains 或 exact

        # 找到匹配的列
        matched_cols = []
        for i, h in enumerate(headers):
            if h is None:
                continue
            h_str = str(h)
            if match_type == "exact":
                if h_str == keyword:
                    matched_cols.append((i, h_str))
            else:
                if keyword in h_str:
                    matched_cols.append((i, h_str))

        for col_idx, col_name in matched_cols:
            col_values = [row[col_idx] if col_idx < len(row) else None for row in data[1:]]
            n = apply_color_scale_to_column(
                spreadsheet_token, sheet_id, col_idx, col_values,
                start_row=2, reverse=reverse
            )
            if n > 0:
                arrow = "↓" if reverse else "↑"
                print(f"    {col_name} {arrow}: {n} 个单元格上色")


def process_sheet(spreadsheet_token: str, name: str, color_specs: list = None):
    """处理单个表格：删空行 + 上色阶"""
    print(f"\n=== {name} ===")
    print(f"  token: {spreadsheet_token}")

    info = get_sheet_info(spreadsheet_token)
    if not info:
        print(f"  ✗ 无法获取 sheet 信息")
        return

    sheet_id = info["sheet_id"]
    print(f"  sheet_id: {sheet_id}")

    # 1. 删除空行
    delete_empty_rows(spreadsheet_token, sheet_id, info["row_count"])

    # 2. 应用色阶
    if color_specs:
        print(f"  应用色阶: {len(color_specs)} 类")
        apply_color_scales(spreadsheet_token, sheet_id, color_specs)

    print(f"  ✓ 完成")


# 8 个表格的色阶配置（参考 SOP）
SHEETS_CONFIG = {
    "4.1": {
        "token": "KqK9smqv0hefZDtomjAc7odcnfD",
        "name": "4.1 服务指标",
        "color_specs": [
            {"col_keyword": "邀请添加企微"},
            {"col_keyword": "一家多娃问询执行率"},
            {"col_keyword": "转介绍执行率"},
            {"col_keyword": "首通_跟进率"},
            {"col_keyword": "首通_及时跟进率"},
            {"col_keyword": "首通_企微绑定率"},
            {"col_keyword": "首通_秒挂占比", "reverse": True},  # 越低越好
            {"col_keyword": "询问上课感受执行率"},
            {"col_keyword": "首课跟进率"},
            {"col_keyword": "首课及时跟进率"},
            {"col_keyword": "首专跟进率"},
            {"col_keyword": "首专及时跟进率"},
        ]
    },
    "4.2": {
        "token": "DLm7sSS0vhEcZZtTc5Ic6pmhnOd",
        "name": "4.2 组班意向",
        "color_specs": []  # 4.2 不需要色阶
    },
    "4.3": {
        "token": "DcW1sDaJ5hpwS5te27ucOVasn8g",
        "name": "4.3 群发消息",
        "color_specs": [
            {"col_keyword": "个人群发占比"}
        ]
    },
    "4.4": {
        "token": "AMLastnqKhoTtYtUO1Rc1NRsnDc",
        "name": "4.4 停课唤醒",
        "color_specs": [
            {"col_keyword": "停课占比", "reverse": True},  # 越低越好
            {"col_keyword": "唤醒率"},
        ]
    },
    "4.5_fuwuyue": {
        "token": "L2fysKrHXhKbxstgxqncysbFn0e",
        "name": "4.5 服务月跟进",
        "color_specs": [
            {"col_keyword": "服务池-外呼跟进率"},
            {"col_keyword": "服务池-综合有效跟进率"},
        ]
    },
    "4.5_sop": {
        "token": "YluVssnLPhq71atvFJRcSr9mn3f",
        "name": "4.5 服务池SOP",
        "color_specs": [
            {"col_keyword": "执行率加和"},
        ]
    },
    "4.6_waihu": {
        "token": "BkKNsK7uwhpjkstd1TicYKtHnKf",
        "name": "4.6 外呼监控",
        "color_specs": [
            {"col_keyword": "整体_覆盖率"},
            {"col_keyword": "外呼接通率"},
            {"col_keyword": "有效接通率"},
        ]
    },
    "4.6_qiwei": {
        "token": "RW1vsAWB3hpqXDtl7NQcn0GhnDh",
        "name": "4.6 企微回复",
        "color_specs": []  # 企微回复比越高越坏（用户特别说明），先不上色
    },
}


def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--only", help="只处理指定板块（如 4.1）")
    args = parser.parse_args()

    print("=" * 60)
    print("飞书表格清理 + 色阶应用")
    print("=" * 60)

    for key, cfg in SHEETS_CONFIG.items():
        if args.only and key != args.only:
            continue
        process_sheet(cfg["token"], cfg["name"], cfg["color_specs"])

    print("\n" + "=" * 60)
    print("✓ 全部完成")


if __name__ == "__main__":
    main()
