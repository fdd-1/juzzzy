"""通用飞书表格生成器（4.2-4.6 板块）

简化版本，只做：
1. 创建电子表格
2. 写入数据
3. 全表居中
4. 表头样式
5. 可选的数据条色阶
"""
from __future__ import annotations
import sys
import json
import subprocess
import shutil
import time
from pathlib import Path
from datetime import date

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, ValueError):
        pass

import pandas as pd

LARK_CLI = shutil.which("lark-cli") or "lark-cli"
FOLDER_TOKEN = "CkpmfbJfTlWwx6d98PscfdOnnoe"


def lark_cli(cmd_args: list[str], timeout: int = 60) -> dict:
    """执行 lark-cli."""
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
    except Exception as e:
        return {"ok": False, "raw": result.stdout, "error": str(e)}


def col_letter(n: int) -> str:
    """列号转字母: 1→A, 27→AA."""
    result = ""
    while n > 0:
        n -= 1
        result = chr(65 + n % 26) + result
        n //= 26
    return result


def build_simple_sheet(df: pd.DataFrame, title: str, color_scale_cols: list[str] = None) -> dict:
    """创建简单飞书表格。

    Args:
        df: 数据框
        title: 表格标题
        color_scale_cols: 需要应用数据条的列名列表

    Returns:
        {token, sheet_id, url}
    """
    print(f"\n=== 创建飞书表格: {title} ===")

    # 1. 创建表格
    resp = lark_cli(["sheets", "+create", "--title", title, "--folder-token", FOLDER_TOKEN])
    if not resp.get("ok"):
        print(f"  [错误] 创建失败")
        return {}

    token = resp.get("data", {}).get("spreadsheet_token") or resp.get("data", {}).get("spreadsheet", {}).get("spreadsheet_token")
    if not token:
        print(f"  [错误] 找不到 token")
        return {}

    print(f"  电子表格 token: {token}")

    # 2. 获取 sheet_id
    resp = lark_cli(["sheets", "+info", "--spreadsheet-token", token])
    sheets = resp.get("data", {}).get("sheets", {}).get("sheets", [])
    if not sheets:
        print(f"  [错误] sheets 为空, response: {json.dumps(resp, ensure_ascii=False)[:300]}")
        return {}
    sheet_id = sheets[0].get("sheet_id")
    if not sheet_id:
        print(f"  [错误] 找不到 sheet_id")
        return {}

    print(f"  sheet_id: {sheet_id}")

    # 3. 准备数据
    headers = list(df.columns)
    data_rows = []
    for _, row in df.iterrows():
        rec = []
        for v in row:
            if pd.isna(v):
                rec.append("")
            elif isinstance(v, pd.Timestamp):
                rec.append(v.strftime("%Y-%m-%d"))
            elif hasattr(v, 'strftime'):
                rec.append(v.strftime("%Y-%m-%d"))
            elif hasattr(v, "item"):
                rec.append(v.item())
            else:
                rec.append(v)
        data_rows.append(rec)

    all_values = [headers] + data_rows
    n_rows = len(all_values)
    n_cols = len(headers)

    print(f"  写入 {n_rows} 行 x {n_cols} 列...")

    # 4. 分批写入
    batch_size = 5
    for i in range(0, n_rows, batch_size):
        batch = all_values[i:i + batch_size]
        for row in batch:
            while len(row) < n_cols:
                row.append("")

        safe_batch = []
        for row in batch:
            sr = []
            for v in row:
                if isinstance(v, str):
                    sr.append(v.replace("<", "<").replace(">", ">"))
                else:
                    sr.append(v)
            safe_batch.append(sr)

        rng = f"{sheet_id}!A{i+1}:{col_letter(n_cols)}{i+len(batch)}"
        lark_cli([
            "sheets", "+write",
            "--spreadsheet-token", token,
            "--range", rng,
            "--values", json.dumps(safe_batch, ensure_ascii=False),
        ])
        time.sleep(0.2)

    # 5. 表头样式
    rng = f"{sheet_id}!A1:{col_letter(n_cols)}1"
    lark_cli([
        "sheets", "+set-style",
        "--spreadsheet-token", token,
        "--range", rng,
        "--style", json.dumps({"font": {"bold": True, "fontSize": 10}, "backColor": "#4472C4", "foreColor": "#FFFFFF", "hAlign": 1, "vAlign": 1}),
    ])

    # 6. 全表居中
    print(f"  设置全表居中...")
    all_range = f"{sheet_id}!A1:{col_letter(n_cols)}{n_rows}"
    lark_cli([
        "sheets", "+set-style",
        "--spreadsheet-token", token,
        "--range", all_range,
        "--style", json.dumps({"hAlign": 1, "vAlign": 1}),
    ])

    # 7. 数据条色阶（如果指定）
    if color_scale_cols:
        print(f"  应用数据条: {len(color_scale_cols)} 列")
        from color_scale import apply_color_scale

        # 为每个指定列应用色阶
        for col_name in color_scale_cols:
            if col_name not in df.columns:
                print(f"    [警告] 列 '{col_name}' 不存在，跳过")
                continue

            col_idx = df.columns.get_loc(col_name)
            # 提取该列的值（跳过表头）
            col_values = df[col_name].tolist()
            apply_color_scale(token, sheet_id, col_idx, col_values, start_row=2)

    url = f"https://hcnig43mb8gp.feishu.cn/sheets/{token}"
    print(f"  ✓ URL: {url}")

    return {"token": token, "sheet_id": sheet_id, "url": url, "title": title}


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", required=True, help="输入 Excel 文件")
    parser.add_argument("--title", required=True, help="表格标题")
    parser.add_argument("--color-cols", help="数据条列名，逗号分隔")
    args = parser.parse_args()

    df = pd.read_excel(args.input)
    color_cols = args.color_cols.split(',') if args.color_cols else None

    result = build_simple_sheet(df, args.title, color_cols)
    print(f"\n=== 结果 ===")
    print(json.dumps(result, ensure_ascii=False, indent=2))
