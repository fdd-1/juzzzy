"""4.1 飞书电子表格构建器
- 创建电子表格
- 写入整合数据
- 应用样式: 表头颜色 + 总计行黄底 + 数据条格式(色阶)
- 命名规则: 4.1 服务指标跟进 {开始日期}-{结束日期}
"""
from __future__ import annotations
import sys
import json
import subprocess
import shutil
import time
from pathlib import Path
from datetime import date
import pandas as pd

sys.path.insert(0, str(Path(__file__).parent))
from color_scale import apply_color_scale, batch_set_style, col_letter

# 飞书云盘文件夹: 周报自动化
FOLDER_TOKEN = "CkpmfbJfTlWwx6d98PscfdOnnoe"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"


def lark_cli(cmd_args: list[str], timeout: int = 60) -> dict:
    """执行 lark-cli 命令并返回 JSON。lark-cli 默认输出 JSON, 不需要 --format。"""
    cmd = [LARK_CLI] + cmd_args
    print(f"  $ {' '.join(cmd[:6])}...")
    result = subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        shell=(sys.platform == "win32"), timeout=timeout
    )
    if result.returncode != 0:
        # exit 10 = high-risk-write 需要 --yes
        if result.returncode == 10:
            print(f"  [high-risk-write] retry with --yes")
            cmd_with_yes = cmd + ["--yes"]
            result = subprocess.run(
                cmd_with_yes, capture_output=True, text=True,
                encoding="utf-8", errors="replace",
                shell=(sys.platform == "win32"), timeout=timeout
            )
            if result.returncode != 0:
                print(f"  [错误] retry failed exit={result.returncode}")
                print(f"  stderr: {result.stderr[:300]}")
                return {"ok": False, "stderr": result.stderr}
        else:
            print(f"  [错误] exit={result.returncode}")
            print(f"  stderr: {result.stderr[:300]}")
            return {"ok": False, "stderr": result.stderr}
    try:
        return json.loads(result.stdout)
    except Exception as e:
        return {"ok": False, "raw": result.stdout, "error": str(e)}


def create_spreadsheet(title: str) -> str | None:
    """创建一个新的飞书电子表格,返回 spreadsheet_token。"""
    resp = lark_cli([
        "sheets", "+create",
        "--title", title,
        "--folder-token", FOLDER_TOKEN,
    ])
    if not resp.get("ok"):
        return None
    data = resp.get("data", {})
    token = data.get("spreadsheet_token") or data.get("spreadsheet", {}).get("spreadsheet_token")
    return token


def get_default_sheet_id(spreadsheet_token: str) -> str | None:
    """获取spreadsheet中第一个sheet的ID。"""
    resp = lark_cli([
        "sheets", "+info",
        "--spreadsheet-token", spreadsheet_token,
    ])
    if not resp.get("ok"):
        return None
    sheets = resp.get("data", {}).get("sheets", {}).get("sheets", [])
    return sheets[0].get("sheet_id") if sheets else None


def write_data_to_sheet(spreadsheet_token: str, sheet_id: str, values: list[list], batch_size: int = 5):
    """分批写入数据到sheet。values 是二维列表,首行可包含表头。"""
    if not values:
        return
    n = len(values)
    n_cols = max((len(r) for r in values), default=0)
    if n_cols == 0:
        return
    col_end = col_letter(n_cols)

    for i in range(0, n, batch_size):
        batch = values[i:i + batch_size]
        # 补齐列数
        for r in batch:
            while len(r) < n_cols:
                r.append("")
        # 全角替换 + NaN/Timestamp 处理
        safe_batch = []
        for row in batch:
            sr = []
            for v in row:
                try:
                    if pd.isna(v):
                        sr.append("")
                        continue
                except (TypeError, ValueError):
                    pass
                # Timestamp / datetime
                if isinstance(v, pd.Timestamp):
                    sr.append(v.strftime("%Y-%m-%d"))
                    continue
                if hasattr(v, 'strftime'):
                    sr.append(v.strftime("%Y-%m-%d"))
                    continue
                if isinstance(v, str):
                    sr.append(v.replace("<", "<").replace(">", ">"))
                else:
                    sr.append(v)
            safe_batch.append(sr)

        rng = f"{sheet_id}!A{i+1}:{col_end}{i+len(batch)}"
        resp = lark_cli([
            "sheets", "+write",
            "--spreadsheet-token", spreadsheet_token,
            "--range", rng,
            "--values", json.dumps(safe_batch, ensure_ascii=False),
        ])
        if not resp.get("ok"):
            print(f"  [写入失败] range={rng}")
        time.sleep(0.2)


def apply_header_style(spreadsheet_token: str, sheet_id: str, n_cols: int, header_rows: int = 1):
    """应用表头样式: 蓝底白字加粗。fontSize 必须 >= 9。"""
    col_end = col_letter(n_cols)
    rng = f"{sheet_id}!A1:{col_end}{header_rows}"
    style = {
        "font": {"bold": True, "fontSize": 10},
        "backColor": "#4472C4",
        "foreColor": "#FFFFFF",
        "hAlign": 1,
        "vAlign": 1,
    }
    lark_cli([
        "sheets", "+set-style",
        "--spreadsheet-token", spreadsheet_token,
        "--range", rng,
        "--style", json.dumps(style),
    ])


def apply_summary_row_style(spreadsheet_token: str, sheet_id: str, row_indexes: list[int], n_cols: int):
    """对汇总行应用黄底加粗样式。row_indexes 是 1-based 行号列表。"""
    if not row_indexes:
        return
    col_end = col_letter(n_cols)
    ops = []
    for r in row_indexes:
        rng = f"{sheet_id}!A{r}:{col_end}{r}"
        ops.append({
            "ranges": [rng],
            "style": {
                "font": {"bold": True},
                "backColor": "#FFF2CC",
            }
        })
    batch_set_style(spreadsheet_token, ops)


def find_data_bar_columns(headers: list[str], target_keywords: list[str]) -> list[int]:
    """根据列名关键词找出需要数据条上色的列(0-based)。"""
    cols = []
    for i, h in enumerate(headers):
        s = str(h)
        if any(kw in s for kw in target_keywords):
            cols.append(i)
    return cols


def build_4_1_spreadsheet(merged_df: pd.DataFrame, start_date: date, end_date: date) -> dict:
    """构建 4.1 板块的电子表格,返回 {token, sheet_id, url}。"""
    title = f"4.1 服务指标跟进 {start_date.strftime('%m%d')}-{end_date.strftime('%m%d')}"
    print(f"\n=== 4.1 飞书电子表格: {title} ===")

    # 1. 创建表
    token = create_spreadsheet(title)
    if not token:
        print("[错误] 创建电子表格失败")
        return {}
    print(f"  电子表格 token: {token}")

    sheet_id = get_default_sheet_id(token)
    print(f"  sheet_id: {sheet_id}")
    if not sheet_id:
        return {"token": token}

    # 2. 准备数据
    headers = list(merged_df.columns)
    data_rows = []
    for _, row in merged_df.iterrows():
        rec = []
        for v in row:
            try:
                if pd.isna(v):
                    rec.append("")
                    continue
            except (TypeError, ValueError):
                pass
            if isinstance(v, pd.Timestamp):
                rec.append(v.strftime("%Y-%m-%d"))
                continue
            if hasattr(v, 'strftime'):
                rec.append(v.strftime("%Y-%m-%d"))
                continue
            if hasattr(v, "item"):
                rec.append(v.item())
            else:
                rec.append(v)
        data_rows.append(rec)

    all_values = [headers] + data_rows
    n_cols = len(headers)

    # 3. 写入
    print(f"  写入 {len(all_values)} 行 x {n_cols} 列...")
    write_data_to_sheet(token, sheet_id, all_values, batch_size=5)

    # 4. 表头样式
    apply_header_style(token, sheet_id, n_cols, header_rows=1)

    # 4.5 全表居中对齐
    print(f"  设置全表居中...")
    all_range = f"{sheet_id}!A1:{col_letter(n_cols)}{len(all_values)}"
    lark_cli([
        "sheets", "+set-style",
        "--spreadsheet-token", token,
        "--range", all_range,
        "--style", json.dumps({"hAlign": 1, "vAlign": 1}),
    ])

    # 5. 找出汇总行(LP=='总计')并加色
    summary_rows = []
    if "LP" in headers:
        lp_col = headers.index("LP")
        for i, row in enumerate(data_rows):
            if str(row[lp_col]) == "总计":
                summary_rows.append(i + 2)  # +2 = 表头(1) + 数据从第2行起
    if summary_rows:
        print(f"  汇总行: {summary_rows}")
        apply_summary_row_style(token, sheet_id, summary_rows, n_cols)

    # 6. 数据条色阶 (按需求中指定的列)
    data_bar_keywords = [
        # 首通语义点执行
        "邀请添加企微", "一家多娃问询执行率", "转介绍执行率",
        # 首通
        "首通_跟进率", "首通_及时跟进率", "首通_企微绑定率", "首通_秒挂占比",
        # 首课语义点执行
        "询问上课感受执行率",
        # 首课
        "首课跟进率", "首课及时跟进率",
        # 首专
        "首专跟进率", "首专及时跟进率",
    ]
    bar_cols = find_data_bar_columns(headers, data_bar_keywords)
    print(f"  数据条列: {len(bar_cols)} 个")

    for col_idx in bar_cols:
        # 取列值(包括汇总行,色阶基于全部数据)
        values = []
        for row in data_rows:
            v = row[col_idx] if col_idx < len(row) else None
            values.append(v)
        # 应用色阶(数据从第2行开始)
        n = apply_color_scale(token, sheet_id, col_idx, values, start_row=2)
        if n > 0:
            print(f"    {headers[col_idx]}: {n} 个单元格上色")

    url = f"https://hcnig43mb8gp.feishu.cn/sheets/{token}"
    print(f"  ✓ URL: {url}")

    return {
        "token": token,
        "sheet_id": sheet_id,
        "url": url,
        "title": title,
    }


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--merged", required=True, help="_merged_4_1.xlsx 路径")
    parser.add_argument("--start-date", required=True)
    parser.add_argument("--end-date", required=True)
    args = parser.parse_args()

    from datetime import datetime
    start = datetime.strptime(args.start_date, "%Y-%m-%d").date()
    end = datetime.strptime(args.end_date, "%Y-%m-%d").date()

    df = pd.read_excel(args.merged)
    result = build_4_1_spreadsheet(df, start, end)
    print(f"\n=== 结果 ===")
    print(json.dumps(result, ensure_ascii=False, indent=2))
