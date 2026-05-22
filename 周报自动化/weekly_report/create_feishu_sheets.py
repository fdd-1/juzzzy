"""创建飞书电子表格并嵌入到文档中

流程：
1. 读取格式化 Excel 数据
2. 用 lark-cli sheets +create 创建两张电子表格（主表 + AI 表）
3. 用 lark-cli docs +update 在文档中嵌入 sheet 块
"""
from __future__ import annotations
import sys, io, json, subprocess, shutil
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pandas as pd
from pathlib import Path

ROOT = Path(__file__).parent
EXPORT_BASE = ROOT.parent / "exports" / "4_1"
MAIN_EXCEL = EXPORT_BASE / "4_1_格式化.xlsx"
AI_EXCEL = EXPORT_BASE / "4_1_AI学情_格式化.xlsx"

FEISHU_DOC_ID = "ZFb3d1CZFobHnTxSSnMcgcVanyg"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"


def excel_to_data(excel_path: Path):
    """读取 Excel 转为 headers + data 数组"""
    df = pd.read_excel(excel_path)
    headers = df.columns.tolist()
    data = []
    for _, row in df.iterrows():
        data.append([str(v) if pd.notna(v) else "" for v in row])
    return headers, data


def create_sheet(title: str) -> tuple[str, str]:
    """创建空的飞书电子表格，返回 (spreadsheet_token, sheet_id)"""
    cmd = [LARK_CLI, "sheets", "+create", "--title", title]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            shell=(sys.platform == "win32"))
    if result.returncode != 0:
        print(f"[ERROR] 创建电子表格失败: {result.stderr}")
        return None, None
    resp = json.loads(result.stdout)
    if not resp.get("ok"):
        print(f"[ERROR] API 错误: {resp}")
        return None, None
    token = resp["data"]["spreadsheet_token"]
    url = resp["data"].get("url", "")
    print(f"  [OK] 创建成功: {title}")
    print(f"      token={token}, url={url}")
    # 获取 sheet_id
    cmd2 = [LARK_CLI, "sheets", "+info", "--spreadsheet-token", token]
    r2 = subprocess.run(cmd2, capture_output=True, text=True, encoding="utf-8",
                        shell=(sys.platform == "win32"))
    if r2.returncode == 0:
        info = json.loads(r2.stdout)
        sheets_data = info.get("data", {}).get("sheets", {})
        if isinstance(sheets_data, dict):
            sheets_list = sheets_data.get("sheets", [])
        else:
            sheets_list = sheets_data
        sheet_id = sheets_list[0]["sheet_id"] if sheets_list else None
    else:
        sheet_id = None
    print(f"      sheet_id={sheet_id}")
    return token, sheet_id


def write_data_in_batches(token: str, sheet_id: str, headers: list, data: list, batch_size: int = 5):
    """分批写入数据到电子表格（小批次以适应 Windows 命令行长度限制）"""
    all_rows = [headers] + data
    total = len(all_rows)
    written = 0

    # 计算结束列字母
    n_cols = len(headers)
    if n_cols <= 26:
        end_col = chr(ord('A') + n_cols - 1)
    else:
        q, r = divmod(n_cols - 1, 26)
        end_col = chr(ord('A') + q - 1) + chr(ord('A') + r)

    for i in range(0, total, batch_size):
        batch = all_rows[i:i + batch_size]
        start_row = i + 1
        end_row = i + len(batch)
        range_str = f"{sheet_id}!A{start_row}:{end_col}{end_row}"

        values_json = json.dumps(batch, ensure_ascii=False)
        cmd = [
            LARK_CLI, "sheets", "+write",
            "--spreadsheet-token", token,
            "--range", range_str,
            "--values", values_json,
        ]
        result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                                shell=(sys.platform == "win32"))
        if result.returncode != 0 or not result.stdout:
            print(f"  [ERROR] rows {start_row}-{end_row}: cmd too long or failed")
            continue
        try:
            resp = json.loads(result.stdout)
        except Exception:
            print(f"  [ERROR] rows {start_row}-{end_row}: invalid response")
            continue
        if resp.get("ok"):
            written += len(batch)
        else:
            err = resp.get("error", {}).get("message", "unknown")
            print(f"  [ERROR] rows {start_row}-{end_row}: {err}")

    print(f"  [OK] 写入完成: {written}/{total} 行")


def embed_sheet_in_doc(doc_id: str, sheet_token: str, title: str):
    """在文档中嵌入电子表格块"""
    # 构建 sheet 嵌入块 XML
    xml = f'<h4>{title}</h4><sheet token="{sheet_token}"></sheet>'
    cmd = [
        LARK_CLI, "docs", "+update", "--api-version", "v2",
        "--doc", doc_id,
        "--command", "append",
        "--content", xml,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                            shell=(sys.platform == "win32"))
    if result.returncode != 0:
        print(f"[ERROR] 嵌入失败: {result.stderr}")
        return False
    resp = json.loads(result.stdout)
    ok = resp.get("ok", False)
    if ok:
        print(f"  [OK] 已嵌入: {title}")
    else:
        print(f"[ERROR] 嵌入失败: {resp}")
    return ok


def main():
    print("=" * 50)
    print("4.1 创建飞书电子表格并嵌入文档")
    print("=" * 50)

    print("\n[1/5] 读取主表数据...")
    main_headers, main_data = excel_to_data(MAIN_EXCEL)
    print(f"  主表: {len(main_headers)} 列 × {len(main_data)} 行")

    print("\n[2/5] 读取 AI 表数据...")
    ai_headers, ai_data = excel_to_data(AI_EXCEL)
    print(f"  AI 表: {len(ai_headers)} 列 × {len(ai_data)} 行")

    print("\n[3/5] 创建主表电子表格...")
    main_token, main_sheet_id = create_sheet("4.1 主表 - 服务指标+语义分析+LP架构")
    if not main_token:
        return
    print("  写入主表数据...")
    write_data_in_batches(main_token, main_sheet_id, main_headers, main_data)

    print("\n[4/5] 创建 AI 表电子表格...")
    ai_token, ai_sheet_id = create_sheet("4.1 AI学情助手汇总")
    if not ai_token:
        return
    print("  写入 AI 表数据...")
    write_data_in_batches(ai_token, ai_sheet_id, ai_headers, ai_data)

    print("\n[5/5] 在文档中嵌入电子表格...")
    embed_sheet_in_doc(FEISHU_DOC_ID, main_token, "主表 - 服务指标 + 语义分析 + LP架构")
    embed_sheet_in_doc(FEISHU_DOC_ID, ai_token, "AI 学情助手")

    print("\n[DONE] 完成！请检查飞书文档：")
    print(f"  https://hcnig43mb8gp.feishu.cn/docx/{FEISHU_DOC_ID}")


if __name__ == "__main__":
    main()
