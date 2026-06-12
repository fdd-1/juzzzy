"""创建最终统一周报文档（含全部修正：完整数据、参考文档结论格式）"""
from __future__ import annotations
import sys
import json
import subprocess
import shutil
from pathlib import Path
from datetime import date

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, ValueError):
        pass

LARK_CLI = shutil.which("lark-cli") or "lark-cli"
FOLDER_TOKEN = "JpSRflVoWlwxZxdBgg7cFbBNnrc"  # 目标文件夹

# 最新版飞书表格 token（2026-06-10 新数据）
SHEET_TOKENS = {
    "4.1": "KqK9smqv0hefZDtomjAc7odcnfD",
    "4.2": "DLm7sSS0vhEcZZtTc5Ic6pmhnOd",
    "4.3": "DcW1sDaJ5hpwS5te27ucOVasn8g",
    "4.4": "AMLastnqKhoTtYtUO1Rc1NRsnDc",
    "4.5_fuwuyue": "L2fysKrHXhKbxstgxqncysbFn0e",
    "4.5_sop": "YluVssnLPhq71atvFJRcSr9mn3f",
    "4.6_waihu": "BkKNsK7uwhpjkstd1TicYKtHnKf",
    "4.6_qiwei": "RW1vsAWB3hpqXDtl7NQcn0GhnDh",
}


def get_sheet_id(token: str) -> str:
    """获取 spreadsheet 的第一个 sheet_id"""
    cmd = [LARK_CLI, "sheets", "+info", "--spreadsheet-token", token]
    result = subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        shell=(sys.platform == "win32"), timeout=30,
    )
    if result.returncode != 0:
        return ""
    try:
        resp = json.loads(result.stdout)
        sheets = resp.get("data", {}).get("sheets", {}).get("sheets", [])
        return sheets[0].get("sheet_id", "") if sheets else ""
    except Exception:
        return ""


def build_final_doc(sheet_ids: dict, callouts: dict) -> dict:
    """构建最终统一文档"""
    parts = ['<title>服务周报 (06.01-06.07)</title>']

    # 4.1
    parts.append('<h3>4.1 服务指标跟进 &amp; 语义分析</h3>')
    parts.append(callouts['4.1'])
    parts.append('<h6>服务指标数据表</h6>')
    parts.append(f'<sheet token="{SHEET_TOKENS["4.1"]}" sheet-id="{sheet_ids["4.1"]}"></sheet>')

    # 4.2
    parts.append('<h3>4.2 组班多意向占比</h3>')
    parts.append(callouts['4.2'])
    parts.append(f'<sheet token="{SHEET_TOKENS["4.2"]}" sheet-id="{sheet_ids["4.2"]}"></sheet>')

    # 4.3
    parts.append('<h3>4.3 群发跟进</h3>')
    parts.append(callouts['4.3'])
    parts.append(f'<sheet token="{SHEET_TOKENS["4.3"]}" sheet-id="{sheet_ids["4.3"]}"></sheet>')

    # 4.4
    parts.append('<h3>4.4 停课唤醒</h3>')
    parts.append(callouts['4.4'])
    parts.append(f'<sheet token="{SHEET_TOKENS["4.4"]}" sheet-id="{sheet_ids["4.4"]}"></sheet>')

    # 4.5
    parts.append('<h3>4.5 服务月跟进</h3>')
    parts.append(callouts['4.5_fuwuyue'])
    parts.append(f'<sheet token="{SHEET_TOKENS["4.5_fuwuyue"]}" sheet-id="{sheet_ids["4.5_fuwuyue"]}"></sheet>')
    parts.append('<p><b>#语义分析</b></p>')
    parts.append(callouts['4.5_sop'])
    parts.append(f'<sheet token="{SHEET_TOKENS["4.5_sop"]}" sheet-id="{sheet_ids["4.5_sop"]}"></sheet>')

    # 4.6
    parts.append('<h3>4.6 系统电话外呼 &amp; 微信回复监控</h3>')
    parts.append('<p><b>#整体系统外呼</b></p>')
    parts.append(callouts['4.6_waihu'])
    parts.append(f'<sheet token="{SHEET_TOKENS["4.6_waihu"]}" sheet-id="{sheet_ids["4.6_waihu"]}"></sheet>')
    parts.append('<p><b>#整体微信发送&amp;回复比</b></p>')
    parts.append(callouts['4.6_qiwei'])
    parts.append(f'<sheet token="{SHEET_TOKENS["4.6_qiwei"]}" sheet-id="{sheet_ids["4.6_qiwei"]}"></sheet>')

    xml = '\n'.join(parts)

    # 写入临时文件并创建文档
    tmp_dir = Path(__file__).parent.parent / "exports" / "_tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_file_name = "weekly_report_v3_20260601.xml"
    tmp_file = tmp_dir / tmp_file_name
    tmp_file.write_text(xml, encoding="utf-8")

    print(f"  XML: {tmp_file}")

    cmd = [
        LARK_CLI, "docs", "+create",
        "--api-version", "v2",
        "--folder-token", FOLDER_TOKEN,
        "--content", f"@{tmp_file_name}",
    ]

    result = subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        shell=(sys.platform == "win32"), timeout=60,
        cwd=str(tmp_dir),
    )

    if result.returncode == 10:
        result = subprocess.run(
            cmd + ["--yes"], capture_output=True, text=True,
            encoding="utf-8", errors="replace",
            shell=(sys.platform == "win32"), timeout=60,
            cwd=str(tmp_dir),
        )

    if result.returncode != 0:
        print(f"  [错误] exit={result.returncode}, stderr={result.stderr[:500]}")
        return {}

    try:
        resp = json.loads(result.stdout)
    except Exception:
        return {}

    data = resp.get("data", {})
    doc_id = data.get("document_id") or data.get("document", {}).get("document_id")
    if not doc_id:
        return {}

    return {
        "doc_id": doc_id,
        "url": f"https://hcnig43mb8gp.feishu.cn/docx/{doc_id}",
    }


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).parent))
    from conclusion_generator_v2 import generate_all_callouts

    base_dir = Path(__file__).parent.parent / "exports" / "weekly_20260601_20260607"

    # 1. 获取所有 sheet_id
    print("=== 获取 sheet_id ===")
    sheet_ids = {}
    for key, token in SHEET_TOKENS.items():
        sid = get_sheet_id(token)
        sheet_ids[key] = sid
        print(f"  {key}: {sid}")

    # 2. 生成结论
    print("\n=== 生成结论 ===")
    callouts = generate_all_callouts(base_dir)

    # 3. 创建最终文档
    print("\n=== 创建最终文档 ===")
    result = build_final_doc(sheet_ids, callouts)

    print(f"\n=== 最终结果 ===")
    print(json.dumps(result, ensure_ascii=False, indent=2))
