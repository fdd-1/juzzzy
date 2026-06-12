"""创建最终统一周报文档（含所有结论 callout 和嵌入表格）

文档结构（参考: https://my.feishu.cn/docx/Veyzd0uGtoBKvyxRGgIcp3MznQc）:

4.1 服务指标跟进 & 语义分析
  - callout 4.1
  - 服务指标数据表
4.2 组班多意向占比
  - callout 4.2
  - 组班意向数据表
4.3 群发跟进
  - callout 4.3
  - 群发消息数据表
4.4 停课唤醒
  - callout 4.4
  - 停课唤醒数据表
4.5 服务月跟进
  - callout 4.5_fuwuyue
  - 服务月跟进数据表
  - 语义分析
  - callout 4.5_sop
  - 服务池SOP数据表
4.6 系统电话外呼 & 微信回复监控
  - callout 4.6_waihu
  - 整体系统外呼数据表
  - callout 4.6_qiwei
  - 整体微信发送&回复比数据表
"""
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
FOLDER_TOKEN = "CkpmfbJfTlWwx6d98PscfdOnnoe"


# 所有飞书表格的 token 和 sheet_id（最新版本）
SHEETS = {
    "4.1": {"token": "Q1ZRsjl8ShEaT6tzfOQceQCZn8w", "sheet_id": "6e8dab"},
    "4.2": {"token": "PLKAsm2LDhZfFttXWGrcOzjNnhc", "sheet_id": "3769a0"},
    "4.3": {"token": "GsS8sMhnghj3iYtCrMLcbduknEh", "sheet_id": "869169"},
    "4.4": {"token": "R66dsCmrChNkGutCF3RcPimGn0b", "sheet_id": "960e5b"},
    "4.5_fuwuyue": {"token": "Nuges5LbChBv53tnqCRcgLjlnKb", "sheet_id": "35065c"},
    "4.5_sop": {"token": "Dk8hsvXFThWH8ht3HKncdzisnCb", "sheet_id": "a0acc2"},
    "4.6_waihu": {"token": "PTX3snLUzhJKRetCeAxcKQpSnQf", "sheet_id": "15daf8"},
    "4.6_qiwei": {"token": "WUl4s8LkNhP61UtMxXucC2S0nFf", "sheet_id": "610ce2"},
}


def build_doc_xml(start_date: date, end_date: date, callouts: dict, callout_4_1: str) -> str:
    """构建完整文档 XML，按参考文档顺序"""
    title = f"服务周报 ({start_date.strftime('%m.%d')}-{end_date.strftime('%m.%d')})"

    parts = [f"<title>{title}</title>"]

    # 4.1 服务指标跟进 & 语义分析
    parts.append('<h3>4.1 服务指标跟进 &amp; 语义分析</h3>')
    parts.append(callout_4_1)
    parts.append('<h6>服务指标数据表</h6>')
    parts.append(f'<sheet token="{SHEETS["4.1"]["token"]}" sheet-id="{SHEETS["4.1"]["sheet_id"]}"></sheet>')

    # 4.2 组班多意向占比
    parts.append('<h3>4.2 组班多意向占比</h3>')
    parts.append(callouts['4.2'])
    parts.append(f'<sheet token="{SHEETS["4.2"]["token"]}" sheet-id="{SHEETS["4.2"]["sheet_id"]}"></sheet>')

    # 4.3 群发跟进
    parts.append('<h3>4.3 群发跟进</h3>')
    parts.append(callouts['4.3'])
    parts.append(f'<sheet token="{SHEETS["4.3"]["token"]}" sheet-id="{SHEETS["4.3"]["sheet_id"]}"></sheet>')

    # 4.4 停课唤醒
    parts.append('<h3>4.4 停课唤醒</h3>')
    parts.append(callouts['4.4'])
    parts.append(f'<sheet token="{SHEETS["4.4"]["token"]}" sheet-id="{SHEETS["4.4"]["sheet_id"]}"></sheet>')

    # 4.5 服务月跟进
    parts.append('<h3>4.5 服务月跟进</h3>')
    parts.append(callouts['4.5_fuwuyue'])
    parts.append(f'<sheet token="{SHEETS["4.5_fuwuyue"]["token"]}" sheet-id="{SHEETS["4.5_fuwuyue"]["sheet_id"]}"></sheet>')
    parts.append('<p><b>#语义分析</b></p>')
    parts.append(callouts['4.5_sop'])
    parts.append(f'<sheet token="{SHEETS["4.5_sop"]["token"]}" sheet-id="{SHEETS["4.5_sop"]["sheet_id"]}"></sheet>')

    # 4.6 系统电话外呼 & 微信回复监控
    parts.append('<h3>4.6 系统电话外呼 &amp; 微信回复监控</h3>')
    parts.append('<p><b>#整体系统外呼</b></p>')
    parts.append(callouts['4.6_waihu'])
    parts.append(f'<sheet token="{SHEETS["4.6_waihu"]["token"]}" sheet-id="{SHEETS["4.6_waihu"]["sheet_id"]}"></sheet>')
    parts.append('<p><b>#整体微信发送&amp;回复比</b></p>')
    parts.append(callouts['4.6_qiwei'])
    parts.append(f'<sheet token="{SHEETS["4.6_qiwei"]["token"]}" sheet-id="{SHEETS["4.6_qiwei"]["sheet_id"]}"></sheet>')

    return '\n'.join(parts)


def create_doc(xml_content: str, start_date: date) -> dict:
    """创建飞书文档"""
    tmp_dir = Path(__file__).parent.parent / "exports" / "_tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_file_name = f"weekly_report_final_{start_date.strftime('%Y%m%d')}.xml"
    tmp_file = tmp_dir / tmp_file_name
    tmp_file.write_text(xml_content, encoding="utf-8")

    print(f"\n=== 创建最终周报文档 ===")
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
        print(f"  [错误] exit={result.returncode}")
        print(f"  stderr: {result.stderr[:500]}")
        return {}

    try:
        resp = json.loads(result.stdout)
    except Exception:
        print(f"  [错误] 解析失败: {result.stdout[:300]}")
        return {}

    data = resp.get("data", {})
    doc_id = data.get("document_id") or data.get("document", {}).get("document_id")
    if not doc_id:
        return {}

    url = f"https://hcnig43mb8gp.feishu.cn/docx/{doc_id}"
    print(f"  ✓ doc_id: {doc_id}")
    print(f"  ✓ URL: {url}")

    return {"doc_id": doc_id, "url": url}


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).parent))
    from conclusion_generator import generate_all_callouts

    base_dir = Path(__file__).parent.parent / "exports" / "weekly_20260601_20260607"

    # 读取 4.1 callout
    callout_4_1_file = base_dir / "4_1" / "_callout_4_1.xml"
    callout_4_1 = callout_4_1_file.read_text(encoding="utf-8") if callout_4_1_file.exists() else "<callout emoji=\"❗\"><p>4.1 数据待补充</p></callout>"

    # 生成 4.2-4.6 callouts
    callouts = generate_all_callouts(base_dir)

    from datetime import datetime
    start = datetime(2026, 6, 1).date()
    end = datetime(2026, 6, 7).date()

    xml = build_doc_xml(start, end, callouts, callout_4_1)
    result = create_doc(xml, start)

    print(f"\n=== 最终结果 ===")
    print(json.dumps(result, ensure_ascii=False, indent=2))
