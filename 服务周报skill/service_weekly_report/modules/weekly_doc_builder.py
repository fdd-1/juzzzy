"""汇总所有板块到统一周报文档"""
from __future__ import annotations
import sys
import json
import subprocess
from pathlib import Path
from datetime import date

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, ValueError):
        pass

import shutil

LARK_CLI = shutil.which("lark-cli") or "lark-cli"
FOLDER_TOKEN = "CkpmfbJfTlWwx6d98PscfdOnnoe"


def create_weekly_report_doc(start_date: date, end_date: date, sections: list[dict], callout_4_1: str) -> dict:
    """创建完整周报文档。

    Args:
        start_date: 开始日期
        end_date: 结束日期
        sections: 各板块信息 [{title, token, sheet_id}, ...]
        callout_4_1: 4.1 的结论 XML

    Returns:
        {doc_id, url}
    """
    title = f"服务周报 ({start_date.strftime('%m.%d')}-{end_date.strftime('%m.%d')})"

    # 构建文档 XML
    xml_parts = [f"<title>{title}</title>"]

    for section in sections:
        sec_title = section['title']
        token = section['token']
        sheet_id = section['sheet_id']

        # 4.1 板块有 callout，其他板块没有
        if '4.1' in sec_title:
            xml_parts.append(f"<h3>{sec_title}</h3>")
            xml_parts.append(callout_4_1)
            xml_parts.append(f"<h6>服务指标数据表</h6>")
        else:
            xml_parts.append(f"<h3>{sec_title}</h3>")

        xml_parts.append(f'<sheet token="{token}" sheet-id="{sheet_id}"></sheet>')

    xml_content = '\n'.join(xml_parts)

    # 写入临时文件
    tmp_dir = Path(__file__).parent.parent / "exports" / "_tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_file_name = f"weekly_report_{start_date.strftime('%Y%m%d')}.xml"
    tmp_file = tmp_dir / tmp_file_name
    tmp_file.write_text(xml_content, encoding="utf-8")

    print(f"\n=== 创建周报文档 ===")
    print(f"  XML 临时文件: {tmp_file}")

    # 创建文档
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
        print(f"  [错误] 解析响应失败: {result.stdout[:300]}")
        return {}

    if not resp.get("ok"):
        print(f"  [错误] 创建文档失败: {resp}")
        return {}

    data = resp.get("data", {})
    doc_id = data.get("document_id") or data.get("document", {}).get("document_id")
    if not doc_id:
        print(f"  [错误] 找不到 document_id")
        return {}

    url = f"https://hcnig43mb8gp.feishu.cn/docx/{doc_id}"
    print(f"  ✓ 文档创建成功")
    print(f"  doc_id: {doc_id}")
    print(f"  URL: {url}")

    return {"doc_id": doc_id, "url": url}


if __name__ == "__main__":
    # 所有板块信息（按顺序）- 最终版本
    sections = [
        {"title": "4.1 服务指标跟进 & 语义分析", "token": "Q1ZRsjl8ShEaT6tzfOQceQCZn8w", "sheet_id": "6e8dab"},
        {"title": "4.2 组班意向", "token": "D7R3sOSHlhXL8KtnJqwc1Lfbn8c", "sheet_id": "1bbf03"},
        {"title": "4.3 群发消息", "token": "YRccsY820hGUkKthjyYcuMRNnsX", "sheet_id": "9bb3ef"},
        {"title": "4.4 停课唤醒", "token": "Ewe6sDcjvhXP1atIMbDczHrknpd", "sheet_id": "bc2309"},
        {"title": "4.5 服务月跟进", "token": "FQaMs2LmnhT4tPttH4ccx80Ansb", "sheet_id": "811ff5"},
        {"title": "4.5 服务池SOP执行", "token": "DPwXs56tah17nitEwGGcJmZXnhh", "sheet_id": "9889f3"},
        {"title": "4.6 外呼监控", "token": "EHRPspEPXhJyCJtc95jcXvCxn4g", "sheet_id": "db6aec"},
        {"title": "4.6 企微回复监控", "token": "Qa0tsoTVHhAZbztQfdocZeeynec", "sheet_id": "601b5f"},
    ]

    # 读取 4.1 的 callout
    callout_file = Path(__file__).parent.parent / "exports" / "weekly_20260601_20260607" / "4_1" / "_callout_4_1.xml"
    callout_4_1 = callout_file.read_text(encoding="utf-8")

    from datetime import datetime
    start = datetime(2026, 6, 1).date()
    end = datetime(2026, 6, 7).date()

    result = create_weekly_report_doc(start, end, sections, callout_4_1)

    print(f"\n=== 结果 ===")
    print(json.dumps(result, ensure_ascii=False, indent=2))
