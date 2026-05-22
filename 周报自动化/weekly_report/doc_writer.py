"""
飞书文档写入模块 - 新建文档并写入周报内容
使用 lark-cli docs 命令操作
"""
import json
import shutil
import subprocess
import sys
from typing import Optional

LARK_CLI = shutil.which("lark-cli") or "lark-cli"


def create_doc(title: str, folder_token: Optional[str] = None) -> Optional[str]:
    """新建飞书文档，返回 document_id"""
    content = f"<title>{title}</title>"
    cmd = [LARK_CLI, "docs", "+create", "--api-version", "v2", "--content", content]
    if folder_token:
        cmd += ["--folder-token", folder_token]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", shell=(sys.platform == "win32"))
    if result.returncode != 0:
        print(f"[ERROR] 创建文档失败: {result.stderr}", file=sys.stderr)
        return None
    data = json.loads(result.stdout)
    if not data.get("ok"):
        print(f"[ERROR] 创建文档 API 错误: {data}", file=sys.stderr)
        return None
    return data["data"]["document"]["document_id"]


def append_content(doc_id: str, xml_content: str) -> bool:
    """向文档末尾追加内容"""
    cmd = [
        LARK_CLI, "docs", "+update", "--api-version", "v2",
        "--doc", doc_id,
        "--command", "append",
        "--content", xml_content,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", shell=(sys.platform == "win32"))
    if result.returncode != 0:
        print(f"[ERROR] 追加内容失败: {result.stderr}", file=sys.stderr)
        return False
    data = json.loads(result.stdout)
    return data.get("ok", False)


def build_heading_xml(text: str, level: int = 3) -> str:
    return f"<h{level}>{text}</h{level}>"


def build_callout_xml(lines: list[str], emoji: str = "❗") -> str:
    """构建 callout 块 XML"""
    inner = ""
    for line in lines:
        if line.startswith("**") or line.startswith("——"):
            inner += f"<p><b>{line}</b></p>"
        else:
            inner += f"<p>{line}</p>"
    return (
        f'<callout background-color="rgb(255,245,235)" '
        f'border-color="rgb(254,212,164)" emoji="{emoji}">'
        f"{inner}</callout>"
    )


def build_table_xml(headers: list[str], rows: list[list]) -> str:
    """构建表格 XML"""
    col_count = len(headers)
    row_count = len(rows) + 1
    xml = f'<table row-count="{row_count}" col-count="{col_count}">'
    # header row
    xml += "<tr>"
    for h in headers:
        xml += f"<td><p><b>{h}</b></p></td>"
    xml += "</tr>"
    # data rows
    for row in rows:
        xml += "<tr>"
        for i, cell in enumerate(row):
            val = _format_cell(cell)
            xml += f"<td><p>{val}</p></td>"
        xml += "</tr>"
    xml += "</table>"
    return xml


def _format_cell(value) -> str:
    if value is None:
        return ""
    if isinstance(value, float):
        if 0 < abs(value) < 1:
            return f"{value*100:.2f}%"
        if abs(value) >= 1000:
            return f"{value:,.0f}"
        return f"{value:.2f}"
    return str(value)


def write_section(
    doc_id: str,
    title: str,
    conclusion_lines: list[str],
    table_headers: list[str],
    table_rows: list[list],
    prev_table_headers: Optional[list[str]] = None,
    prev_table_rows: Optional[list[list]] = None,
) -> bool:
    """写入一个完整的周报子章节"""
    xml_parts = []
    xml_parts.append(build_heading_xml(title, level=6))
    xml_parts.append(build_callout_xml(conclusion_lines))
    if table_headers and table_rows:
        xml_parts.append(build_table_xml(table_headers, table_rows))
    if prev_table_headers and prev_table_rows:
        xml_parts.append("<p>#上周数据</p>")
        xml_parts.append(build_table_xml(prev_table_headers, prev_table_rows))
    full_xml = "".join(xml_parts)
    return append_content(doc_id, full_xml)
