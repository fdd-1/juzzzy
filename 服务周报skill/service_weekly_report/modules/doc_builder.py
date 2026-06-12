"""创建飞书文档并嵌入电子表格 (4.1 板块)
- 调用 lark-cli docs +create 创建文档
- 用 XML 格式: <h3>4.1 ...</h3> + callout + <sheet token=... sheet-id=...></sheet>
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

# 飞书云盘文件夹: 周报自动化
FOLDER_TOKEN = "CkpmfbJfTlWwx6d98PscfdOnnoe"


def lark_cli(cmd_args: list[str], timeout: int = 60) -> dict:
    """执行 lark-cli 命令并返回 JSON。"""
    cmd = [LARK_CLI] + cmd_args
    print(f"  $ {' '.join(cmd[:6])}...")
    result = subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        shell=(sys.platform == "win32"), timeout=timeout
    )
    if result.returncode != 0:
        if result.returncode == 10:
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
            print(f"  stderr: {result.stderr[:500]}")
            return {"ok": False, "stderr": result.stderr}
    try:
        return json.loads(result.stdout)
    except Exception as e:
        return {"ok": False, "raw": result.stdout, "error": str(e)}


def build_doc_xml(start_date: date, end_date: date, sheet_token: str, sheet_id: str, callout_xml: str) -> str:
    """构造文档 XML 结构: 标题 + h3 + callout + h6 + sheet 嵌入。"""
    title = f"服务周报 ({start_date.strftime('%m.%d')}-{end_date.strftime('%m.%d')})"
    xml = f"""<title>{title}</title>
<h3>4.1 服务指标跟进 &amp; 语义分析</h3>
{callout_xml}
<h6>服务指标数据表</h6>
<sheet token="{sheet_token}" sheet-id="{sheet_id}"></sheet>
"""
    return xml


def create_feishu_doc(start_date: date, end_date: date, sheet_token: str, sheet_id: str, callout_xml: str) -> dict:
    """创建飞书文档并写入内容。返回 {doc_id, url}。"""
    xml_content = build_doc_xml(start_date, end_date, sheet_token, sheet_id, callout_xml)

    # 写入临时文件 (用相对路径调用)
    tmp_dir = Path(__file__).parent.parent / "exports" / "_tmp"
    tmp_dir.mkdir(parents=True, exist_ok=True)
    tmp_file_name = f"doc_4_1_{start_date.strftime('%Y%m%d')}.xml"
    tmp_file = tmp_dir / tmp_file_name
    tmp_file.write_text(xml_content, encoding="utf-8")

    print(f"\n=== 创建飞书文档 ===")
    print(f"  XML 临时文件: {tmp_file}")

    # 切到 tmp 目录,用相对路径
    cmd = [
        LARK_CLI, "docs", "+create",
        "--api-version", "v2",
        "--folder-token", FOLDER_TOKEN,
        "--content", f"@{tmp_file_name}",
    ]
    print(f"  $ (cwd={tmp_dir}) {' '.join(cmd[:6])}...")
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
        print(f"  [错误] 找不到 document_id, response: {json.dumps(data, ensure_ascii=False)[:300]}")
        return {}

    url = f"https://hcnig43mb8gp.feishu.cn/docx/{doc_id}"
    print(f"  ✓ 文档创建成功")
    print(f"  doc_id: {doc_id}")
    print(f"  URL: {url}")

    return {"doc_id": doc_id, "url": url}


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--sheet-token", required=True)
    parser.add_argument("--sheet-id", required=True)
    parser.add_argument("--callout-file", required=True, help="callout XML 文件路径")
    parser.add_argument("--start-date", required=True)
    parser.add_argument("--end-date", required=True)
    args = parser.parse_args()

    from datetime import datetime
    start = datetime.strptime(args.start_date, "%Y-%m-%d").date()
    end = datetime.strptime(args.end_date, "%Y-%m-%d").date()

    callout = Path(args.callout_file).read_text(encoding="utf-8")
    result = create_feishu_doc(start, end, args.sheet_token, args.sheet_id, callout)
    print(f"\n=== 结果 ===")
    print(json.dumps(result, ensure_ascii=False, indent=2))
