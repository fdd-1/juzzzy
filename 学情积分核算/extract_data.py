"""提取 HTML 周报中 4.1-4.6 的结论文本，以及 v6.xlsx 的 9 个 sheet 数据，输出为 JSON。"""
from __future__ import annotations
import sys, io, json, re
from pathlib import Path

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

import openpyxl
from bs4 import BeautifulSoup

ROOT = Path(r"C:\Users\fengjianyi\Desktop\周报自动化\20260528_服务")
HTML_PATH = ROOT / "周报_v4_5_01_5_24.html"
XLSX_PATH = ROOT / "周报数据汇总_v6.xlsx"
OUT_PATH = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json")


def strip_html(html_frag: str) -> str:
    soup = BeautifulSoup(html_frag, "html.parser")
    for br in soup.find_all("br"):
        br.replace_with("\n")
    text = soup.get_text("\n")
    lines = [ln.strip() for ln in text.splitlines() if ln.strip()]
    return "\n".join(lines)


def extract_conclusions():
    html = HTML_PATH.read_text(encoding="utf-8")
    soup = BeautifulSoup(html, "html.parser")
    sections = {}
    for h2 in soup.find_all("h2"):
        title = h2.get_text(" ", strip=True)
        m = re.match(r"^(4\.\d+)\s*(.*)", title)
        if not m:
            continue
        sec_id = m.group(1)
        sec_title = m.group(2).strip()
        # 找到此 h2 之后到下一个 4.x h2 之间的所有内容
        nodes = []
        for sib in h2.next_siblings:
            if sib.name == "h2":
                t2 = sib.get_text(" ", strip=True)
                if re.match(r"^4\.\d+", t2):
                    break
            nodes.append(sib)
        # 从 nodes 中提取首个"结论"块下面的内容
        conclusion_text = []
        capture = False
        stop = False
        for n in nodes:
            if stop:
                break
            if hasattr(n, "name") and n.name == "h2":
                heading = n.get_text(" ", strip=True)
                if heading == "结论":
                    capture = True
                    continue
                if capture and heading != "结论":
                    stop = True
                    break
            if capture and hasattr(n, "get_text"):
                txt = strip_html(str(n))
                if txt:
                    conclusion_text.append(txt)
        sections[sec_id] = {
            "title": sec_title,
            "conclusion": "\n".join(conclusion_text).strip(),
        }
    return sections


def excel_to_rows():
    wb = openpyxl.load_workbook(XLSX_PATH, data_only=True)
    sheets = []
    for ws in wb.worksheets:
        rows = []
        for row in ws.iter_rows(values_only=True):
            rows.append(["" if v is None else (v if isinstance(v, (int, float)) else str(v)) for v in row])
        # 去除完全空的尾行
        while rows and all(c == "" for c in rows[-1]):
            rows.pop()
        sheets.append({
            "title": ws.title,
            "rows": rows,
            "row_count": len(rows),
            "col_count": len(rows[0]) if rows else 0,
        })
    return sheets


def main():
    conclusions = extract_conclusions()
    sheets = excel_to_rows()
    out = {
        "conclusions": conclusions,
        "sheets": sheets,
    }
    OUT_PATH.write_text(json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"saved -> {OUT_PATH}")
    print("sections found:", list(conclusions.keys()))
    for s in sheets:
        print(f"  sheet [{s['title']}]: {s['row_count']} rows x {s['col_count']} cols")


if __name__ == "__main__":
    main()
