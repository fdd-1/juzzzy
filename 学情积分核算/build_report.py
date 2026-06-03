"""把 9 个 Excel sheet 数据写到飞书 spreadsheet，并重建飞书文档结构。

流程：
1. lark-cli sheets +info 拿当前所有 sheet（默认有一个 Sheet1）
2. 按顺序 +create-sheet 创建 9 个新 sheet（用 Excel 的 sheet 名）
3. 分批 +write 写入每个 sheet 数据
4. +delete-sheet 删掉默认 Sheet1
5. lark-cli docs +update --command overwrite 用新结构覆盖文档
"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

EXTRACTED = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json")
SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
DOC_ID = "TBX2dTCIqoc3zJxHVdXcA1bUnoh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"


def run(cmd, timeout=120):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=(sys.platform == "win32"), timeout=timeout)
    if r.returncode != 0:
        return False, r.stderr or r.stdout
    try:
        return True, json.loads(r.stdout)
    except Exception:
        return False, r.stdout


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def get_sheets():
    ok, resp = run([LARK_CLI, "sheets", "+info", "--spreadsheet-token", SPREADSHEET_TOKEN])
    if not ok:
        raise RuntimeError(f"info failed: {resp}")
    data = resp["data"]["sheets"]
    if isinstance(data, dict):
        data = data.get("sheets", [])
    return data


def create_sheet(title, idx):
    ok, resp = run([LARK_CLI, "sheets", "+create-sheet",
                    "--spreadsheet-token", SPREADSHEET_TOKEN,
                    "--title", title, "--index", str(idx)])
    if not ok:
        print(f"  [ERR] create-sheet {title}: {resp}")
        return None
    d = resp.get("data", {})
    return d.get("sheet_id") or d.get("replies", [{}])[0].get("addSheet", {}).get("properties", {}).get("sheetId")


def delete_sheet(sheet_id):
    ok, resp = run([LARK_CLI, "sheets", "+delete-sheet",
                    "--spreadsheet-token", SPREADSHEET_TOKEN,
                    "--sheet-id", sheet_id, "--yes"])
    print(f"  delete sheet_id={sheet_id}: {'OK' if ok else 'FAIL '+str(resp)[:200]}")
    return ok


def write_rows(sheet_id, rows, batch_size=20):
    if not rows:
        return 0
    n_cols = max(len(r) for r in rows)
    end_col = col_letter(n_cols)
    written = 0
    for i in range(0, len(rows), batch_size):
        batch = rows[i:i + batch_size]
        # 补齐列
        batch = [r + [""] * (n_cols - len(r)) for r in batch]
        start_row = i + 1
        end_row = i + len(batch)
        rng = f"{sheet_id}!A{start_row}:{end_col}{end_row}"
        values_json = json.dumps(batch, ensure_ascii=False)
        # Windows cmdline 长度限制：单批 JSON 太大就再切
        if len(values_json) > 7000 and batch_size > 1:
            half = max(1, batch_size // 2)
            return write_rows(sheet_id, rows, batch_size=half)
        ok, resp = run([LARK_CLI, "sheets", "+write",
                        "--spreadsheet-token", SPREADSHEET_TOKEN,
                        "--range", rng, "--values", values_json])
        if not ok:
            print(f"  [ERR] write rows {start_row}-{end_row}: {str(resp)[:300]}")
            time.sleep(1.0)
            continue
        written += len(batch)
        time.sleep(0.3)
    return written


# 4.6 重写后的结论文本（回复比越高 = 用户回复越少）
CONCLUSION_46_NEW = """整体系统外呼：
全部活跃学员 22945 人，外呼覆盖率 51.70%（月环比下降 0.65%），生均呼次 2.88（月环比上升 3.04%），外呼接通率 69.90%，有效接通率 18.14%
分池子表现：
续费池覆盖率 90.25%（最高，月环比 -0.72%）/ 服务池 94.95%（月环比 +3.23%，提升明显）/ M1-M3 93.33% / 上月底续费 54.22% / 其他非做工池 23.83%（最低）
—— 美澳3组整体覆盖率 41.3%、美澳1组 40.4%、美澳4组 43.5%，整体覆盖率明显低于均值（51.7%），需要主管督促 LP 提高活跃学员触达
—— 美澳5组 95.6% / 港澳1组 68.1% / 港澳2组 66.7%，整体覆盖率较好

整体微信发送 & 回复比（口径：回复比 = 发送数 / 回复数，数值越高表示用户回复越少）：
整体发送消息 848,633 条（生均 36.84），整体回复比 8.90
系统推送 325,809 条（生均 14.15，回复比 3.42，用户回复最频密）
LP 个人发送 522,824 条（生均 22.70，回复比 5.48，用户回复频度居中）
分池子：服务池消息 65,137 条（回复比 12.17，用户回复最稀疏，互动最弱）/ 续费池 271,029 条（回复比 8.41，互动一般）
—— 系统推送回复比最低（3.42），说明系统模板话术对学员的回应吸引力较强；LP 个人发送回复比 5.48，仍有优化空间
—— 服务池回复比 12.17（最稀疏），需要重点关注：服务池 LP 触达数量大但学员回应少，建议复盘话术内容、增加学员关心的真实价值点
—— 待办：抽样审视高回复比小组的群发话术，参考低回复比小组（系统推送类）的写法做模板优化"""


# Excel sheet 名 → 文档内显示名 + 节
SHEET_MAP = {
    "4.1 服务指标": ("服务指标数据表", "4.1"),
    "4.1 AI学情": ("AI 学情助手数据表", "4.1"),
    "4.2 组班意向": ("组班多意向数据表", "4.2"),
    "4.3 群发消息": ("群发消息数据表", "4.3"),
    "4.4 停课监控": ("停课学员执行监控数据表", "4.4"),
    "4.5 服务池跟进": ("服务池跟进数据表", "4.5"),
    "4.5 服务池SOP": ("服务池 SOP 语义执行数据表", "4.5"),
    "4.6 系统外呼监控": ("LP 系统外呼监控（分池子）数据表", "4.6"),
    "4.6 企微回复比": ("LP 企微回复比监控（分池子）数据表", "4.6"),
}


def xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def build_doc_xml(sheet_token_id_map, conclusions):
    """sheet_token_id_map: {excel_sheet_title: sheet_id}"""
    parts = []
    parts.append("<docx>")
    parts.append("<h1>0528 服务周报（数据范围 5.01-5.24）</h1>")
    parts.append("<p>本文档由 Excel 数据 + HTML 周报结论自动生成，每个小节按「结论 → 数据表」顺序排列，每个数据表内嵌为独立 sheet 页签。</p>")

    # 节顺序
    section_order = ["4.1", "4.2", "4.3", "4.4", "4.5", "4.6"]
    # 节标题（来自参考样板和 HTML）
    section_title = {
        "4.1": "4.1 服务指标跟进 & 语义分析",
        "4.2": "4.2 组班多意向占比",
        "4.3": "4.3 群发消息汇总",
        "4.4": "4.4 停课学员执行监控",
        "4.5": "4.5 服务月跟进",
        "4.6": "4.6 系统电话外呼 & 微信回复监控",
    }
    section_emoji = {
        "4.1": "❗", "4.2": "📋", "4.3": "📢",
        "4.4": "❗", "4.5": "❗", "4.6": "❗",
    }
    # 每节对应的 Excel sheet（顺序）
    section_sheets = {
        "4.1": ["4.1 服务指标", "4.1 AI学情"],
        "4.2": ["4.2 组班意向"],
        "4.3": ["4.3 群发消息"],
        "4.4": ["4.4 停课监控"],
        "4.5": ["4.5 服务池跟进", "4.5 服务池SOP"],
        "4.6": ["4.6 系统外呼监控", "4.6 企微回复比"],
    }

    for sec in section_order:
        parts.append(f"<h3>{xml_escape(section_title[sec])}</h3>")
        # 结论
        if sec == "4.6":
            conc = CONCLUSION_46_NEW
        else:
            conc = conclusions.get(sec, {}).get("conclusion", "")
            # 移除原结论开头的 emoji（因为我们用 callout emoji 属性了）
            lines = [ln for ln in conc.splitlines() if ln.strip() not in {"❗", "📋", "📢", "✏️", "💬"}]
            conc = "\n".join(lines)

        parts.append(f'<callout emoji="{section_emoji[sec]}">')
        parts.append("<p><b>结论</b></p>")
        for line in conc.split("\n"):
            line = line.strip()
            if not line:
                parts.append("<p></p>")
            else:
                parts.append(f"<p>{xml_escape(line)}</p>")
        parts.append("</callout>")

        # 数据表
        for excel_name in section_sheets[sec]:
            display_name, _ = SHEET_MAP[excel_name]
            sid = sheet_token_id_map[excel_name]
            parts.append(f"<h6>{xml_escape(display_name)}</h6>")
            parts.append(f'<sheet sheet-id="{sid}" token="{SPREADSHEET_TOKEN}"></sheet>')

    parts.append("</docx>")
    return "".join(parts)


def main():
    extracted = json.loads(EXTRACTED.read_text(encoding="utf-8"))
    conclusions = extracted["conclusions"]
    excel_sheets = extracted["sheets"]

    # === Step 1: 创建 9 个 sheet ===
    print("[1/4] 创建 9 个 sheet 页签...")
    excel_to_sheet_id = {}
    target_order = list(SHEET_MAP.keys())
    # 先获取当前 sheets（应该有默认 Sheet1）
    before = get_sheets()
    print(f"  已有 sheet: {[(s.get('title'), s.get('sheet_id')) for s in before]}")
    default_ids = [s.get("sheet_id") for s in before]

    for idx, excel_name in enumerate(target_order):
        # idx 从 0 开始，新建在末尾就让 idx=后续递增；这里固定用 idx + len(default)
        sid = create_sheet(excel_name, idx + len(default_ids))
        if not sid:
            print(f"  [ERR] 创建失败: {excel_name}")
            return
        excel_to_sheet_id[excel_name] = sid
        print(f"  + {excel_name} -> sheet_id={sid}")
        time.sleep(0.4)

    # === Step 2: 写入数据 ===
    print("\n[2/4] 写入数据...")
    excel_data_map = {s["title"]: s["rows"] for s in excel_sheets}
    for excel_name in target_order:
        rows = excel_data_map.get(excel_name, [])
        sid = excel_to_sheet_id[excel_name]
        # 清掉非字符串/数字的内容，防 lark-cli 报错
        cleaned = []
        for r in rows:
            cleaned_row = []
            for v in r:
                if isinstance(v, (int, float)):
                    cleaned_row.append(v)
                else:
                    cleaned_row.append(str(v))
            cleaned.append(cleaned_row)
        n = write_rows(sid, cleaned, batch_size=15)
        print(f"  ✓ {excel_name}: 写入 {n}/{len(cleaned)} 行")

    # === Step 3: 删默认 Sheet1 ===
    print("\n[3/4] 删除默认 Sheet1...")
    for sid in default_ids:
        delete_sheet(sid)
        time.sleep(0.5)

    # === Step 4: 重建文档 ===
    print("\n[4/4] 重建飞书文档...")
    xml = build_doc_xml(excel_to_sheet_id, conclusions)
    xml_path = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_doc.xml")
    xml_path.write_text(xml, encoding="utf-8")
    print(f"  XML saved: {xml_path} ({len(xml)} chars)")

    # 用 @file 方式传入超长 XML
    ok, resp = run([LARK_CLI, "docs", "+update", "--api-version", "v2",
                    "--doc", DOC_ID, "--command", "overwrite",
                    "--doc-format", "xml",
                    "--content", f"@{xml_path}"], timeout=300)
    if not ok:
        print(f"  [ERR] overwrite failed: {str(resp)[:500]}")
    else:
        result = resp.get("data", {}).get("result", "?")
        warnings = resp.get("data", {}).get("warnings", [])
        print(f"  result={result}, warnings={warnings}")

    # 输出结果
    out = {
        "doc_url": f"https://hcnig43mb8gp.feishu.cn/docx/{DOC_ID}",
        "sheet_url": f"https://hcnig43mb8gp.feishu.cn/sheets/{SPREADSHEET_TOKEN}",
        "sheet_id_map": excel_to_sheet_id,
    }
    Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_result.json").write_text(
        json.dumps(out, ensure_ascii=False, indent=2), encoding="utf-8")
    print("\n[DONE]")
    print(json.dumps(out, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
