"""补写之前因命令行长度限制丢失的行；用 stdin 把 XML overwrite 到飞书文档。"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

EXTRACTED = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json")
SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
DOC_ID = "TBX2dTCIqoc3zJxHVdXcA1bUnoh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"

# 上一轮记录到的 sheet_id 映射
SHEET_ID_MAP = {
    "4.1 服务指标": "2RK2qc",
    "4.1 AI学情": "2S16wg",
    "4.2 组班意向": "2SgAAo",
    "4.3 群发消息": "2SwpXO",
    "4.4 停课监控": "2SOCo8",
    "4.5 服务池跟进": "2T3bYQ",
    "4.5 服务池SOP": "2Ti90Y",
    "4.6 系统外呼监控": "2TwScg",
    "4.6 企微回复比": "2TLJU4",
}


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def run(cmd, stdin_text=None, timeout=180):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=(sys.platform == "win32"), timeout=timeout,
                       input=stdin_text)
    if r.returncode != 0:
        return False, r.stderr or r.stdout
    try:
        return True, json.loads(r.stdout)
    except Exception:
        return False, r.stdout


def write_one_batch(sid, batch, start_row, n_cols):
    end_col = col_letter(n_cols)
    end_row = start_row + len(batch) - 1
    rng = f"{sid}!A{start_row}:{end_col}{end_row}"
    values_json = json.dumps(batch, ensure_ascii=False)
    if len(values_json) > 7500:
        # 拆半递归
        if len(batch) <= 1:
            # 单行还超长就强写（截断字符串）
            row = batch[0]
            row = [str(c)[:1000] if isinstance(c, str) and len(str(c)) > 1000 else c for c in row]
            return write_one_batch(sid, [row], start_row, n_cols)
        half = len(batch) // 2
        a = write_one_batch(sid, batch[:half], start_row, n_cols)
        b = write_one_batch(sid, batch[half:], start_row + half, n_cols)
        return a and b
    ok, resp = run([LARK_CLI, "sheets", "+write",
                    "--spreadsheet-token", SPREADSHEET_TOKEN,
                    "--range", rng, "--values", values_json])
    if not ok:
        msg = str(resp)[:200]
        # 重试一次（可能是限流）
        time.sleep(1.5)
        ok, resp = run([LARK_CLI, "sheets", "+write",
                        "--spreadsheet-token", SPREADSHEET_TOKEN,
                        "--range", rng, "--values", values_json])
        if not ok:
            print(f"    [ERR] rows {start_row}-{end_row}: {str(resp)[:200]}")
            return False
    time.sleep(0.25)
    return True


def write_full_sheet(sid, rows):
    """从第 1 行开始全量重写（覆盖之前残缺的写入）"""
    if not rows:
        return 0
    n_cols = max(len(r) for r in rows)
    rows = [r + [""] * (n_cols - len(r)) for r in rows]
    written = 0
    batch_size = 8  # 用更小批次
    i = 0
    while i < len(rows):
        batch = rows[i:i + batch_size]
        if write_one_batch(sid, batch, i + 1, n_cols):
            written += len(batch)
        i += batch_size
    return written


# ----- 4.6 改写后的结论 -----
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


SHEET_DISPLAY = {
    "4.1 服务指标": "服务指标数据表",
    "4.1 AI学情": "AI 学情助手数据表",
    "4.2 组班意向": "组班多意向数据表",
    "4.3 群发消息": "群发消息数据表",
    "4.4 停课监控": "停课学员执行监控数据表",
    "4.5 服务池跟进": "服务池跟进数据表",
    "4.5 服务池SOP": "服务池 SOP 语义执行数据表",
    "4.6 系统外呼监控": "LP 系统外呼监控（分池子）数据表",
    "4.6 企微回复比": "LP 企微回复比监控（分池子）数据表",
}


def xml_escape(s):
    return (s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;"))


def build_doc_xml(conclusions):
    parts = ["<docx>",
             "<h1>0528 服务周报（数据范围 5.01-5.24）</h1>",
             "<p>本文档由 Excel 数据 + HTML 周报结论自动生成，每个小节按「结论 → 数据表」顺序排列，每个数据表内嵌为独立 sheet 页签。</p>"]

    section_order = ["4.1", "4.2", "4.3", "4.4", "4.5", "4.6"]
    section_title = {
        "4.1": "4.1 服务指标跟进 & 语义分析",
        "4.2": "4.2 组班多意向占比",
        "4.3": "4.3 群发消息汇总",
        "4.4": "4.4 停课学员执行监控",
        "4.5": "4.5 服务月跟进",
        "4.6": "4.6 系统电话外呼 & 微信回复监控",
    }
    section_emoji = {"4.1": "❗", "4.2": "📋", "4.3": "📢", "4.4": "❗", "4.5": "❗", "4.6": "❗"}
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
        if sec == "4.6":
            conc = CONCLUSION_46_NEW
        else:
            conc = conclusions.get(sec, {}).get("conclusion", "")
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

        for excel_name in section_sheets[sec]:
            display_name = SHEET_DISPLAY[excel_name]
            sid = SHEET_ID_MAP[excel_name]
            parts.append(f"<h6>{xml_escape(display_name)}</h6>")
            parts.append(f'<sheet sheet-id="{sid}" token="{SPREADSHEET_TOKEN}"></sheet>')

    parts.append("</docx>")
    return "".join(parts)


def main():
    extracted = json.loads(EXTRACTED.read_text(encoding="utf-8"))
    conclusions = extracted["conclusions"]
    excel_data_map = {s["title"]: s["rows"] for s in extracted["sheets"]}

    # === 全量重写所有 sheet（确保数据完整）===
    print("[1/2] 全量重写所有 sheet（小批次，覆盖之前的残缺）...")
    for excel_name, sid in SHEET_ID_MAP.items():
        rows = excel_data_map[excel_name]
        # 数字保持数字
        cleaned = []
        for r in rows:
            cleaned_row = [v if isinstance(v, (int, float)) else str(v) for v in r]
            cleaned.append(cleaned_row)
        n = write_full_sheet(sid, cleaned)
        print(f"  ✓ {excel_name} ({sid}): {n}/{len(cleaned)} 行")

    # === overwrite 文档（用 stdin 传 XML，绕开路径限制）===
    print("\n[2/2] overwrite 文档（用 stdin 传 XML）...")
    xml = build_doc_xml(conclusions)
    print(f"  XML 长度: {len(xml)} 字符")
    ok, resp = run([LARK_CLI, "docs", "+update", "--api-version", "v2",
                    "--doc", DOC_ID, "--command", "overwrite",
                    "--doc-format", "xml",
                    "--content", "-"], stdin_text=xml, timeout=300)
    if not ok:
        print(f"  [ERR] overwrite: {str(resp)[:600]}")
        return
    data = resp.get("data", {})
    print(f"  result={data.get('result')}, warnings={data.get('warnings')}")

    print("\n[DONE]")
    print(f"文档：https://hcnig43mb8gp.feishu.cn/docx/{DOC_ID}")
    print(f"表格：https://hcnig43mb8gp.feishu.cn/sheets/{SPREADSHEET_TOKEN}")


if __name__ == "__main__":
    main()
