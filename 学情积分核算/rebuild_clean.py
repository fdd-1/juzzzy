"""完整重做：清洗 sheet → 删尾 → 表头加粗 → 列宽 → 重写文档（参考样板格式）"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

EXTRACTED = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json")
SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
DOC_ID = "TBX2dTCIqoc3zJxHVdXcA1bUnoh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"

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

# 当前每个 sheet 在飞书侧的总行数（来自之前的 +info）
CURRENT_ROWS = {
    "4.1 服务指标": 200, "4.1 AI学情": 200, "4.2 组班意向": 200,
    "4.3 群发消息": 200, "4.4 停课监控": 200, "4.5 服务池跟进": 200,
    "4.5 服务池SOP": 200, "4.6 系统外呼监控": 200, "4.6 企微回复比": 297,
}

# 表头行数（前 N 行是表头，剩下是数据）
HEADER_ROWS = {
    "4.1 服务指标": 2, "4.1 AI学情": 2, "4.2 组班意向": 2,
    "4.3 群发消息": 2, "4.4 停课监控": 3, "4.5 服务池跟进": 2,
    "4.5 服务池SOP": 2, "4.6 系统外呼监控": 2, "4.6 企微回复比": 2,
}


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def safe_text(v):
    if isinstance(v, (int, float)):
        return v
    s = str(v)
    return s.replace("<", "＜").replace(">", "＞")


def run(cmd, stdin_text=None, timeout=180):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=(sys.platform == "win32"), timeout=timeout, input=stdin_text)
    if r.returncode != 0:
        return False, r.stderr or r.stdout
    try:
        return True, json.loads(r.stdout)
    except Exception:
        return False, r.stdout


def find_caliber_start(rows):
    for i, r in enumerate(rows):
        first = str(r[0]) if r else ""
        if any(kw in first for kw in ["口径说明", "注意：", "说明："]):
            return i
        if re.match(r"^\d+[、）)\.]\s*[^,]+[:：]", first):
            other = [str(c) for c in r[1:] if str(c).strip()]
            if len(other) <= 2:
                return i
    return len(rows)


def clean_rows(rows):
    cut = find_caliber_start(rows)
    rows = rows[:cut]
    # 去空行
    rows = [r for r in rows if any(str(c).strip() for c in r)]
    return rows


def write_rows_safe(sid, rows, batch_size=8):
    if not rows:
        return 0
    n_cols = max(len(r) for r in rows)
    rows = [list(r) + [""] * (n_cols - len(r)) for r in rows]
    rows = [[safe_text(c) for c in r] for r in rows]
    end_col = col_letter(n_cols)
    written = 0

    def write_batch(batch, start_row):
        end_row = start_row + len(batch) - 1
        rng = f"{sid}!A{start_row}:{end_col}{end_row}"
        values = json.dumps(batch, ensure_ascii=False)
        if len(values) > 7000:
            if len(batch) == 1:
                # 列分段
                row = batch[0]
                chunk_c = 15
                i = 0
                while i < len(row):
                    sub = row[i:i + chunk_c]
                    sj = json.dumps([sub], ensure_ascii=False)
                    if len(sj) > 7000 and chunk_c > 1:
                        chunk_c = max(1, chunk_c // 2)
                        continue
                    rg = f"{sid}!{col_letter(i+1)}{start_row}:{col_letter(i+len(sub))}{start_row}"
                    ok, _ = run([LARK_CLI, "sheets", "+write",
                                 "--spreadsheet-token", SPREADSHEET_TOKEN,
                                 "--range", rg, "--values", sj])
                    if not ok:
                        time.sleep(1.0)
                        ok, _ = run([LARK_CLI, "sheets", "+write",
                                     "--spreadsheet-token", SPREADSHEET_TOKEN,
                                     "--range", rg, "--values", sj])
                    i += chunk_c
                    time.sleep(0.15)
                return True
            half = len(batch) // 2
            return write_batch(batch[:half], start_row) and write_batch(batch[half:], start_row + half)
        ok, _ = run([LARK_CLI, "sheets", "+write",
                     "--spreadsheet-token", SPREADSHEET_TOKEN,
                     "--range", rng, "--values", values])
        if not ok:
            time.sleep(1.0)
            ok, _ = run([LARK_CLI, "sheets", "+write",
                         "--spreadsheet-token", SPREADSHEET_TOKEN,
                         "--range", rng, "--values", values])
        return ok

    i = 0
    while i < len(rows):
        batch = rows[i:i + batch_size]
        if write_batch(batch, i + 1):
            written += len(batch)
        i += batch_size
        time.sleep(0.2)
    return written


def delete_extra_rows(sid, keep_rows, current_rows):
    """删除从 keep_rows+1 到 current_rows 的多余行"""
    if current_rows <= keep_rows:
        return
    ok, resp = run([LARK_CLI, "sheets", "+delete-dimension",
                    "--spreadsheet-token", SPREADSHEET_TOKEN,
                    "--sheet-id", sid, "--dimension", "ROWS",
                    "--start-index", str(keep_rows + 1), "--end-index", str(current_rows)])
    print(f"  delete rows {keep_rows+1}-{current_rows}: {'OK' if ok else 'FAIL ' + str(resp)[:200]}")


def style_header(sid, n_header_rows, n_cols):
    """表头加粗 + 浅蓝背景 + 居中"""
    end_col = col_letter(n_cols)
    rng = f"{sid}!A1:{end_col}{n_header_rows}"
    style = json.dumps({
        "font": {"bold": True, "fontSize": 10},
        "backColor": "#E1F0FE",
        "hAlign": 1,  # 居中
        "vAlign": 1,
        "borderType": "FULL_BORDER",
    }, ensure_ascii=False)
    ok, resp = run([LARK_CLI, "sheets", "+set-style",
                    "--spreadsheet-token", SPREADSHEET_TOKEN,
                    "--range", rng, "--style", style])
    print(f"  style header rows 1-{n_header_rows}: {'OK' if ok else 'FAIL ' + str(resp)[:150]}")


def style_total_rows(sid, total_row_indices, n_cols):
    """汇总行（总计行）加粗 + 浅黄背景"""
    if not total_row_indices:
        return
    end_col = col_letter(n_cols)
    style = json.dumps({
        "font": {"bold": True},
        "backColor": "#FFF8DC",
    }, ensure_ascii=False)
    for r in total_row_indices:
        rng = f"{sid}!A{r}:{end_col}{r}"
        run([LARK_CLI, "sheets", "+set-style",
             "--spreadsheet-token", SPREADSHEET_TOKEN,
             "--range", rng, "--style", style])
        time.sleep(0.15)


def set_column_widths(sid, n_cols):
    """统一列宽：前 2 列稍宽，其余 110px"""
    # 1-2 列 130px
    run([LARK_CLI, "sheets", "+update-dimension",
         "--spreadsheet-token", SPREADSHEET_TOKEN,
         "--sheet-id", sid, "--dimension", "COLUMNS",
         "--start-index", "1", "--end-index", "2", "--fixed-size", "130"])
    time.sleep(0.2)
    # 其余 110px
    if n_cols >= 3:
        run([LARK_CLI, "sheets", "+update-dimension",
             "--spreadsheet-token", SPREADSHEET_TOKEN,
             "--sheet-id", sid, "--dimension", "COLUMNS",
             "--start-index", "3", "--end-index", str(n_cols), "--fixed-size", "110"])
    time.sleep(0.2)


# ===== 文档结论文本（按参考样板格式重写） =====

CONCLUSIONS = {
    "4.1": {
        "emoji": "❗",
        "title": "4.1 服务指标跟进 ＆ 语义分析",
        "paragraphs": [
            ("b", "跟进："),
            ("b", "首通：总体跟进率 97.82%；及时跟进率 93.32%，未达本月目标（95%）"),
            ("p", "美澳1组及时跟进率 84.0%，需注意（王智文 57% / 董筠宜 60% / 刘明飞 88% / 孙鹏飞 89%）"),
            ("p", "港澳1组及时跟进率 91.9%，需注意（温宇豪 50% / 王凯智 80% / 罗思情 83% / 陈慧妍 83%）"),
            ("p", "美澳3组及时跟进率 92.3%，需注意（韩明昭 80% / 张弛 83% / 田淇 83% / 敖勇 88%）"),
            ("p", "港澳组及时跟进率 93.1%，需注意（阮忻妍 0% / 王新宇 82% / 尤鹤 88% / 郭欣怡 89%）"),
            ("b", "首课：总体跟进率 95.44%；及时跟进率 89.53%，达到本月目标（85%）"),
            ("p", "美澳4组及时跟进率 79.2%，需注意（孙玮 67% / 于卓言 75% / 李鑫 80% / 邓翔睿 80%）"),
            ("p", "美澳3组及时跟进率 81.7%，需注意（曲慧 50% / 敖勇 57% / 刘彤 67% / 韩明昭 67%）"),
            ("p", "美澳2组及时跟进率 82.9%，需注意（周子雅 0% / 石晓杰 40% / 冷林鑫 50% / 梁媛 75%）"),
            ("b", "首专：总体跟进率 89.20%；及时跟进率 81.02%，未达本月目标（85%）"),
            ("p", "美澳4组及时跟进率仅 51.7%，需注意（林聪 0% / 董思雨 0% / 邓翔睿 33% / 李鑫 40%）"),
            ("p", "美澳5组及时跟进率 55.6%，需注意（高庆兰 47% / 刘博 57% / 李欣旭 67%）"),
            ("p", "美澳1组及时跟进率 70.2%，需注意（王小苗 20% / 王智文 50% / 孙浩 67%）"),
            ("p", "—— ＜b＞企微绑定率＜/b＞ 整体为 ＜b＞67.8%＜/b＞"),
            ("b", "语义分析："),
            ("p", "—— 添加企微/WS/Line 执行率 42.6%"),
            ("p", "—— 一家多娃问询执行率 77%；转介绍执行率 73%"),
            ("p", "—— 美澳2组邀请添加企微/一家多娃问询/转介绍执行率仅 0% / 42% / 42%（梁媛/冷林鑫/石晓杰/李松榕执行率为 0）"),
            ("b", "AI 学情助手跟进（港澳2组 / 美澳5组）："),
            ("p", "—— 港澳2组：王一强 / 田淇 / 刘宇 首课 AI 学情干预中学员 100% 未干预，需提醒跟进"),
            ("p", "—— 美澳5组：首专未干预学员较多，主管需提醒 LP 跟进"),
        ],
    },
    "4.2": {
        "emoji": "📋",
        "title": "4.2 组班多意向占比",
        "paragraphs": [
            ("b", "多意向占比 57.1%，达预期目标"),
            ("p", "当前意向等待学员 14 人，其中 2 个意向及以上 8 人"),
            ("p", "—— 多意向占比高于整体的小组：台湾组（100%）/ 美澳2组（100%）/ 美澳4组（100%）/ 港澳1组（75%）"),
            ("p", "—— 多意向占比低于整体的小组：美澳1组（33.3%）/ 港澳组（0%）/ 美澳3组（0%）"),
            ("b", "分场景占比（海外团队整体）："),
            ("p", "—— 复课·首消 M1-M3：等待 2 人，多意向占比 100%，人均等待 8.5 天"),
            ("p", "—— 复课·首消 M4-M12+：等待 7 人，多意向占比 71.4%，人均等待 6.3 天，等待 ＞14 天 1 人"),
            ("p", "—— 调课·首消 M1-M3：等待 3 人，多意向占比 33.3%，人均等待 17.7 天，等待 ＞14 天 3 人"),
            ("p", "—— 调课·首消 M4-M12+：等待 2 人，多意向占比 0%，人均等待 6.5 天"),
        ],
    },
    "4.3": {
        "emoji": "📢",
        "title": "4.3 群发消息汇总",
        "paragraphs": [
            ("b", "营销类群发：消息数 71，环比 ↑20.3%，覆盖 3,975 学员"),
            ("b", "服务类群发：消息数 39，环比 ↑39.3%，覆盖 4,576 学员"),
            ("p", "—— 个人群发占比 50.7%（个人群发占比高，说明 LP 个人主动触达多）"),
        ],
    },
    "4.4": {
        "emoji": "❗",
        "title": "4.4 停课学员执行监控",
        "paragraphs": [
            ("b", "当前整体停课占比 6.84%，超出本月目标（6%）"),
            ("p", "当前实际停课学员 1,718 人，执行中学员停课率 17.05%"),
            ("b", "分组停课占比表现："),
            ("p", "—— 美澳3组 8.58% / 美澳1组 8.36% / 美澳2组 7.99%，超目标较多，主管需重点关注"),
            ("p", "—— 港澳组 6.57% 略超目标；美澳4组 6.00% 刚好达标；其余组（港澳1/2组、美澳5组、台湾组）控制在 6% 以内"),
            ("b", "执行中学员停课率（停课唤醒缺口）："),
            ("p", "—— 港澳1组 29.3% / 港澳组 25.6% / 台湾组 20.6%，唤醒推进偏慢，需提醒 LP 加强唤醒目标学员的跟进"),
            ("p", "—— 美澳5组 7.1% / 美澳3组 10.2%，唤醒推进较好"),
        ],
    },
    "4.5": {
        "emoji": "❗",
        "title": "4.5 服务月跟进",
        "paragraphs": [
            ("b", "服务池跟进："),
            ("p", "整体服务池学员 2,130 人，外呼跟进率 93.85%，综合有效跟进率 36.90%；微信覆盖率 85.21%，微信有效回复率 29.34%"),
            ("p", "—— 美澳3组综合有效跟进率仅 31.6%，外呼跟进率 96.4% 但有效跟进偏低，需关注通话质量"),
            ("p", "—— 美澳2组 33.3%，美澳1组 34.4%，均低于整体"),
            ("p", "—— 港澳组 57.2% / 美澳5组 46.7% / 港澳1组 38.6%，表现相对较好"),
            ("p", "整体生均外呼次数 1.76，带 R 数 52，带 R 效率 2.44%，服务池转化端有提升空间"),
            ("b", "服务池语义执行（SOP）："),
            ("p", "整体语义点加和 1.82，未达服务池语义点加和目标（2.4）；命中服务池学员数 140 人"),
            ("p", "—— 分项执行率：学情反馈 58.6% / 学习规划 39.3% / 转介绍权益 58.6% / 告知剩余课时 25.7%（最低，需重点提升）"),
            ("p", "—— 台湾组 0 / 港澳1组 0.63 / 美澳5组 1.0 / 美澳4组 1.22，语义执行严重落后"),
            ("p", "—— 港澳组 2.72 / 美澳3组 2.46，已达目标，可作为标杆"),
        ],
    },
    "4.6": {
        "emoji": "❗",
        "title": "4.6 系统电话外呼 ＆ 微信回复监控",
        "paragraphs": [
            ("b", "整体系统外呼："),
            ("p", "全部活跃学员 22,945 人，外呼覆盖率 51.70%（月环比 -0.65%），生均呼次 2.88（月环比 +3.04%）"),
            ("p", "外呼接通率 69.90%，有效接通率 18.14%"),
            ("b", "分池子表现："),
            ("p", "—— 续费池 90.25% / 服务池 94.95%（月环比 +3.23%，提升明显）/ M1-M3 93.33% / 上月底续费 54.22% / 其他非做工池 23.83%（最低）"),
            ("p", "—— 美澳3组 41.3% / 美澳1组 40.4% / 美澳4组 43.5%，明显低于均值，需要主管督促 LP 提高活跃学员触达"),
            ("p", "—— 美澳5组 95.6% / 港澳1组 68.1% / 港澳2组 66.7%，整体覆盖率较好"),
            ("b", "整体微信发送 ＆ 回复比（口径：回复比 = 发送数 / 回复数，数值越高表示用户回复越少）："),
            ("p", "整体发送消息 848,633 条（生均 36.84），整体回复比 8.90"),
            ("p", "—— 系统推送 325,809 条（生均 14.15，回复比 3.42，用户回复最频密）"),
            ("p", "—— LP 个人发送 522,824 条（生均 22.70，回复比 5.48，居中）"),
            ("p", "—— 分池子：服务池 12.17（最稀疏，互动最弱）/ 续费池 8.41（一般）"),
            ("p", "待办：复盘高回复比组的话术，参考低回复比（系统推送类）的写法做模板优化"),
        ],
    },
}

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

SECTION_SHEETS = {
    "4.1": ["4.1 服务指标", "4.1 AI学情"],
    "4.2": ["4.2 组班意向"],
    "4.3": ["4.3 群发消息"],
    "4.4": ["4.4 停课监控"],
    "4.5": ["4.5 服务池跟进", "4.5 服务池SOP"],
    "4.6": ["4.6 系统外呼监控", "4.6 企微回复比"],
}


def xml_escape(s):
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def render_paragraph(kind, text):
    # 全角 ＜b＞ ＜/b＞ 占位符 → 真正的 b 标签
    text_escaped = xml_escape(text).replace("＜b＞", "<b>").replace("＜/b＞", "</b>")
    if kind == "b":
        return f"<p><b>{text_escaped}</b></p>"
    return f"<p>{text_escaped}</p>"


def build_doc_xml():
    parts = ["<docx>",
             "<h1>0528 服务周报（数据范围 5.01-5.24）</h1>",
             "<p>本文档由 Excel 数据 + HTML 周报结论自动生成。每节顺序：「结论 → 数据表」，每个数据表内嵌为独立 sheet 页签，数据已剔除口径行与空行，仅保留团队/LP 数据。</p>"]
    for sec_id in ["4.1", "4.2", "4.3", "4.4", "4.5", "4.6"]:
        sec = CONCLUSIONS[sec_id]
        parts.append(f"<h3>{xml_escape(sec['title'])}</h3>")
        parts.append(f'<callout emoji="{sec["emoji"]}">')
        parts.append("<p><b>结论</b></p>")
        for kind, text in sec["paragraphs"]:
            parts.append(render_paragraph(kind, text))
        parts.append("</callout>")
        for excel_name in SECTION_SHEETS[sec_id]:
            display = SHEET_DISPLAY[excel_name]
            sid = SHEET_ID_MAP[excel_name]
            parts.append(f"<h6>{xml_escape(display)}</h6>")
            parts.append(f'<sheet sheet-id="{sid}" token="{SPREADSHEET_TOKEN}"></sheet>')
    parts.append("</docx>")
    return "".join(parts)


def main():
    extracted = json.loads(EXTRACTED.read_text(encoding="utf-8"))
    excel_data_map = {s["title"]: s["rows"] for s in extracted["sheets"]}

    # === Stage 1: 清洗 + 写入每个 sheet ===
    print("[1/4] 清洗并重写每个 sheet 的数据...")
    cleaned_lengths = {}
    for excel_name, sid in SHEET_ID_MAP.items():
        rows = excel_data_map[excel_name]
        cleaned = clean_rows(rows)
        n = write_rows_safe(sid, cleaned, batch_size=8)
        cleaned_lengths[excel_name] = len(cleaned)
        print(f"  ✓ {excel_name} (sid={sid}): 写入 {n}/{len(cleaned)} 行 (原 {len(rows)} 行)")

    # === Stage 2: 删除多余尾行 ===
    print("\n[2/4] 删除多余尾行...")
    for excel_name, sid in SHEET_ID_MAP.items():
        keep = cleaned_lengths[excel_name]
        cur = CURRENT_ROWS[excel_name]
        if cur > keep:
            delete_extra_rows(sid, keep, cur)
        time.sleep(0.4)

    # === Stage 3: 表头加粗 + 列宽 + 总计行高亮 ===
    print("\n[3/4] 表头加粗 + 列宽设置 + 总计行高亮...")
    for excel_name, sid in SHEET_ID_MAP.items():
        rows = clean_rows(excel_data_map[excel_name])
        n_cols = max(len(r) for r in rows) if rows else 1
        n_header = HEADER_ROWS[excel_name]
        print(f"  -- {excel_name} (sid={sid}, {n_cols} cols, {n_header} header rows) --")
        style_header(sid, n_header, n_cols)
        # 找总计行（"总计" 通常在第 2 或 3 列）
        total_indices = []
        for i, r in enumerate(rows):
            if i < n_header:
                continue
            cells = [str(c).strip() for c in r[:4]]
            if "总计" in cells:
                total_indices.append(i + 1)
        if total_indices:
            print(f"     总计行: {total_indices[:10]}{'...' if len(total_indices) > 10 else ''}")
            style_total_rows(sid, total_indices, n_cols)
        set_column_widths(sid, n_cols)
        time.sleep(0.4)

    # === Stage 4: 重写文档 ===
    print("\n[4/4] 重写飞书文档...")
    xml = build_doc_xml()
    print(f"  XML 长度: {len(xml)} 字符")
    ok, resp = run([LARK_CLI, "docs", "+update", "--api-version", "v2",
                    "--doc", DOC_ID, "--command", "overwrite",
                    "--doc-format", "xml", "--content", "-"],
                   stdin_text=xml, timeout=300)
    if not ok:
        print(f"  [ERR] overwrite: {str(resp)[:600]}")
    else:
        data = resp.get("data", {})
        print(f"  result={data.get('result')}, warnings={data.get('warnings')}")

    print("\n[DONE]")
    print(f"文档：https://hcnig43mb8gp.feishu.cn/docx/{DOC_ID}")
    print(f"表格：https://hcnig43mb8gp.feishu.cn/sheets/{SPREADSHEET_TOKEN}")


if __name__ == "__main__":
    main()
