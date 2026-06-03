"""综合修复脚本 - 处理所有 sheet 的格式问题"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

EXTRACTED = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json")
SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
DOC_ID = "TBX2dTCIqoc3zJxHVdXcA1bUnoh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"

SHEET_ID_MAP = {
    "4.1 服务指标": "2RK2qc", "4.1 AI学情": "2S16wg",
    "4.2 组班意向": "2SgAAo", "4.3 群发消息": "2SwpXO",
    "4.4 停课监控": "2SOCo8", "4.5 服务池跟进": "2T3bYQ",
    "4.5 服务池SOP": "2Ti90Y", "4.6 系统外呼监控": "2TwScg",
    "4.6 企微回复比": "2TLJU4",
}


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def run(cmd, timeout=120):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=(sys.platform == "win32"), timeout=timeout)
    if r.returncode != 0:
        return False, r.stderr or r.stdout
    try:
        return True, json.loads(r.stdout)
    except Exception:
        return False, r.stdout


def batch_set_style(data):
    dj = json.dumps(data, ensure_ascii=False)
    if len(dj) > 7500:
        half = len(data) // 2
        return batch_set_style(data[:half]) and batch_set_style(data[half:])
    ok, _ = run([LARK_CLI, "sheets", "+batch-set-style",
                 "--spreadsheet-token", SPREADSHEET_TOKEN, "--data", dj])
    return ok


def set_style(sid, rng, style):
    ok, _ = run([LARK_CLI, "sheets", "+set-style",
                 "--spreadsheet-token", SPREADSHEET_TOKEN,
                 "--range", f"{sid}!{rng}", "--style", json.dumps(style, ensure_ascii=False)])
    return ok


def merge_cells(sid, rng):
    ok, _ = run([LARK_CLI, "sheets", "+merge-cells",
                 "--spreadsheet-token", SPREADSHEET_TOKEN,
                 "--range", f"{sid}!{rng}", "--merge-type", "MERGE_ALL"])
    return ok


def delete_columns(sid, cols_to_delete):
    """从右到左删除列（避免索引偏移）"""
    for col in sorted(cols_to_delete, reverse=True):
        run([LARK_CLI, "sheets", "+delete-dimension",
             "--spreadsheet-token", SPREADSHEET_TOKEN,
             "--sheet-id", sid, "--dimension", "COLUMNS",
             "--start-index", str(col), "--end-index", str(col)])
        time.sleep(0.3)


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
    return [r for r in rows[:cut] if any(str(c).strip() for c in r)]


def safe_text(v):
    if isinstance(v, str):
        return v.replace("<", "＜").replace(">", "＞")
    return v


def to_float(v):
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("%", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def lerp_color(ratio):
    if ratio >= 0.5:
        t = (ratio - 0.5) * 2
        r = int(255 * (1 - t) + 76 * t)
        g = int(215 * (1 - t) + 175 * t)
        b = int(0 * (1 - t) + 80 * t)
    else:
        t = ratio * 2
        r = int(244 * (1 - t) + 255 * t)
        g = int(67 * (1 - t) + 215 * t)
        b = int(54 * (1 - t) + 0 * t)
    return f"#{r:02X}{g:02X}{b:02X}"


def write_batch(sid, batch, start_row, n_cols):
    end_col = col_letter(n_cols)
    end_row = start_row + len(batch) - 1
    rng = f"{sid}!A{start_row}:{end_col}{end_row}"
    values = json.dumps(batch, ensure_ascii=False)
    if len(values) > 7000:
        if len(batch) == 1:
            row = batch[0]
            i = 0
            chunk = 15
            while i < len(row):
                sub = row[i:i + chunk]
                sj = json.dumps([sub], ensure_ascii=False)
                if len(sj) > 7000 and chunk > 1:
                    chunk = max(1, chunk // 2)
                    continue
                rg = f"{sid}!{col_letter(i+1)}{start_row}:{col_letter(i+len(sub))}{start_row}"
                run([LARK_CLI, "sheets", "+write", "--spreadsheet-token",
                     SPREADSHEET_TOKEN, "--range", rg, "--values", sj])
                i += chunk
                time.sleep(0.15)
            return True
        half = len(batch) // 2
        return write_batch(sid, batch[:half], start_row, n_cols) and \
               write_batch(sid, batch[half:], start_row + half, n_cols)
    ok, _ = run([LARK_CLI, "sheets", "+write", "--spreadsheet-token",
                 SPREADSHEET_TOKEN, "--range", rng, "--values", values])
    if not ok:
        time.sleep(1.0)
        ok, _ = run([LARK_CLI, "sheets", "+write", "--spreadsheet-token",
                     SPREADSHEET_TOKEN, "--range", rng, "--values", values])
    return ok


def apply_color_scale(sid, col_indices, data_rows, n_header):
    """对指定列应用绿黄红色阶"""
    n_cols_max = max(len(r) for r in data_rows) if data_rows else 0
    ops = []
    for col_idx in col_indices:
        values = []
        for row_i, r in enumerate(data_rows):
            padded = list(r) + [""] * (n_cols_max - len(r))
            cell = padded[col_idx] if col_idx < len(padded) else ""
            v = to_float(cell)
            values.append((row_i, v))
        nums = [v for _, v in values if v is not None and v != 0]
        if not nums or max(nums) == min(nums):
            continue
        min_v, max_v = min(nums), max(nums)
        col_l = col_letter(col_idx + 1)
        for row_i, v in values:
            if v is None or v == 0:
                continue
            ratio = (v - min_v) / (max_v - min_v)
            color = lerp_color(ratio)
            actual_row = n_header + 1 + row_i
            rng = f"{sid}!{col_l}{actual_row}:{col_l}{actual_row}"
            ops.append({"ranges": [rng], "style": {"backColor": color}})
    sent = 0
    for i in range(0, len(ops), 30):
        batch = ops[i:i + 30]
        if batch_set_style(batch):
            sent += len(batch)
        time.sleep(0.3)
    return sent


def group_merge_col(sid, rows, col_idx, n_header):
    """对指定列做向下合并（相同值的连续行合并）"""
    col_l = col_letter(col_idx + 1)
    data_rows = rows[n_header:]
    cur_val = None
    cur_start = None
    merges = []
    for i, r in enumerate(data_rows):
        padded = list(r) + [""] * 5
        v = str(padded[col_idx]).strip()
        if v and v != cur_val:
            if cur_val and cur_start is not None and i - cur_start > 1:
                merges.append((cur_start + n_header + 1, i + n_header))
            cur_val = v
            cur_start = i
        elif v == "" or v == cur_val:
            pass
    if cur_val and cur_start is not None and len(data_rows) - cur_start > 1:
        merges.append((cur_start + n_header + 1, len(data_rows) + n_header))
    for s, e in merges:
        merge_cells(sid, f"{col_l}{s}:{col_l}{e}")
        time.sleep(0.15)
    return len(merges)


def main():
    extracted = json.loads(EXTRACTED.read_text(encoding="utf-8"))
    excel_data_map = {s["title"]: s["rows"] for s in extracted["sheets"]}

    # ========== Step 1: 4.4 删除微信关键词覆盖率和1V1推报率列 ==========
    print("[Step 1] 4.4 删除微信关键词覆盖率/1V1推报率列...")
    sid_44 = SHEET_ID_MAP["4.4 停课监控"]
    # 列号（1-based）：30,31,44,45,58,59,69,70,80,81
    cols_del = [30, 31, 44, 45, 58, 59, 69, 70, 80, 81]
    delete_columns(sid_44, cols_del)
    print(f"  删除了 {len(cols_del)} 列")

    # ========== Step 2: 清除所有 sheet 的数据区背景色 ==========
    print("\n[Step 2] 清除所有 sheet 数据区背景色...")
    for title, sid in SHEET_ID_MAP.items():
        rows = clean_rows(excel_data_map[title])
        if not rows:
            continue
        n_cols = max(len(r) for r in rows)
        n_header = 3 if title == "4.4 停课监控" else 2
        n_rows = len(rows)
        if title == "4.4 停课监控":
            n_cols -= len(cols_del)  # 删列后列数减少
        end_col = col_letter(n_cols)
        set_style(sid, f"A{n_header+1}:{end_col}{n_rows}", {"clean": True})
        time.sleep(0.3)
    print("  全部清除完成")

    # ========== Step 3: 4.1 修复执行率加和（改回小数）==========
    print("\n[Step 3] 4.1 修复执行率加和列...")
    sid_41 = SHEET_ID_MAP["4.1 服务指标"]
    rows_41 = clean_rows(excel_data_map["4.1 服务指标"])
    # col 4 和 col 16 是执行率加和，当前被错误转成了百分比
    # 重新写入这两列的数据行（row 3 开始）
    for col_idx in [3, 15]:  # 0-based
        col_l = col_letter(col_idx + 1)
        for i, r in enumerate(rows_41[2:], 3):  # 从第3行开始
            padded = list(r) + [""] * (27 - len(r))
            v = padded[col_idx]
            fv = to_float(v)
            if fv is not None:
                cell_val = f"{fv:.2f}"
            else:
                cell_val = safe_text(str(v))
            rng = f"{sid_41}!{col_l}{i}:{col_l}{i}"
            run([LARK_CLI, "sheets", "+write", "--spreadsheet-token",
                 SPREADSHEET_TOKEN, "--range", rng,
                 "--values", json.dumps([[cell_val]], ensure_ascii=False)])
            time.sleep(0.1)
    print("  col 4 + col 16 执行率加和已改回小数")

    # ========== Step 4: 4.1 首课语义点执行表头合并 ==========
    print("\n[Step 4] 4.1 首课语义点执行表头合并...")
    merge_cells(sid_41, "P1:Q1")  # col 16-17
    print("  P1:Q1 合并完成")

    # ========== Step 5: 4.3 基础表头合并 ==========
    print("\n[Step 5] 4.3 基础表头合并...")
    sid_43 = SHEET_ID_MAP["4.3 群发消息"]
    merge_cells(sid_43, "A1:B1")  # 基础 = 小组+LP
    # 其他表头向上合并（row1 的"其他"只有1列=col27）
    merge_cells(sid_43, "AA1:AA2")  # 其他 向下合并到 row2
    print("  A1:B1 + AA1:AA2 合并完成")

    # ========== Step 6: 4.4 表头向上合并 + 小组向下合并 ==========
    print("\n[Step 6] 4.4 表头合并 + 小组合并...")
    # 4.4 删列后需要重新读取当前 sheet 状态
    # 原 row1 和 row2 中相同的表头需要纵向合并
    # 由于删了列，我需要重新 fetch 当前 sheet 的表头
    # 简化：直接对 row1-row2 中相同值的列做纵向合并
    ok, info = run([LARK_CLI, "sheets", "+read", "--spreadsheet-token",
                    SPREADSHEET_TOKEN, "--range", f"{sid_44}!A1:BV3"])
    if ok:
        r1_data = info.get("data", {}).get("values", [[]])[0] if info.get("data") else []
        r2_data = info.get("data", {}).get("values", [[]])[1] if info.get("data") and len(info["data"].get("values", [])) > 1 else []
        # 纵向合并：row1 和 row2 值相同的列
        for ci in range(len(r1_data)):
            v1 = str(r1_data[ci]).strip() if ci < len(r1_data) and r1_data[ci] else ""
            v2 = str(r2_data[ci]).strip() if ci < len(r2_data) and r2_data[ci] else ""
            if v1 and v1 == v2:
                cl = col_letter(ci + 1)
                merge_cells(sid_44, f"{cl}1:{cl}2")
                time.sleep(0.1)
        print("  表头纵向合并完成")
    # 小组向下合并（col 1, 0-based=0）
    # 需要重新读取数据
    ok2, info2 = run([LARK_CLI, "sheets", "+read", "--spreadsheet-token",
                      SPREADSHEET_TOKEN, "--range", f"{sid_44}!A1:A99"])
    if ok2:
        col_data = [row[0] if row else "" for row in info2.get("data", {}).get("values", [])]
        n_h = 3
        cur_val = None
        cur_start = None
        merges = []
        for i in range(n_h, len(col_data)):
            v = str(col_data[i]).strip() if col_data[i] else ""
            if v and v != cur_val:
                if cur_val and cur_start is not None and i - cur_start > 1:
                    merges.append((cur_start + 1, i))
                cur_val = v
                cur_start = i
        if cur_val and cur_start is not None and len(col_data) - cur_start > 1:
            merges.append((cur_start + 1, len(col_data)))
        for s, e in merges:
            merge_cells(sid_44, f"A{s}:A{e}")
            time.sleep(0.15)
        print(f"  小组向下合并 {len(merges)} 组")

    # ========== Step 7: 4.5 仅保留外呼跟进率(col4)和综合有效跟进率(col9)色阶 ==========
    print("\n[Step 7] 4.5 重新应用色阶（仅外呼跟进率+综合有效跟进率）...")
    sid_45 = SHEET_ID_MAP["4.5 服务池跟进"]
    rows_45 = clean_rows(excel_data_map["4.5 服务池跟进"])
    data_45 = rows_45[2:]
    n_sent = apply_color_scale(sid_45, [3, 8], data_45, 2)  # 0-based: col4=idx3, col9=idx8
    print(f"  {n_sent} 单元格上色")

    # ========== Step 8: 4.4 仅唤醒率列色阶 ==========
    print("\n[Step 8] 4.4 唤醒率列色阶...")
    # 删列后唤醒率位置变了，需要重新定位
    ok3, info3 = run([LARK_CLI, "sheets", "+read", "--spreadsheet-token",
                      SPREADSHEET_TOKEN, "--range", f"{sid_44}!A3:BV3"])
    if ok3:
        r3_vals = info3.get("data", {}).get("values", [[]])[0]
        wake_cols = [i for i, v in enumerate(r3_vals) if v and "唤醒率" in str(v)]
        print(f"  唤醒率列（0-based）: {wake_cols}")
        # 读取数据行
        ok4, info4 = run([LARK_CLI, "sheets", "+read", "--spreadsheet-token",
                          SPREADSHEET_TOKEN, "--range", f"{sid_44}!A4:BV99"])
        if ok4:
            data_44 = info4.get("data", {}).get("values", [])
            n_sent2 = apply_color_scale(sid_44, wake_cols, data_44, 3)
            print(f"  {n_sent2} 单元格上色")

    # ========== Step 9: 4.6 小组向下合并 ==========
    print("\n[Step 9] 4.6 小组向下合并...")
    for title in ["4.6 系统外呼监控", "4.6 企微回复比"]:
        sid = SHEET_ID_MAP[title]
        ok5, info5 = run([LARK_CLI, "sheets", "+read", "--spreadsheet-token",
                          SPREADSHEET_TOKEN, "--range", f"{sid}!A1:A300"])
        if ok5:
            col_data = [row[0] if row else "" for row in info5.get("data", {}).get("values", [])]
            n_h = 2
            cur_val = None
            cur_start = None
            merges = []
            for i in range(n_h, len(col_data)):
                v = str(col_data[i]).strip() if col_data[i] else ""
                if v and v != cur_val:
                    if cur_val and cur_start is not None and i - cur_start > 1:
                        merges.append((cur_start + 1, i))
                    cur_val = v
                    cur_start = i
            if cur_val and cur_start is not None and len(col_data) - cur_start > 1:
                merges.append((cur_start + 1, len(col_data)))
            for s, e in merges:
                merge_cells(sid, f"A{s}:A{e}")
                time.sleep(0.15)
            print(f"  [{title}] 小组合并 {len(merges)} 组")

    # ========== Step 10: 重新上色（表头+总计行）==========
    print("\n[Step 10] 重新上色表头+总计行...")
    for title, sid in SHEET_ID_MAP.items():
        rows = clean_rows(excel_data_map[title])
        if not rows:
            continue
        n_cols = max(len(r) for r in rows)
        if title == "4.4 停课监控":
            n_cols -= len(cols_del)
        n_header = 3 if title == "4.4 停课监控" else 2
        n_rows = len(rows)
        end_col = col_letter(n_cols)
        ops = [{"ranges": [f"{sid}!A1:{end_col}{n_header}"],
                "style": {"font": {"bold": True, "font_size": 10},
                          "backColor": "#E1F0FE", "hAlign": 1, "vAlign": 1}}]
        total_idx = []
        for i, r in enumerate(rows):
            if i < n_header:
                continue
            cells4 = [str(c).strip() for c in r[:4]]
            if "总计" in cells4:
                total_idx.append(i + 1)
        if total_idx:
            ops.append({"ranges": [f"{sid}!A{r}:{end_col}{r}" for r in total_idx],
                        "style": {"font": {"bold": True}, "backColor": "#FFF8DC"}})
        batch_set_style(ops)
        time.sleep(0.3)
    print("  完成")

    # ========== Step 11: 重新嵌入文档 ==========
    print("\n[Step 11] 重新嵌入文档...")
    # 文档结构不变，只需 overwrite 触发 sheet 刷新
    # 实际上 sheet 内嵌是引用关系，格式变化会自动反映
    # 但为确保，我重新 overwrite 一次文档
    xml_path = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_doc.xml")
    if xml_path.exists():
        xml = xml_path.read_text(encoding="utf-8")
    else:
        # 重新构建（复用之前的结构）
        from fix_and_overwrite import build_doc_xml, EXTRACTED as E2
        d2 = json.loads(E2.read_text(encoding="utf-8"))
        xml = build_doc_xml(d2["conclusions"])
    ok_doc, resp_doc = run([LARK_CLI, "docs", "+update", "--api-version", "v2",
                            "--doc", DOC_ID, "--command", "overwrite",
                            "--doc-format", "xml", "--content", "-"],
                           timeout=300)
    # 用 stdin 方式
    r = subprocess.run([LARK_CLI, "docs", "+update", "--api-version", "v2",
                        "--doc", DOC_ID, "--command", "overwrite",
                        "--doc-format", "xml", "--content", "-"],
                       capture_output=True, text=True, encoding="utf-8",
                       shell=False, timeout=300, input=xml)
    if r.returncode == 0:
        print("  文档 overwrite 成功")
    else:
        print(f"  文档 overwrite: {r.stdout[:300]}")

    print("\n[DONE]")
    print(f"文档：https://hcnig43mb8gp.feishu.cn/docx/{DOC_ID}")
    print(f"表格：https://hcnig43mb8gp.feishu.cn/sheets/{SPREADSHEET_TOKEN}")


if __name__ == "__main__":
    main()
