"""把 8 行第 1 列含 < > 的内容做全角替换写入。"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")

SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"

# (sheet_id, row_num, col_index_1based, original_text)
TARGETS = [
    ("2SOCo8", 104, 1, "4）当前实际停课人数： 学员在结束日期前的最后一次正式课上课签到时间，与结束日期相隔30天及以上，且剩余基础课时>0"),
    ("2SOCo8", 105, 1, "5）当前实际停课执行中学员数：学员在结束日期前的最后一次正式课上课签到时间，与结束日期相隔30天及以上，且剩余基础课时>0，学员状态 = 执行中"),
    ("2SOCo8", 108, 1, "8）停课GAP：当停课占比>停课占比目标，停课GAP =（停课占比 - 停课占比目标）* 有效在读人数，当停课占比<=停课占比目标，停课GAP = 0"),
    ("2SOCo8", 114, 1, None),  # 用文件里的真值
    ("2Ti90Y", 102, 1, "首通场景的分母为拨通且命中首通场景新生数（语义分析命中首通场景下的当月新生，且系统累计外呼时长 > 120秒或者企微累计外呼时长 > 0 ）；"),
    ("2TwScg", 106, 1, "7、有效接通率：单次外呼时长>120秒的学员数 / 外呼人数"),
    ("2TwScg", 107, 1, "8、生均有效通次：单次外呼时长>120秒的总外呼次数 / 有效外呼学员数"),
    ("2TwScg", 108, 1, "9、生均有效通时_min：单次外呼时长>120秒的总时长 / 有效外呼学员数"),
]

# 加载 row 114 真值
import json as _json
from pathlib import Path
d = _json.loads(Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json").read_text(encoding="utf-8"))
sm = {s["title"]: s["rows"] for s in d["sheets"]}
row114 = sm["4.4 停课监控"][113]
TARGETS[3] = ("2SOCo8", 114, 1, str(row114[0]))


def col_letter(n):
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def safe_text(s):
    # 半角 < > 替换为全角，避免 cmd.exe 重定向
    return s.replace("<", "＜").replace(">", "＞")


def run(cmd, timeout=60):
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=(sys.platform == "win32"), timeout=timeout)
    if r.returncode != 0:
        return False, r.stderr or r.stdout
    try:
        return True, json.loads(r.stdout)
    except Exception:
        return False, r.stdout


for sid, rn, ci, raw in TARGETS:
    safe = safe_text(raw)
    rng = f"{sid}!{col_letter(ci)}{rn}:{col_letter(ci)}{rn}"
    values = json.dumps([[safe]], ensure_ascii=False)
    ok, resp = run([LARK_CLI, "sheets", "+write",
                    "--spreadsheet-token", SPREADSHEET_TOKEN,
                    "--range", rng, "--values", values])
    print(f"  {sid}!{col_letter(ci)}{rn}: {'OK' if ok else 'FAIL ' + str(resp)[:200]}")
    time.sleep(0.3)
