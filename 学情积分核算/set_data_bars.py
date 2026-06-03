"""通过 lark-cli api + stdin 设置数据条条件格式。

飞书条件格式 API：POST /open-apis/sheets/v2/spreadsheets/:token/condition_formats
需要用 PowerShell 避免 Git Bash 路径转换问题。
"""
from __future__ import annotations
import sys, io, json, subprocess, shutil, time, re
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
from pathlib import Path

EXTRACTED = Path(r"C:\Users\fengjianyi\Desktop\学情积分核算\_extracted.json")
SPREADSHEET_TOKEN = "M4NHsX8DHhcbTptLdFmcim4HnJh"
LARK_CLI = shutil.which("lark-cli") or "lark-cli"

SHEET_ID_MAP = {
    "4.1 服务指标": "2RK2qc", "4.1 AI学情": "2S16wg",
    "4.2 组班意向": "2SgAAo", "4.3 群发消息": "2SwpXO",
    "4.4 停课监控": "2SOCo8", "4.5 服务池跟进": "2T3bYQ",
    "4.5 服务池SOP": "2Ti90Y", "4.6 系统外呼监控": "2TwScg",
    "4.6 企微回复比": "2TLJU4",
}
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


def find_pct_cols(header_row):
    cols = []
    for i, cell in enumerate(header_row):
        v = str(cell).strip()
        if "占比" in v or "率" in v:
            cols.append(i)
    return cols


def call_api(method, path, body=None):
    """通过 subprocess 调 lark-cli api，stdin 传 body 避免 shell 转义问题"""
    cmd = [LARK_CLI, "api", method, path, "--as", "user"]
    if body:
        cmd += ["--data", "-"]
    r = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8",
                       shell=False, timeout=60,
                       input=json.dumps(body, ensure_ascii=False) if body else None)
    try:
        resp = json.loads(r.stdout)
        return resp.get("ok", False), resp
    except Exception:
        return False, r.stderr or r.stdout


def create_condition_formats(sheet_id, formats_list):
    """创建条件格式"""
    path = f"/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/condition_formats"
    body = {
        "sheet_id": sheet_id,
        "condition_formats": formats_list,
    }
    ok, resp = call_api("POST", path, body)
    return ok, resp


def main():
    extracted = json.loads(EXTRACTED.read_text(encoding="utf-8"))
    excel_data_map = {s["title"]: s["rows"] for s in extracted["sheets"]}

    # 先测试 API 是否可达
    print("测试条件格式 API...")
    test_path = f"/open-apis/sheets/v2/spreadsheets/{SPREADSHEET_TOKEN}/condition_formats"
    ok, resp = call_api("POST", test_path, {
        "sheet_id": "2RK2qc",
        "condition_formats": [{
            "ranges": ["2RK2qc!F3:F10"],
            "rule_type": "color_scale",
            "attrs": [
                {"type": "min", "color": "#FFFFFF"},
                {"type": "max", "color": "#5B8FF9"},
            ],
        }],
    })
    print(f"  测试结果: ok={ok}")
    if not ok:
        err = resp if isinstance(resp, str) else json.dumps(resp, ensure_ascii=False)[:500]
        print(f"  错误: {err}")
        # 如果 API 确实不可用，尝试另一种格式
        # 飞书数据条的 rule_type 可能是 "dataBar" 而不是 "data_bar"
        ok2, resp2 = call_api("POST", test_path, {
            "sheet_id": "2RK2qc",
            "condition_formats": [{
                "ranges": ["2RK2qc!F3:F10"],
                "rule_type": "dataBar",
                "attrs": {
                    "min_point": {"type": "min", "value": ""},
                    "max_point": {"type": "max", "value": ""},
                    "bar_color": "#5B8FF9",
                },
            }],
        })
        print(f"  尝试 dataBar: ok={ok2}")
        if not ok2:
            err2 = resp2 if isinstance(resp2, str) else json.dumps(resp2, ensure_ascii=False)[:500]
            print(f"  错误2: {err2}")
            if "404" in str(err2):
                print("\n  [结论] 条件格式 API 不可达（404），可能是 lark-cli 版本过旧不支持此路由。")
                print("  建议：升级 lark-cli 到最新版（lark-cli update），或手动在飞书表格中设置数据条。")
                return
    else:
        print("  API 可达！开始批量设置...")

    # 如果测试成功，批量为所有 sheet 的占比/率列设置色阶格式
    for title, sid in SHEET_ID_MAP.items():
        rows = excel_data_map[title]
        cut = find_caliber_start(rows)
        rows = [r for r in rows[:cut] if any(str(c).strip() for c in r)]
        if not rows:
            continue
        n_rows = len(rows)
        n_header = HEADER_ROWS[title]
        detail_row = rows[n_header - 1]
        pct_cols = find_pct_cols(detail_row)
        if not pct_cols:
            print(f"  [{title}] 无占比/率列，跳过")
            continue

        # 构造条件格式列表
        formats = []
        for col_idx in pct_cols:
            col_l = col_letter(col_idx + 1)
            rng = f"{sid}!{col_l}{n_header + 1}:{col_l}{n_rows}"
            formats.append({
                "ranges": [rng],
                "rule_type": "color_scale",
                "attrs": [
                    {"type": "min", "color": "#FFFFFF"},
                    {"type": "max", "color": "#5B8FF9"},
                ],
            })

        # 分批提交（每次最多 10 个）
        for i in range(0, len(formats), 10):
            batch = formats[i:i + 10]
            ok, resp = create_condition_formats(sid, batch)
            status = "OK" if ok else "FAIL"
            print(f"  [{title}] batch {i//10+1}: {status} ({len(batch)} cols)")
            if not ok:
                err = resp if isinstance(resp, str) else json.dumps(resp, ensure_ascii=False)[:300]
                print(f"    {err}")
            time.sleep(0.5)

    print("\n[DONE]")


if __name__ == "__main__":
    main()
