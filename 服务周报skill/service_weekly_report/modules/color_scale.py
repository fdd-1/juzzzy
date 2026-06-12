"""色阶工具 - 用渐变背景色模拟数据条
绿(高,好) → 黄(中) → 红(低,差) 渐变
"""
from __future__ import annotations
import json
import subprocess
import shutil
import sys
import time
from pathlib import Path

LARK_CLI = shutil.which("lark-cli") or "lark-cli"


def col_letter(n: int) -> str:
    """列号转字母 (1->A, 27->AA)。"""
    s = ""
    while n > 0:
        n, r = divmod(n - 1, 26)
        s = chr(ord("A") + r) + s
    return s


def lerp_color(ratio: float) -> str:
    """ratio 0.0(差,红) → 0.5(中,黄) → 1.0(好,绿)
    返回 HEX 色值字符串。
    """
    if ratio >= 0.5:
        # 黄→绿
        t = (ratio - 0.5) * 2
        r = int(255 * (1 - t) + 99 * t)
        g = int(235 * (1 - t) + 190 * t)
        b = int(132 * (1 - t) + 123 * t)
    else:
        # 红→黄
        t = ratio * 2
        r = int(248 * (1 - t) + 255 * t)
        g = int(105 * (1 - t) + 235 * t)
        b = int(107 * (1 - t) + 132 * t)
    return f"#{r:02X}{g:02X}{b:02X}"


def to_float(v) -> float | None:
    """尝试转浮点数,失败返回 None。"""
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace("%", "").replace(",", "")
    try:
        return float(s)
    except ValueError:
        return None


def _run_lark(cmd: list[str], timeout: int = 60) -> tuple[bool, dict | str]:
    r = subprocess.run(
        cmd, capture_output=True, text=True,
        encoding="utf-8", errors="replace",
        shell=(sys.platform == "win32"), timeout=timeout
    )
    if r.returncode != 0:
        return False, r.stderr or r.stdout
    try:
        return True, json.loads(r.stdout)
    except Exception:
        return False, r.stdout


def batch_set_style(spreadsheet_token: str, ops: list[dict]) -> int:
    """批量应用单元格样式。
    ops 每项: {"ranges": [...], "style": {"backColor": "#XXXXXX"}}
    自动按 7500 字节切分。
    返回成功数。
    """
    if not ops:
        return 0
    payload = json.dumps(ops, ensure_ascii=False)
    if len(payload) > 7500:
        mid = len(ops) // 2
        return batch_set_style(spreadsheet_token, ops[:mid]) + batch_set_style(spreadsheet_token, ops[mid:])

    cmd = [
        LARK_CLI, "sheets", "+batch-set-style",
        "--spreadsheet-token", spreadsheet_token,
        "--data", payload,
    ]
    ok, _resp = _run_lark(cmd)
    if not ok:
        # 尝试 --yes
        ok, _resp = _run_lark(cmd + ["--yes"])
    return len(ops) if ok else 0


def apply_color_scale(
    spreadsheet_token: str,
    sheet_id: str,
    col_idx: int,
    values: list,
    start_row: int = 2,
) -> int:
    """对单列应用色阶上色。
    - col_idx: 0-based 列索引
    - values: 数据行的该列值列表(从 start_row 起对应)
    - start_row: 数据起始行号(1-based, 表头之后第一行)
    返回成功上色的单元格数。
    """
    nums = [(i, to_float(v)) for i, v in enumerate(values)]
    valid = [(i, v) for i, v in nums if v is not None and v != 0]
    if len(valid) < 2:
        return 0

    vs = [v for _, v in valid]
    min_v = min(vs)
    max_v = max(vs)
    if max_v == min_v:
        return 0

    col_l = col_letter(col_idx + 1)
    ops = []
    for i, v in valid:
        ratio = (v - min_v) / (max_v - min_v)
        color = lerp_color(ratio)
        row = start_row + i
        rng = f"{sheet_id}!{col_l}{row}:{col_l}{row}"
        ops.append({"ranges": [rng], "style": {"backColor": color}})

    sent = 0
    for i in range(0, len(ops), 30):
        batch = ops[i:i + 30]
        sent += batch_set_style(spreadsheet_token, batch)
        time.sleep(0.3)
    return sent


if __name__ == "__main__":
    print("色阶工具 - 测试用例")
    print(f"  ratio=0.0 → {lerp_color(0.0)} (红)")
    print(f"  ratio=0.5 → {lerp_color(0.5)} (黄)")
    print(f"  ratio=1.0 → {lerp_color(1.0)} (绿)")
    print(f"  col_letter(1) = {col_letter(1)}")
    print(f"  col_letter(27) = {col_letter(27)}")
    print(f"  col_letter(28) = {col_letter(28)}")
