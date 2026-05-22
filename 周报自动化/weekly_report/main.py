"""
周报自动化 - 服务模块
从飞书电子表格读取数据，分析指标，生成结论，写入新飞书文档
"""
import yaml
import json
import subprocess
import sys
from pathlib import Path
from datetime import date, timedelta
from typing import Optional

CONFIG_PATH = Path(__file__).parent / "report_config.yaml"
LARK_CLI = "lark-cli"
if sys.platform == "win32":
    import shutil
    _found = shutil.which("lark-cli")
    if _found:
        LARK_CLI = _found


def load_config():
    return yaml.safe_load(CONFIG_PATH.read_text(encoding="utf-8"))


def get_week_range():
    today = date.today()
    monday = today - timedelta(days=today.weekday())
    sunday = monday + timedelta(days=6)
    last_monday = monday - timedelta(days=7)
    last_sunday = monday - timedelta(days=1)
    return {
        "current": (monday.isoformat(), sunday.isoformat()),
        "previous": (last_monday.isoformat(), last_sunday.isoformat()),
        "display": f"{monday.month}.{monday.day}-{sunday.month}.{sunday.day}",
    }


def read_sheet(spreadsheet_token: str, sheet_id: str, range_str: str = "A1:AZ200"):
    """通过 lark-cli 读取飞书电子表格数据"""
    cmd = [
        LARK_CLI, "sheets", "+read",
        "--spreadsheet-token", spreadsheet_token,
        "--sheet-id", sheet_id,
        "--range", range_str,
    ]
    result = subprocess.run(cmd, capture_output=True, text=True, encoding="utf-8", shell=(sys.platform == "win32"))
    if result.returncode != 0:
        print(f"[ERROR] 读取 sheet {sheet_id} 失败: {result.stderr}", file=sys.stderr)
        return None
    data = json.loads(result.stdout)
    if not data.get("ok"):
        print(f"[ERROR] API 返回错误: {data}", file=sys.stderr)
        return None
    return data["data"]["valueRange"]["values"]


def parse_table(raw_values: list[list], header_rows: int = 2) -> list[dict]:
    """将飞书表格的多行表头 + 数据行解析为字典列表"""
    if not raw_values or len(raw_values) < header_rows + 1:
        return []
    headers = []
    for col_idx in range(len(raw_values[0])):
        parts = []
        for row_idx in range(header_rows):
            if row_idx < len(raw_values) and col_idx < len(raw_values[row_idx]):
                val = raw_values[row_idx][col_idx]
                if val is not None:
                    parts.append(str(val).strip())
        headers.append("_".join(parts) if parts else f"col_{col_idx}")
    records = []
    for row in raw_values[header_rows:]:
        if not row or all(v is None for v in row):
            continue
        record = {}
        for i, h in enumerate(headers):
            record[h] = row[i] if i < len(row) else None
        records.append(record)
    return records


def fmt_pct(value, digits=2) -> str:
    """格式化百分比"""
    if value is None:
        return "N/A"
    if isinstance(value, (int, float)):
        if value <= 1:
            return f"{value * 100:.{digits}f}%"
        return f"{value:.{digits}f}%"
    return str(value)


def fmt_change(current, previous) -> str:
    """格式化环比变化"""
    if current is None or previous is None:
        return "数据缺失"
    if previous == 0:
        return "上周为0"
    change = (current - previous) / previous
    direction = "上升" if change > 0 else "下降"
    return f"较上周{direction}{abs(change)*100:.0f}%（{fmt_pct(previous)}）"
