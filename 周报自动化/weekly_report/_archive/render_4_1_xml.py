"""4.1 统一渲染器 — 生成 Feishu XML 内容
- 输出三段：结论 callout、主表（必要时分多张以满足 2000 cell 限制）、AI 学情表
- 支持 --print 模式（仅打印 XML）和 --push 模式（写入目标文档）
"""
from __future__ import annotations
import sys, io, argparse, subprocess, shutil, json, math
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pandas as pd
from pathlib import Path

ROOT = Path(__file__).parent
sys.path.insert(0, str(ROOT))
from conclusions_4_1 import generate as gen_conclusions

MERGED_MAIN = ROOT.parent / "exports" / "4_1" / "_merged_4_1.xlsx"
MERGED_AI = ROOT.parent / "exports" / "4_1" / "_merged_4_1_ai.xlsx"

LARK_CLI = shutil.which("lark-cli") or "lark-cli"
FEISHU_CELL_LIMIT = 1500  # 留余量，规避 2000 上限

# ── 主表分组（与 export_4_1_excel.py 保持一致，不含"层级"）──
MAIN_GROUPS = [
    ("基础", ["团队/小组", "LP", "在职数"]),
    ("首通语义分析执行", [
        "命中首通新生数",
        "SOP_首通_邀请添加企微执行率",
        "SOP_首通_一家多娃问询执行率",
        "SOP_首通_转介绍执行率",
        "SOP_首通_执行率加和",
    ]),
    ("首通", [
        "首通_新生数", "首通_拨通数",
        "首通_及时跟进率", "首通_生均接通时长",
        "首通_企微绑定率", "首通_秒挂占比",
        "首通_follow率", "首通_生均外呼次数",
    ]),
    ("首课SOP执行", ["SOP_首课_执行率加和"]),
    ("首课", ["首课_学员数", "首课_跟进率", "首课_及时跟进率", "首课_作业完成率"]),
    ("首专", ["首专_学员数", "首专_跟进率", "首专_及时跟进率", "首专_作业完成率"]),
    ("LP架构", ["入职时间", "入职月份", "状态"]),
]
MAIN_LABELS = {
    "团队/小组": "团队/小组", "LP": "LP姓名", "在职数": "到岗",
    "命中首通新生数": "拨通且命中首通场景新生数",
    "SOP_首通_邀请添加企微执行率": "邀请添加企微/WS/Line执行率",
    "SOP_首通_一家多娃问询执行率": "一家多娃问询执行率",
    "SOP_首通_转介绍执行率": "转介绍执行率",
    "SOP_首通_执行率加和": "首通执行率加和",
    "首通_新生数": "首通新生数", "首通_拨通数": "首通拨通数",
    "首通_及时跟进率": "首通及时跟进率",
    "首通_生均接通时长": "首通生均接通时长",
    "首通_企微绑定率": "首通企微绑定率",
    "首通_秒挂占比": "首通秒挂占比",
    "首通_follow率": "首通follow率",
    "首通_生均外呼次数": "首通生均外呼次数",
    "SOP_首课_执行率加和": "首课执行率加和",
    "首课_学员数": "首课学员数", "首课_跟进率": "首课跟进率",
    "首课_及时跟进率": "首课及时跟进率", "首课_作业完成率": "首课作业完成率",
    "首专_学员数": "首专学员数", "首专_跟进率": "首专跟进率",
    "首专_及时跟进率": "首专及时跟进率", "首专_作业完成率": "首专作业完成率",
    "入职时间": "入职时间", "入职月份": "LP工龄(月)", "状态": "人员状态",
}

# ── AI 表分组 ──
AI_GROUPS = [
    ("基础", ["团队/小组", "LP", "状态", "入职月份"]),
    ("总览", ["任务总数", "覆盖学员数"]),
    ("首课AI干预", ["首课_任务总数", "首课_干预中", "首课_干预中占比"]),
    ("首专AI干预", ["首专_任务总数", "首专_干预中", "首专_干预中占比"]),
]
AI_LABELS = {
    "团队/小组": "团队/小组", "LP": "LP姓名",
    "状态": "状态", "入职月份": "工龄(月)",
    "任务总数": "任务总数", "覆盖学员数": "覆盖学员数",
    "首课_任务总数": "首课任务总数", "首课_干预中": "首课干预中",
    "首课_干预中占比": "首课干预中占比",
    "首专_任务总数": "首专任务总数", "首专_干预中": "首专干预中",
    "首专_干预中占比": "首专干预中占比",
}

PCT_COLS = {
    "SOP_首通_邀请添加企微执行率", "SOP_首通_一家多娃问询执行率", "SOP_首通_转介绍执行率",
    "首通_及时跟进率", "首通_企微绑定率", "首通_秒挂占比", "首通_follow率",
    "首课_跟进率", "首课_及时跟进率", "首课_作业完成率",
    "首专_跟进率", "首专_及时跟进率", "首专_作业完成率",
    "首课_干预中占比", "首专_干预中占比",
}
DEC2_COLS = {
    "SOP_首通_执行率加和", "SOP_首课_执行率加和",
    "首通_生均接通时长", "首通_生均外呼次数",
}
INT_COLS = {
    "在职数", "命中首通新生数", "首通_新生数", "首通_拨通数",
    "首课_学员数", "首专_学员数", "入职月份",
    "任务总数", "覆盖学员数",
    "首课_任务总数", "首课_干预中", "首专_任务总数", "首专_干预中",
}


def _fmt(col, val):
    if val is None or (isinstance(val, float) and pd.isna(val)) or val == "":
        return "-"
    if col in PCT_COLS:
        try:
            return f"{float(val)*100:.1f}%"
        except Exception:
            return str(val)
    if col in DEC2_COLS:
        try:
            return f"{float(val):.2f}"
        except Exception:
            return str(val)
    if col in INT_COLS:
        try:
            return f"{int(val)}"
        except Exception:
            return str(val)
    if col == "入职时间":
        try:
            return pd.to_datetime(val).strftime("%Y-%m-%d")
        except Exception:
            return str(val)
    return str(val)


def _xml_escape(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;")
            .replace(">", "&gt;").replace('"', "&quot;"))


def _flat_cols(groups, df_cols):
    cols = []
    for _, gcols in groups:
        for c in gcols:
            if c in df_cols:
                cols.append(c)
    return cols


def _build_table_xml(df: pd.DataFrame, groups, labels):
    cols = _flat_cols(groups, df.columns)
    df = df[cols].copy()
    n_cols = len(cols)
    # 表头：分组行 + 列名行
    rows_xml = []
    # 分组行（合并单元格）
    head1 = "<tr>"
    for gname, gcols in groups:
        present = [c for c in gcols if c in cols]
        if not present:
            continue
        span = len(present)
        if span == 1:
            head1 += f"<td><p><b>{_xml_escape(gname)}</b></p></td>"
        else:
            # Feishu XML 表格不直接支持合并，退化为重复 + 加边界
            head1 += f"<td><p><b>{_xml_escape(gname)}</b></p></td>"
            for _ in range(span - 1):
                head1 += "<td><p></p></td>"
    head1 += "</tr>"
    rows_xml.append(head1)
    head2 = "<tr>"
    for c in cols:
        head2 += f"<td><p><b>{_xml_escape(labels.get(c, c))}</b></p></td>"
    head2 += "</tr>"
    rows_xml.append(head2)
    # 数据
    for _, r in df.iterrows():
        rx = "<tr>"
        for c in cols:
            rx += f"<td><p>{_xml_escape(_fmt(c, r[c]))}</p></td>"
        rx += "</tr>"
        rows_xml.append(rx)

    n_rows = len(rows_xml)
    table = f'<table row-count="{n_rows}" col-count="{n_cols}">' + "".join(rows_xml) + "</table>"
    return table, n_rows * n_cols


def _split_by_team_summary_then_detail(df: pd.DataFrame, n_cols: int, limit: int):
    """主表分页：尽量保证团队汇总在同一张表，只切分个人明细。"""
    summary_df = df[df["层级"] == "汇总"]
    detail_df = df[df["层级"] == "个人"]
    head_rows = 2
    # 第 1 张表 = 汇总 + 部分明细
    max_first_rows = max(0, (limit // n_cols) - head_rows - len(summary_df))
    if max_first_rows >= len(detail_df):
        return [df]
    chunks = []
    chunks.append(pd.concat([summary_df, detail_df.iloc[:max_first_rows]], ignore_index=True))
    # 后续每张：仅明细（也带表头）
    rest = detail_df.iloc[max_first_rows:]
    chunk_size = max(1, (limit // n_cols) - head_rows)
    for s in range(0, len(rest), chunk_size):
        chunks.append(rest.iloc[s:s + chunk_size].reset_index(drop=True))
    return chunks


def render_main():
    df = pd.read_excel(MERGED_MAIN)
    cols = _flat_cols(MAIN_GROUPS, df.columns)
    n_cols = len(cols)
    chunks = _split_by_team_summary_then_detail(df, n_cols, FEISHU_CELL_LIMIT)
    xml_parts = []
    for i, ch in enumerate(chunks):
        suffix = f"（{i + 1}/{len(chunks)}）" if len(chunks) > 1 else ""
        xml_parts.append(f"<h4>合并宽表{suffix}</h4>")
        tbl, _ = _build_table_xml(ch, MAIN_GROUPS, MAIN_LABELS)
        xml_parts.append(tbl)
    return "\n".join(xml_parts)


def render_ai():
    df = pd.read_excel(MERGED_AI)
    cols = _flat_cols(AI_GROUPS, df.columns)
    n_cols = len(cols)
    head_rows = 2
    chunk_size = max(1, (FEISHU_CELL_LIMIT // n_cols) - head_rows)
    parts = []
    summary_df = df[df["层级"] == "汇总"]
    detail_df = df[df["层级"] == "个人"]
    if (len(summary_df) + len(detail_df) + head_rows) * n_cols <= FEISHU_CELL_LIMIT:
        chunks = [df]
    else:
        chunks = [pd.concat([summary_df, detail_df.iloc[:chunk_size - len(summary_df)]], ignore_index=True)]
        rest = detail_df.iloc[chunk_size - len(summary_df):]
        for s in range(0, len(rest), chunk_size):
            chunks.append(rest.iloc[s:s + chunk_size].reset_index(drop=True))
    for i, ch in enumerate(chunks):
        suffix = f"（{i + 1}/{len(chunks)}）" if len(chunks) > 1 else ""
        parts.append(f"<h4>AI学情助手汇总{suffix}</h4>")
        tbl, _ = _build_table_xml(ch, AI_GROUPS, AI_LABELS)
        parts.append(tbl)
    return "\n".join(parts)


def render_callout(text: str) -> str:
    out = ['<callout background-color="rgb(255,245,235)" border-color="rgb(254,212,164)" emoji="❗">']
    for line in text.split("\n"):
        if line.strip():
            out.append(f"<p>{_xml_escape(line)}</p>")
        else:
            out.append("<p></p>")
    out.append("</callout>")
    return "\n".join(out)


def render_all():
    conclusions = gen_conclusions()
    parts = [
        "<h2>4.1 服务指标跟进&amp;语义分析（5.18-5.24）</h2>",
        "<h3>结论</h3>",
        render_callout(conclusions),
        "<h3>主表 - 服务指标 + 语义分析 + LP架构</h3>",
        render_main(),
        "<h3>AI 学情助手</h3>",
        render_ai(),
    ]
    return "\n".join(parts)


def push_to_doc(doc_id: str, xml_file: Path) -> bool:
    # lark-cli 要求 --content @file 是相对路径
    cwd = xml_file.parent
    rel = xml_file.name
    cmd = [LARK_CLI, "docs", "+update", "--api-version", "v2",
           "--doc", doc_id, "--command", "append", "--content", f"@{rel}"]
    res = subprocess.run(cmd, cwd=str(cwd), capture_output=True, text=True, encoding="utf-8",
                         shell=(sys.platform == "win32"))
    if res.returncode != 0:
        print(f"[ERROR] {res.stderr}", file=sys.stderr)
        return False
    try:
        ok = json.loads(res.stdout).get("ok", False)
        print(res.stdout)
        return ok
    except Exception:
        print(res.stdout)
        return False


def main():
    p = argparse.ArgumentParser()
    p.add_argument("--push", help="目标飞书文档 ID")
    p.add_argument("--out", default=str(ROOT / "_4_1_content.xml"), help="保存 XML 到本地文件")
    args = p.parse_args()

    xml = render_all()
    out_path = Path(args.out)
    out_path.write_text(xml, encoding="utf-8")
    print(f"[OK] XML 已保存: {out_path}  ({len(xml)} chars)")
    if args.push:
        ok = push_to_doc(args.push, out_path)
        print(f"[push] ok={ok}")


if __name__ == "__main__":
    main()
