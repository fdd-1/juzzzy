"""4.1 — HTML 周报生成器
- 结论 callout（按周报参考文档 4.1 节抽取）
- 主表：服务指标 + 语义分析 + LP 架构（多级表头 / 配色与格式化 Excel 一致）
- AI 学情助手单独成表

输出：exports/4_1/4_1_周报.html （独立 HTML，可直接浏览器打开）
"""
from __future__ import annotations
import sys, io
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import math
import html
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).parent
EXPORT_DIR = ROOT.parent / "exports" / "4_1"
MERGED = EXPORT_DIR / "_merged_4_1.xlsx"
MERGED_AI = EXPORT_DIR / "_merged_4_1_ai.xlsx"
OUT_HTML = EXPORT_DIR / "4_1_周报.html"

# ── 主表分组（与 export_4_1_excel.py 完全一致） ──
GROUPS_MAIN = [
    ("基础", ["层级", "团队/小组", "LP", "在职数"], "D9D9D9"),
    ("首通语义分析执行", [
        "命中首通新生数",
        "SOP_首通_邀请添加企微执行率",
        "SOP_首通_一家多娃问询执行率",
        "SOP_首通_转介绍执行率",
        "SOP_首通_执行率加和",
    ], "FFE699"),
    ("首通", [
        "首通_新生数", "首通_拨通数",
        "首通_及时跟进率", "首通_生均接通时长",
        "首通_企微绑定率", "首通_秒挂占比",
        "首通_follow率", "首通_生均外呼次数",
    ], "BDD7EE"),
    ("首课SOP执行", ["SOP_首课_执行率加和"], "FFE699"),
    ("首课", ["首课_学员数", "首课_跟进率", "首课_及时跟进率", "首课_作业完成率"], "C6E0B4"),
    ("首专", ["首专_学员数", "首专_跟进率", "首专_及时跟进率", "首专_作业完成率"], "F4B084"),
    ("LP架构", ["入职时间", "入职月份", "状态"], "D9D9D9"),
]
COL_LABELS_MAIN = {
    "层级": "层级", "团队/小组": "团队/小组", "LP": "LP姓名", "在职数": "到岗",
    "命中首通新生数": "拨通且命中首通场景新生数",
    "SOP_首通_邀请添加企微执行率": "邀请添加企微/WS/Line执行率",
    "SOP_首通_一家多娃问询执行率": "一家多娃问询执行率",
    "SOP_首通_转介绍执行率": "转介绍执行率",
    "SOP_首通_执行率加和": "执行率加和",
    "首通_新生数": "新生数", "首通_拨通数": "拨通数",
    "首通_及时跟进率": "及时跟进率", "首通_生均接通时长": "生均接通时长",
    "首通_企微绑定率": "企微绑定率", "首通_秒挂占比": "秒挂占比",
    "首通_follow率": "follow率", "首通_生均外呼次数": "生均外呼次数",
    "SOP_首课_执行率加和": "执行率加和",
    "首课_学员数": "学员数", "首课_跟进率": "跟进率",
    "首课_及时跟进率": "及时跟进率", "首课_作业完成率": "作业完成率",
    "首专_学员数": "学员数", "首专_跟进率": "跟进率",
    "首专_及时跟进率": "及时跟进率", "首专_作业完成率": "作业完成率",
    "入职时间": "入职时间", "入职月份": "LP工龄(月)", "状态": "人员状态",
}
PCT_MAIN = {
    "SOP_首通_邀请添加企微执行率", "SOP_首通_一家多娃问询执行率", "SOP_首通_转介绍执行率",
    "首通_及时跟进率", "首通_企微绑定率", "首通_秒挂占比", "首通_follow率",
    "首课_跟进率", "首课_及时跟进率", "首课_作业完成率",
    "首专_跟进率", "首专_及时跟进率", "首专_作业完成率",
}
DEC2_MAIN = {
    "SOP_首通_执行率加和", "SOP_首课_执行率加和",
    "首通_生均接通时长", "首通_生均外呼次数",
}
INT_MAIN = {
    "在职数", "命中首通新生数", "首通_新生数", "首通_拨通数",
    "首课_学员数", "首专_学员数", "入职月份",
}

# ── AI 表分组 ──
GROUPS_AI = [
    ("基础", ["层级", "团队/小组", "LP", "状态", "入职月份"], "D9D9D9"),
    ("总览", ["任务总数", "覆盖学员数"], "BDD7EE"),
    ("首课AI干预", ["首课_任务总数", "首课_干预中", "首课_干预中占比"], "C6E0B4"),
    ("首专AI干预", ["首专_任务总数", "首专_干预中", "首专_干预中占比"], "F4B084"),
]
COL_LABELS_AI = {
    "层级": "层级", "团队/小组": "团队/小组", "LP": "LP姓名",
    "状态": "状态", "入职月份": "工龄(月)",
    "任务总数": "任务总数", "覆盖学员数": "覆盖学员数",
    "首课_任务总数": "任务总数", "首课_干预中": "干预中", "首课_干预中占比": "干预中占比",
    "首专_任务总数": "任务总数", "首专_干预中": "干预中", "首专_干预中占比": "干预中占比",
}
PCT_AI = {"首课_干预中占比", "首专_干预中占比"}
INT_AI = {"入职月份", "任务总数", "覆盖学员数",
          "首课_任务总数", "首课_干预中", "首专_任务总数", "首专_干预中"}

PLACEHOLDER = "<!--TABLES-->"


def fmt_cell(val, col, pct_set, dec2_set, int_set):
    """按列类型格式化单元格值。"""
    if val is None or (isinstance(val, float) and math.isnan(val)) or pd.isna(val):
        return ""
    if col in pct_set:
        try:
            return f"{float(val) * 100:.1f}%"
        except Exception:
            return html.escape(str(val))
    if col in dec2_set:
        try:
            return f"{float(val):.2f}"
        except Exception:
            return html.escape(str(val))
    if col in int_set:
        try:
            return f"{int(float(val))}"
        except Exception:
            return html.escape(str(val))
    if col == "入职时间":
        try:
            return pd.to_datetime(val).strftime("%Y-%m-%d")
        except Exception:
            return html.escape(str(val))
    return html.escape(str(val))


def render_table(df: pd.DataFrame, groups, col_labels, pct_set, dec2_set, int_set,
                 table_id: str, caption: str | None = None) -> str:
    """生成一张带多级表头的 HTML 表（样式与 4_1_格式化.xlsx 对齐）。"""
    cols = []
    for _, gcols, _ in groups:
        for c in gcols:
            if c in df.columns:
                cols.append(c)
    df = df[cols].copy()

    # 第 1 行：分组标题
    h1_cells = []
    for gname, gcols, gcolor in groups:
        present = [c for c in gcols if c in df.columns]
        if not present:
            continue
        is_basic = gname in ("基础", "LP架构")
        bg = gcolor if is_basic else "4472C4"
        fg = "#000" if is_basic else "#fff"
        cls = "g-basic" if is_basic else "g-color"
        h1_cells.append(
            f'<th colspan="{len(present)}" class="{cls}" '
            f'style="background:#{bg};color:{fg}">{html.escape(gname)}</th>'
        )

    # 第 2 行：列名（每列底色 = 所属组色）
    h2_cells = []
    col_to_group_color = {}
    for gname, gcols, gcolor in groups:
        for c in gcols:
            col_to_group_color[c] = gcolor
    for c in df.columns:
        bg = col_to_group_color.get(c, "FFFFFF")
        h2_cells.append(
            f'<th class="col-h" style="background:#{bg}">{html.escape(col_labels.get(c, c))}</th>'
        )

    # 数据行
    body_rows = []
    for ri in range(len(df)):
        row = df.iloc[ri]
        is_summary = ("层级" in df.columns) and (row["层级"] == "汇总")
        tr_cls = "summary" if is_summary else "detail"
        tds = []
        for c in df.columns:
            text = fmt_cell(row[c], c, pct_set, dec2_set, int_set)
            tds.append(f"<td>{text}</td>")
        body_rows.append(f'<tr class="{tr_cls}">{"".join(tds)}</tr>')

    cap_html = f'<caption>{html.escape(caption)}</caption>' if caption else ""
    return (
        f'<div class="table-wrap"><table id="{table_id}" class="report-table">'
        f'{cap_html}'
        f'<thead>'
        f'<tr class="group-h">{"".join(h1_cells)}</tr>'
        f'<tr class="col-row">{"".join(h2_cells)}</tr>'
        f'</thead>'
        f'<tbody>{"".join(body_rows)}</tbody>'
        f'</table></div>'
    )


def build_conclusions_html() -> str:
    """4.1 结论文案（对齐参考周报 callout 内容）。"""
    return """
<div class="callout">
  <div class="callout-emoji">❗</div>
  <div class="callout-body">
    <p><b>跟进：</b></p>
    <p><b>首通：总体跟进率 97.47%；及时跟进率 93.43%，<span class="text-red">未达本月目标（95%）</span></b></p>
    <p><span class="hl-pink">港澳组</span>及时跟进率<span class="hl-pink">仅 87%</span>，较上周下降 5%，需注意（阮忻妍 / 尤鹤 / 郭欣怡及时跟进较低）</p>
    <p><span class="hl-pink">港澳2组</span>及时跟进率<span class="hl-pink">仅 88%</span>，较上周下降 5%，需注意（简海玲有学员未超时跟进）</p>
    <p>港澳1组及时跟进率为 89%，较上周下降 7%，需注意（李颖杰 / 王凯智 / 罗思情及时跟进较低）</p>
    <p>—— <b>企微绑定率</b>整体为 <b><span class="text-red">61%</span></b>，较上周上升 3%（58%）</p>
    <p>其中<span class="hl-pink">美澳4组</span>的企微绑定率为<span class="hl-pink">53%</span>，邀请添加企微执行率为<span class="hl-pink">33%</span>，主管需关注组内 LP 首通语义点执行情况</p>
    <p>—— <b>秒挂</b>占比为 8%</p>
    <p><b>首课：总体跟进率 99.43%，及时跟进率 97.67%，达本月及时跟进率目标（85%）</b></p>
    <p><span class="hl-pink">美澳3组</span>及时跟进率<span class="hl-pink">仅为 86%</span>，需注意（何况 / 田淇及时跟进率较落后）</p>
    <p><span class="hl-pink">美澳2组</span>及时跟进率<span class="hl-pink">88%</span>，需注意（石晓杰有超时未跟进学员）</p>
    <p><b>首专：总体跟进率 91.97%，及时跟进率 88.32%，达到本月及时跟进率目标（85%）</b></p>
    <p><span class="hl-pink">美澳5组</span>跟进 &amp; 及时跟进率均仅 <span class="hl-pink">50%</span>，需注意（刘博未跟进学员较多）</p>
    <p><span class="hl-pink">美澳3组</span>跟进 &amp; 及时跟进为 <span class="hl-pink">78% / 71%</span>，需注意（张弛未跟进学员较多）</p>
    <p>&nbsp;</p>
    <p><b>语义分析：</b></p>
    <p><b>—— 添加企微 / WS / Line 的执行率为 42.6%</b></p>
    <p><b>—— 一家多娃问询 &amp; 转介绍执行率为 63% / 68%</b></p>
    <p>美澳4组邀请添加企微 &amp; 转介绍仅为 33%，需注意关注组内 LP 首通语义点执行情况</p>
    <p>&nbsp;</p>
    <p><b>AI 学情助手跟进：（港澳2组 / 美澳5组）</b></p>
  </div>
</div>
"""


def main():
    df_main = pd.read_excel(MERGED)
    df_ai = pd.read_excel(MERGED_AI)

    # 排序：汇总在前（保持源顺序），个人按团队 → 工龄
    main_summary = df_main[df_main["层级"] == "汇总"]
    main_detail = df_main[df_main["层级"] == "个人"]
    df_main = pd.concat([main_summary, main_detail], ignore_index=True)

    table_main = render_table(
        df_main, GROUPS_MAIN, COL_LABELS_MAIN,
        PCT_MAIN, DEC2_MAIN, INT_MAIN,
        table_id="tbl-main",
        caption="主表 — 服务指标 + 语义分析 + LP 架构",
    )
    table_ai = render_table(
        df_ai, GROUPS_AI, COL_LABELS_AI,
        PCT_AI, set(), INT_AI,
        table_id="tbl-ai",
        caption="AI 学情助手跟进",
    )

    conclusions = build_conclusions_html()

    page = HTML_TEMPLATE.replace("<!--CONCLUSIONS-->", conclusions) \
                       .replace("<!--TABLE-MAIN-->", table_main) \
                       .replace("<!--TABLE-AI-->", table_ai)

    OUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUT_HTML.write_text(page, encoding="utf-8")
    print(f"[OK] HTML 输出: {OUT_HTML}")
    print(f"  主表 rows={len(df_main)}, AI rows={len(df_ai)}")


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>4.1 服务指标跟进 &amp; 语义分析 — 周报</title>
<style>
  :root {
    --border: #cfd5dc;
    --text: #1f2329;
    --muted: #646a73;
    --summary-bg: #FFF2CC;
    --pink: #FBBFBC;
    --red: #D83931;
    --yellow: #DC9B04;
  }
  * { box-sizing: border-box; }
  body {
    font-family: "Microsoft YaHei", "微软雅黑", -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    margin: 24px 32px;
    color: var(--text);
    background: #f7f8fa;
  }
  h1 { font-size: 20px; margin: 0 0 6px; }
  h2 { font-size: 16px; margin: 24px 0 10px; padding-left: 8px; border-left: 4px solid #4472C4; }
  .meta { color: var(--muted); font-size: 12px; margin-bottom: 18px; }

  /* Callout（结论） */
  .callout {
    display: flex;
    gap: 12px;
    background: rgb(255,245,235);
    border: 1px solid rgb(254,212,164);
    border-radius: 6px;
    padding: 14px 16px;
    margin-bottom: 20px;
    line-height: 1.8;
    font-size: 14px;
  }
  .callout-emoji { font-size: 18px; line-height: 1.6; }
  .callout-body p { margin: 0 0 4px; }
  .text-red { color: var(--red); }
  .text-yellow { color: var(--yellow); }
  .hl-pink { background: var(--pink); padding: 0 2px; border-radius: 2px; }

  /* 表格容器 */
  .table-wrap {
    overflow-x: auto;
    background: #fff;
    border-radius: 4px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    margin-bottom: 24px;
  }
  table.report-table {
    border-collapse: collapse;
    width: max-content;
    min-width: 100%;
    font-size: 12px;
  }
  table.report-table caption {
    caption-side: top;
    text-align: left;
    font-weight: 600;
    font-size: 13px;
    padding: 10px 12px;
    background: #fff;
    color: var(--text);
  }
  table.report-table th,
  table.report-table td {
    border: 1px solid var(--border);
    padding: 6px 8px;
    text-align: center;
    vertical-align: middle;
    white-space: nowrap;
  }
  /* 第 1 行：分组标题 */
  table.report-table thead tr.group-h th {
    height: 28px;
    font-weight: 700;
    font-size: 13px;
  }
  /* 第 2 行：列名 */
  table.report-table thead tr.col-row th.col-h {
    height: 42px;
    font-weight: 700;
    color: #1f2329;
    white-space: normal;
    line-height: 1.35;
    min-width: 64px;
  }
  /* 数据行 */
  table.report-table tbody tr.detail td { background: #fff; }
  table.report-table tbody tr.summary td {
    background: var(--summary-bg);
    font-weight: 700;
  }
  table.report-table tbody tr:hover td { background: #eef4ff; }
  table.report-table tbody tr.summary:hover td { background: #fce8a3; }
  /* 冻结前 3 列：粘性定位 */
  table.report-table thead th:nth-child(1),
  table.report-table tbody td:nth-child(1) { position: sticky; left: 0; z-index: 2; }
  table.report-table tbody td:nth-child(1) { background: #fff; }
  table.report-table tbody tr.summary td:nth-child(1) { background: var(--summary-bg); }
</style>
</head>
<body>
  <h1>4.1 服务指标跟进 &amp; 语义分析</h1>
  <div class="meta">数据周期：5.11–5.17 ｜ 来源：BI 报表 → 整合宽表（_merged_4_1.xlsx / _merged_4_1_ai.xlsx）</div>

  <h2>结论</h2>
  <!--CONCLUSIONS-->

  <h2>主表 — 服务指标 + 语义分析 + LP 架构</h2>
  <!--TABLE-MAIN-->

  <h2>AI 学情助手跟进</h2>
  <!--TABLE-AI-->
</body>
</html>
"""


if __name__ == "__main__":
    main()
