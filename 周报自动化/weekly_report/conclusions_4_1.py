"""4.1 结论生成 v2 — 基于 _merged_4_1.xlsx"""
from __future__ import annotations
import pandas as pd
from pathlib import Path

MERGED = Path(__file__).parent.parent / "exports" / "4_1" / "_merged_4_1.xlsx"

TARGETS = {
    "首通及时跟进率": 0.95,
    "首课及时跟进率": 0.85,
    "首专及时跟进率": 0.85,
}


def _fmt_pct(v):
    if pd.isna(v):
        return "-"
    return f"{v:.0%}"


def _team_low_lps(detail: pd.DataFrame, team: str, metric: str, n: int = 3, threshold: float = None):
    sub = detail[(detail["团队/小组"] == team) & (detail[metric].notna())]
    if threshold is not None:
        sub = sub[sub[metric] < threshold]
    sub = sub.sort_values(metric, ascending=True).head(n)
    return sub["LP"].tolist()


EXCLUDE_TEAMS = {"台湾组"}


def generate():
    df = pd.read_excel(MERGED)
    # 撰写结论时不看台湾组
    df = df[~df["团队/小组"].isin(EXCLUDE_TEAMS)]
    summary = df[df["层级"] == "汇总"].set_index("团队/小组")
    detail = df[df["层级"] == "个人"]

    overall = summary.loc["海外团队（合）"]

    lines = []

    # === 跟进 ===
    lines.append("跟进：")

    # 首通
    js = overall["首通_及时跟进率"]
    gj = overall["首通_跟进率"]
    qw = overall["首通_企微绑定率"]
    mg = overall["首通_秒挂占比"]
    target = TARGETS["首通及时跟进率"]
    status = "达本月目标" if js >= target else f"未达本月目标（{target:.0%}）"
    lines.append(f"  首通：总体跟进率{gj:.2%}；及时跟进率{js:.2%}，{status}")

    # 各组（按及时跟进率从低到高，列出低于整体值的组）
    teams = summary.drop("海外团队（合）", errors="ignore")
    low_fc = teams[(teams["首通_及时跟进率"].notna()) & (teams["首通_及时跟进率"] < js)]
    low_fc = low_fc.sort_values("首通_及时跟进率")
    for team, r in low_fc.iterrows():
        lps = _team_low_lps(detail, team, "首通_及时跟进率", n=3, threshold=r["首通_及时跟进率"])
        suffix = f"（{'/'.join(lps)}及时跟进较低）" if lps else ""
        lines.append(f"    {team}及时跟进率仅{r['首通_及时跟进率']:.0%}{suffix}")

    lines.append(f"  企微绑定率整体为{qw:.0%}")
    low_qw = teams[(teams["首通_企微绑定率"].notna()) & (teams["首通_企微绑定率"] < qw)]
    low_qw = low_qw.sort_values("首通_企微绑定率")
    for team, r in low_qw.iterrows():
        lines.append(f"    {team}企微绑定率为{r['首通_企微绑定率']:.0%}，主管需关注")
    lines.append(f"  秒挂占比为{mg:.0%}")
    lines.append("")

    # 首课
    sk_gj = overall["首课_跟进率"]
    sk_js = overall["首课_及时跟进率"]
    target = TARGETS["首课及时跟进率"]
    status = "达本月目标" if sk_js >= target else f"未达本月目标（{target:.0%}）"
    lines.append(f"  首课：总体跟进率{sk_gj:.2%}，及时跟进率{sk_js:.2%}，{status}")
    low_sk = teams[(teams["首课_及时跟进率"].notna()) & (teams["首课_及时跟进率"] < sk_js)]
    low_sk = low_sk.sort_values("首课_及时跟进率")
    for team, r in low_sk.iterrows():
        lps = _team_low_lps(detail, team, "首课_及时跟进率", n=3, threshold=r["首课_及时跟进率"])
        suffix = f"（{'/'.join(lps)}及时跟进率较落后）" if lps else ""
        lines.append(f"    {team}及时跟进率仅为{r['首课_及时跟进率']:.0%}{suffix}")
    lines.append("")

    # 首专
    sz_gj = overall["首专_跟进率"]
    sz_js = overall["首专_及时跟进率"]
    target = TARGETS["首专及时跟进率"]
    status = "达本月目标" if sz_js >= target else f"未达本月目标（{target:.0%}）"
    lines.append(f"  首专：总体跟进率{sz_gj:.2%}，及时跟进率{sz_js:.2%}，{status}")
    low_sz = teams[(teams["首专_及时跟进率"].notna()) & (teams["首专_及时跟进率"] < sz_js)]
    low_sz = low_sz.sort_values("首专_及时跟进率")
    for team, r in low_sz.iterrows():
        lps = _team_low_lps(detail, team, "首专_及时跟进率", n=3, threshold=r["首专_及时跟进率"])
        suffix = f"（{'/'.join(lps)}未跟进学员较多）" if lps else ""
        lines.append(f"    {team}跟进&及时跟进率为{r['首专_跟进率']:.0%}/{r['首专_及时跟进率']:.0%}{suffix}")
    lines.append("")

    # === 语义分析 ===
    lines.append("语义分析：")
    qw_sop = overall["SOP_首通_邀请添加企微执行率"]
    yj_sop = overall["SOP_首通_一家多娃问询执行率"]
    zj_sop = overall["SOP_首通_转介绍执行率"]
    lines.append(f"  添加企微/WS/Line的执行率为{qw_sop:.1%}")
    lines.append(f"  一家多娃问询&转介绍执行率为{yj_sop:.0%}/{zj_sop:.0%}")
    low_sop = teams[(teams["SOP_首通_邀请添加企微执行率"].notna()) & (teams["SOP_首通_邀请添加企微执行率"] < qw_sop)]
    low_sop = low_sop.sort_values("SOP_首通_邀请添加企微执行率")
    for team, r in low_sop.iterrows():
        lines.append(f"    {team}邀请添加企微&转介绍仅为{r['SOP_首通_邀请添加企微执行率']:.0%}/{r['SOP_首通_转介绍执行率']:.0%}")
    lines.append("")

    # === AI 学情助手 ===
    lines.append("AI学情助手跟进：")
    ai_teams = teams[teams["AI_首课干预中占比"].notna()].copy()
    if not ai_teams.empty:
        # 按首课_干预中占比从高到低排
        ai_teams = ai_teams.sort_values("AI_首课干预中占比", ascending=False)
        for team, r in ai_teams.iterrows():
            sk = r.get("AI_首课干预中占比")
            sz = r.get("AI_首专干预中占比")
            sk_s = f"{sk:.0%}" if pd.notna(sk) else "-"
            sz_s = f"{sz:.0%}" if pd.notna(sz) else "-"
            lines.append(f"  {team}：首课_干预中占比{sk_s}，首专_干预中占比{sz_s}")

    return "\n".join(lines)


if __name__ == "__main__":
    print(generate())
