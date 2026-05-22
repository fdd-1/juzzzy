"""4.1 服务指标跟进&语义分析 — 数据整合与结论生成

读取 5 份 BI 导出，整合为周报展示表 + 生成结论文本。
"""
import pandas as pd
import numpy as np
from pathlib import Path

EXPORT_BASE = Path(__file__).parent.parent / "exports" / "4_1"

# 本月目标
TARGETS = {
    "首通及时跟进率": 0.95,
    "首课及时跟进率": 0.85,
    "首专及时跟进率": 0.85,
}


def load_first_call():
    """04_first_call: 首通监控 → 团队+个人级别"""
    f = EXPORT_BASE / "04_first_call" / "益智海外新生首通监控.xlsx"
    raw = pd.read_excel(f, header=None)
    row2 = raw.iloc[2].tolist()

    col_idx = {}
    seen = {}
    for i, v in enumerate(row2):
        if pd.notna(v):
            key = str(v).strip()
            count = seen.get(key, 0)
            if count == 0:
                col_idx[key] = i
            else:
                col_idx[f"{key}_{count}"] = i
            seen[key] = count + 1

    records = []
    current_team = None
    for i in range(3, len(raw)):
        row = raw.iloc[i].tolist()
        if pd.notna(row[1]):
            current_team = str(row[1]).strip()
        lp = row[3] if pd.notna(row[3]) else ""
        if not lp or current_team is None:
            continue
        rec = {
            "小组": current_team,
            "LP": lp,
            "新生数": row[4] if pd.notna(row[4]) else 0,
        }
        for metric in ["跟进率", "及时跟进率", "企微绑定率", "秒挂占比"]:
            idx = col_idx.get(metric)
            rec[metric] = row[idx] if idx and pd.notna(row[idx]) else None
        records.append(rec)

    df = pd.DataFrame(records)
    return df


def load_first_class():
    """02_lp_first_class: 首课指标 (-48h)"""
    f = EXPORT_BASE / "02_lp_first_class" / "海外思维学管服务指标统计表.xlsx"
    raw = pd.read_excel(f, header=None)
    hdr = raw.iloc[1].tolist()

    col_map = {}
    for i, h in enumerate(hdr):
        if pd.notna(h):
            col_map[str(h).strip()] = i

    records = []
    current_team = None
    for i in range(2, len(raw)):
        row = raw.iloc[i].tolist()
        team_val = row[col_map["团队"]]
        lp_val = row[col_map["LP名称"]]
        if pd.notna(team_val):
            current_team = str(team_val).strip()
        if pd.isna(lp_val):
            continue
        rec = {
            "小组": current_team,
            "LP": str(lp_val).strip(),
            "首课学员数": row[col_map.get("首课学员数", 17)],
            "首课跟进率": row[col_map.get("首课跟进率", 22)],
            "首课及时跟进率": row[col_map.get("首课及时跟进率", 24)],
        }
        records.append(rec)
    return pd.DataFrame(records)


def load_first_specialty():
    """03_lp_first_specialty: 首专指标 (-72h)"""
    f = EXPORT_BASE / "03_lp_first_specialty" / "海外思维学管服务指标统计表.xlsx"
    raw = pd.read_excel(f, header=None)
    hdr = raw.iloc[1].tolist()

    col_map = {}
    for i, h in enumerate(hdr):
        if pd.notna(h):
            col_map[str(h).strip()] = i

    records = []
    current_team = None
    for i in range(2, len(raw)):
        row = raw.iloc[i].tolist()
        team_val = row[col_map["团队"]]
        lp_val = row[col_map["LP名称"]]
        if pd.notna(team_val):
            current_team = str(team_val).strip()
        if pd.isna(lp_val):
            continue
        rec = {
            "小组": current_team,
            "LP": str(lp_val).strip(),
            "首专学员数": row[col_map.get("首专学员数", 26)],
            "首专跟进率": row[col_map.get("首专跟进率", 30)],
            "首专及时跟进率": row[col_map.get("首专及时跟进率", 32)],
        }
        records.append(rec)
    return pd.DataFrame(records)


def load_sop():
    """01_sop: 语义分析执行率"""
    f = EXPORT_BASE / "01_sop" / "海外思维服务SOP执行情况.xlsx"
    raw = pd.read_excel(f, header=None)

    records = []
    current_team = None
    for i in range(4, len(raw)):
        row = raw.iloc[i].tolist()
        if pd.notna(row[1]):
            current_team = str(row[1]).strip()
        lp = row[3] if pd.notna(row[3]) else ""
        xinsheng = row[4] if pd.notna(row[4]) else None
        if not lp or xinsheng is None or current_team is None:
            continue
        if "口径" in str(current_team):
            break
        rec = {
            "小组": current_team,
            "LP": lp,
            "命中首通新生数": xinsheng,
            "邀请添加企微执行率": row[10] if pd.notna(row[10]) else None,
            "一家多娃问询执行率": row[11] if pd.notna(row[11]) else None,
            "转介绍执行率": row[12] if pd.notna(row[12]) else None,
            "首通执行率加和": row[13] if pd.notna(row[13]) else None,
        }
        records.append(rec)
    return pd.DataFrame(records)


def load_ai_learning():
    """05_ai_learning: AI学情助手"""
    f = EXPORT_BASE / "05_ai_learning" / "AI学情助手完成情况明细（验收中）.xlsx"
    df = pd.read_excel(f, header=3)
    summary = df.groupby("LP小组").agg(
        总任务数=("任务名称", "count"),
        学员数=("学员id", "nunique"),
    ).reset_index()
    # 干预成功率 per group
    success = df[df["干预状态名称"] == "干预成功"].groupby("LP小组").size().reset_index(name="干预成功数")
    summary = summary.merge(success, on="LP小组", how="left")
    summary["干预成功数"] = summary["干预成功数"].fillna(0).astype(int)
    summary["干预成功率"] = summary["干预成功数"] / summary["总任务数"]
    return summary


def generate_conclusions():
    """生成 4.1 结论文本"""
    fc = load_first_call()
    fclass = load_first_class()
    fspec = load_first_specialty()
    sop = load_sop()
    ai = load_ai_learning()

    lines = []

    # === 首通 ===
    fc_total = fc[fc["LP"] == "总计"]
    fc_overall = fc_total[fc_total["小组"] == "海外教学服务部"].iloc[0]
    gj = fc_overall["跟进率"]
    js = fc_overall["及时跟进率"]
    qw = fc_overall["企微绑定率"]
    mg = fc_overall["秒挂占比"]

    target_st = TARGETS["首通及时跟进率"]
    status = "达到本月目标" if js >= target_st else f"未达本月目标（{target_st:.0%}）"
    lines.append(f"首通：总体跟进率{gj:.2%}；及时跟进率{js:.2%}，{status}")

    # 各组异常
    fc_teams = fc_total[(fc_total["小组"] != "海外教学服务部")]
    low_teams = fc_teams[fc_teams["及时跟进率"].notna() & (fc_teams["及时跟进率"] < js)]
    for _, t in low_teams.iterrows():
        team_lps = fc[(fc["小组"] == t["小组"]) & (fc["LP"] != "总计")]
        low_lps = team_lps.nsmallest(3, "及时跟进率")
        lp_names = "/".join(low_lps["LP"].tolist())
        lines.append(f"  {t['小组']}及时跟进率仅{t['及时跟进率']:.0%}（{lp_names}及时跟进较低）")

    lines.append(f"  企微绑定率整体为{qw:.0%}")
    # 企微低的组
    low_qw = fc_teams[fc_teams["企微绑定率"].notna() & (fc_teams["企微绑定率"] < qw)]
    for _, t in low_qw.iterrows():
        lines.append(f"  {t['小组']}企微绑定率为{t['企微绑定率']:.0%}，主管需关注")
    lines.append(f"  秒挂占比为{mg:.0%}")
    lines.append("")

    # === 首课 ===
    fclass_total = fclass[fclass["LP"] == "总计"]
    fclass_overall = fclass_total[fclass_total["小组"] == "海外团队"].iloc[0]
    sk_gj = fclass_overall["首课跟进率"]
    sk_js = fclass_overall["首课及时跟进率"]
    target_sk = TARGETS["首课及时跟进率"]
    status_sk = "达到本月目标" if sk_js >= target_sk else f"未达本月目标（{target_sk:.0%}）"
    lines.append(f"首课：总体跟进率{sk_gj:.2%}，及时跟进率{sk_js:.2%}，{status_sk}")

    sk_teams = fclass_total[fclass_total["小组"] != "海外团队"]
    low_sk = sk_teams[sk_teams["首课及时跟进率"].notna() & (sk_teams["首课及时跟进率"] < sk_js)]
    for _, t in low_sk.iterrows():
        team_lps = fclass[(fclass["小组"] == t["小组"]) & (fclass["LP"] != "总计")]
        low_lps = team_lps[team_lps["首课及时跟进率"].notna()].nsmallest(3, "首课及时跟进率")
        lp_names = "/".join(low_lps["LP"].tolist())
        lines.append(f"  {t['小组']}及时跟进率仅为{t['首课及时跟进率']:.0%}（{lp_names}及时跟进率较落后）")
    lines.append("")

    # === 首专 ===
    fspec_total = fspec[fspec["LP"] == "总计"]
    fspec_overall = fspec_total[fspec_total["小组"] == "海外团队"].iloc[0]
    sz_gj = fspec_overall["首专跟进率"]
    sz_js = fspec_overall["首专及时跟进率"]
    target_sz = TARGETS["首专及时跟进率"]
    status_sz = "达到本月目标" if sz_js >= target_sz else f"未达本月目标（{target_sz:.0%}）"
    lines.append(f"首专：总体跟进率{sz_gj:.2%}，及时跟进率{sz_js:.2%}，{status_sz}")

    sz_teams = fspec_total[fspec_total["小组"] != "海外团队"]
    low_sz = sz_teams[sz_teams["首专及时跟进率"].notna() & (sz_teams["首专及时跟进率"] < sz_js)]
    for _, t in low_sz.iterrows():
        team_lps = fspec[(fspec["小组"] == t["小组"]) & (fspec["LP"] != "总计")]
        low_lps = team_lps[team_lps["首专及时跟进率"].notna()].nsmallest(3, "首专及时跟进率")
        lp_names = "/".join(low_lps["LP"].tolist())
        lines.append(f"  {t['小组']}跟进&及时跟进率为{t['首专跟进率']:.0%}/{t['首专及时跟进率']:.0%}（{lp_names}未跟进学员较多）")
    lines.append("")

    # === 语义分析 ===
    sop_total = sop[sop["LP"] == "总计"]
    sop_overall = sop_total[sop_total["小组"] == "海外团队"].iloc[0]
    lines.append("语义分析：")
    lines.append(f"  添加企微/WS/Line的执行率为{sop_overall['邀请添加企微执行率']:.1%}")
    lines.append(f"  一家多娃问询&转介绍执行率为{sop_overall['一家多娃问询执行率']:.0%}/{sop_overall['转介绍执行率']:.0%}")
    # 低于平均的组
    sop_teams = sop_total[sop_total["小组"] != "海外团队"]
    avg_qw = sop_overall["邀请添加企微执行率"]
    low_sop = sop_teams[sop_teams["邀请添加企微执行率"].notna() & (sop_teams["邀请添加企微执行率"] < avg_qw)]
    for _, t in low_sop.iterrows():
        lines.append(f"  {t['小组']}邀请添加企微仅为{t['邀请添加企微执行率']:.0%}")
    lines.append("")

    # === AI学情 ===
    lines.append(f"AI学情助手跟进：（{'/'.join(ai['LP小组'].tolist())}）")

    return "\n".join(lines)


def build_summary_table():
    """构建周报展示表（团队级汇总）"""
    fc = load_first_call()
    fclass = load_first_class()
    fspec = load_first_specialty()

    fc_teams = fc[fc["LP"] == "总计"].copy()
    fc_teams = fc_teams.rename(columns={
        "新生数": "首通新生数",
        "跟进率": "首通跟进率",
        "及时跟进率": "首通及时跟进率",
        "企微绑定率": "首通企微绑定率",
    })[["小组", "首通新生数", "首通跟进率", "首通及时跟进率", "首通企微绑定率", "秒挂占比"]]

    sk_teams = fclass[fclass["LP"] == "总计"][["小组", "首课学员数", "首课跟进率", "首课及时跟进率"]]
    sz_teams = fspec[fspec["LP"] == "总计"][["小组", "首专学员数", "首专跟进率", "首专及时跟进率"]]

    # Merge
    merged = fc_teams.merge(sk_teams, on="小组", how="outer")
    merged = merged.merge(sz_teams, on="小组", how="outer")
    return merged


if __name__ == "__main__":
    print("=" * 60)
    print("4.1 服务指标跟进&语义分析 — 结论")
    print("=" * 60)
    print()
    print(generate_conclusions())
    print()
    print("=" * 60)
    print("展示表")
    print("=" * 60)
    tbl = build_summary_table()
    pd.set_option("display.max_columns", 20)
    pd.set_option("display.width", 200)
    print(tbl.to_string(index=False))
