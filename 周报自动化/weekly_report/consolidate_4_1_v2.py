"""4.1 服务指标跟进&语义分析 — 数据整合 v2

整合 6 份数据 + AI 学情：
  01_sop                 海外思维服务SOP执行情况（首通/首课语义分析执行率）
  02_lp_first_class      海外思维学管服务指标统计表（-48h，首课）
  03_lp_first_specialty  海外思维学管服务指标统计表（-72h，首专）
  04_first_call          益智海外新生首通监控
  06_lp_arch             海外思维LP架构表
  07_ai_summary          AI 学情助手（仅供结论引用，不在格式化 Excel 中渲染）

输出：宽表 _merged_4_1.xlsx
  - 列序对齐图片版式（基础 / 首通语义分析 / 首通 / 首课SOP / 首课 / 首专 / LP架构）
  - 末尾追加 AI / 结论辅助列（首通_跟进率/接通率），仅供 conclusions 使用
"""
from __future__ import annotations
import pandas as pd
import numpy as np
from pathlib import Path

EXPORT_BASE = Path(__file__).parent.parent / "exports" / "4_1"
OUT_PATH = EXPORT_BASE / "_merged_4_1.xlsx"

TEAM_ORDER = ["港澳1组", "港澳2组", "港澳组",
              "美澳1组", "美澳2组", "美澳3组", "美澳4组", "美澳5组", "台湾组"]


# ── 各源加载 ─────────────────────────────────────────────────────

def load_sop():
    """SOP — header 4 行，data 从 row4 开始
    col1=小组, col2=负责人(主管), col3=LP
    col4=拨通且命中首通场景新生数
    col10=邀请添加企微/WS/Line执行率, col11=一家多娃问询执行率, col12=转介绍执行率, col13=首通执行率加和
    col23=首课执行率加和
    col39=服务池语义点执行率加和
    """
    f = EXPORT_BASE / "01_sop" / "海外思维服务SOP执行情况.xlsx"
    raw = pd.read_excel(f, header=None)
    rows = []
    cur, cur_mgr = None, None
    for i in range(4, len(raw)):
        row = raw.iloc[i].tolist()
        if pd.notna(row[1]):
            cur = str(row[1]).strip()
            cur_mgr = str(row[2]).strip() if pd.notna(row[2]) else None
        if pd.isna(row[3]):
            continue
        team_str = cur or ""
        if "口径" in team_str or team_str.startswith(("1、", "2、", "3、", "4、", "5、")):
            break
        rows.append({
            "小组": cur, "主管": cur_mgr, "LP": str(row[3]).strip(),
            "命中首通新生数": row[4],
            "首通_邀请添加企微执行率": row[10],
            "首通_一家多娃问询执行率": row[11],
            "首通_转介绍执行率": row[12],
            "首通_执行率加和": row[13],
            "首课_执行率加和": row[23],
            "服务池_语义点执行率加和": row[39],
        })
    return pd.DataFrame(rows)


def load_first_call():
    f = EXPORT_BASE / "04_first_call" / "益智海外新生首通监控.xlsx"
    raw = pd.read_excel(f, header=None)
    rows = []
    cur_team, cur_mgr = None, None
    for i in range(3, len(raw)):
        row = raw.iloc[i].tolist()
        if pd.notna(row[1]):
            cur_team = str(row[1]).strip()
        if pd.notna(row[2]):
            cur_mgr = str(row[2]).strip() if str(row[2]).strip() != "总计" else cur_mgr
        if pd.isna(row[3]):
            continue
        rows.append({
            "小组": cur_team, "主管": cur_mgr, "LP": str(row[3]).strip(),
            "首通_新生数": row[4],
            "首通_拨通数": row[13],
            "首通_一家多娃新生数": row[9],
            "首通_跟进率": row[18],
            "首通_接通率": row[24],
            "首通_及时跟进率": row[20],
            "首通_生均接通时长": row[25],
            "首通_企微绑定率": row[26],
            "首通_48小时企微绑定率": row[27],
            "首通_秒挂占比": row[28],
            "首通_follow率": row[17],
            "首通_生均外呼次数": row[30],
        })
    return pd.DataFrame(rows)


def _load_lp_metric(file_path: str, col_prefix: str):
    raw = pd.read_excel(file_path, header=None)
    hdr = raw.iloc[1].tolist()
    cm = {str(h).strip(): i for i, h in enumerate(hdr) if pd.notna(h)}
    rows = []
    cur = None
    for i in range(2, len(raw)):
        row = raw.iloc[i].tolist()
        if pd.notna(row[cm["团队"]]):
            cur = str(row[cm["团队"]]).strip()
        lp = row[cm["LP名称"]]
        if pd.isna(lp):
            continue
        rec = {"小组_raw": cur, "LP": str(lp).strip()}
        rec[f"{col_prefix}_学员数"] = row[cm[f"{col_prefix}学员数"]]
        rec[f"{col_prefix}_跟进率"] = row[cm[f"{col_prefix}跟进率"]]
        rec[f"{col_prefix}_及时跟进率"] = row[cm[f"{col_prefix}及时跟进率"]]
        rec[f"{col_prefix}_作业完成率"] = row[cm[f"{col_prefix}作业完成率"]]
        rows.append(rec)
    return pd.DataFrame(rows)


def load_first_class():
    df = _load_lp_metric(EXPORT_BASE / "02_lp_first_class" / "海外思维学管服务指标统计表.xlsx", "首课")
    return df.rename(columns={"小组_raw": "小组_首课"})


def load_first_specialty():
    df = _load_lp_metric(EXPORT_BASE / "03_lp_first_specialty" / "海外思维学管服务指标统计表.xlsx", "首专")
    return df.rename(columns={"小组_raw": "小组_首专"})


def load_lp_arch():
    f = EXPORT_BASE / "06_lp_arch" / "海外思维LP架构表.xlsx"
    df = pd.read_excel(f, header=3)
    df = df.rename(columns={"姓名": "LP", "入职时间": "入职时间", "入职时长（月份）": "入职月份",
                            "入职时长分组": "入职时长", "是否在职": "状态"})
    return df[["小组", "主管", "LP", "入职时间", "入职月份", "入职时长", "状态", "职位1"]]


def load_ai_summary():
    """AI 学情：返回 LP 级 干预中占比 + 团队聚合"""
    f = EXPORT_BASE / "07_ai_summary" / "AI学情助手完成情况汇总（验收中）.xlsx"
    df = pd.read_excel(f, header=3).rename(columns={"班主任姓名": "LP", "LP小组": "小组"})
    df["AI_首课干预中占比"] = np.where(df["首课_任务总数"] > 0, df["首课_干预中"] / df["首课_任务总数"], np.nan)
    df["AI_首专干预中占比"] = np.where(df["首专_任务总数"] > 0, df["首专_干预中"] / df["首专_任务总数"], np.nan)
    return df


# ── 合并 ─────────────────────────────────────────────────────────

def build_merged():
    sop = load_sop()
    fc = load_first_call()
    fclass = load_first_class()
    fspec = load_first_specialty()
    arch = load_lp_arch()
    ai = load_ai_summary()

    team_sop = sop[sop["LP"] == "总计"].copy()
    team_fc = fc[fc["LP"] == "总计"].copy()
    team_fclass = fclass[fclass["LP"] == "总计"].copy()
    team_fspec = fspec[fspec["LP"] == "总计"].copy()

    arch_agg = arch.groupby("小组", as_index=False).agg(
        在职数=("状态", lambda s: (s == "在职").sum()),
        总人数=("LP", "count"),
        平均入职月=("入职月份", "mean"),
    )
    ai_agg = ai.groupby("小组", as_index=False).agg(
        AI_首课任务总数=("首课_任务总数", "sum"),
        AI_首课干预中=("首课_干预中", "sum"),
        AI_首专任务总数=("首专_任务总数", "sum"),
        AI_首专干预中=("首专_干预中", "sum"),
    )
    ai_agg["AI_首课干预中占比"] = np.where(
        ai_agg["AI_首课任务总数"] > 0,
        ai_agg["AI_首课干预中"] / ai_agg["AI_首课任务总数"], np.nan)
    ai_agg["AI_首专干预中占比"] = np.where(
        ai_agg["AI_首专任务总数"] > 0,
        ai_agg["AI_首专干预中"] / ai_agg["AI_首专任务总数"], np.nan)

    team_rows = []
    # 海外团队（合）
    rec = {"层级": "汇总", "团队/小组": "海外团队（合）", "LP": "总计"}
    if not team_fc[team_fc["小组"] == "海外教学服务部"].empty:
        r = team_fc[team_fc["小组"] == "海外教学服务部"].iloc[0]
        for k in ["首通_新生数", "首通_拨通数", "首通_一家多娃新生数",
                  "首通_跟进率", "首通_接通率", "首通_及时跟进率",
                  "首通_生均接通时长", "首通_企微绑定率", "首通_48小时企微绑定率",
                  "首通_秒挂占比", "首通_follow率", "首通_生均外呼次数"]:
            rec[k] = r.get(k)
    if not team_fclass[team_fclass["小组_首课"] == "海外团队"].empty:
        r = team_fclass[team_fclass["小组_首课"] == "海外团队"].iloc[0]
        for k in ["首课_学员数", "首课_跟进率", "首课_及时跟进率", "首课_作业完成率"]:
            rec[k] = r.get(k)
    if not team_fspec[team_fspec["小组_首专"] == "海外团队"].empty:
        r = team_fspec[team_fspec["小组_首专"] == "海外团队"].iloc[0]
        for k in ["首专_学员数", "首专_跟进率", "首专_及时跟进率", "首专_作业完成率"]:
            rec[k] = r.get(k)
    if not team_sop[team_sop["小组"] == "海外团队"].empty:
        r = team_sop[team_sop["小组"] == "海外团队"].iloc[0]
        for k in ["命中首通新生数", "首通_邀请添加企微执行率", "首通_一家多娃问询执行率",
                  "首通_转介绍执行率", "首通_执行率加和", "首课_执行率加和", "服务池_语义点执行率加和"]:
            v = r.get(k)
            rec[("SOP_" + k) if k != "命中首通新生数" else k] = v
    if not arch_agg.empty:
        rec["在职数"] = int(arch_agg["在职数"].sum())
        rec["平均入职月"] = round(arch_agg["平均入职月"].mean(), 1) if pd.notna(arch_agg["平均入职月"].mean()) else None
    # AI 全团队
    if not ai_agg.empty:
        sk_t = ai_agg["AI_首课任务总数"].sum()
        sz_t = ai_agg["AI_首专任务总数"].sum()
        rec["AI_首课干预中占比"] = ai_agg["AI_首课干预中"].sum() / sk_t if sk_t > 0 else np.nan
        rec["AI_首专干预中占比"] = ai_agg["AI_首专干预中"].sum() / sz_t if sz_t > 0 else np.nan
    team_rows.append(rec)

    for team in TEAM_ORDER:
        rec = {"层级": "汇总", "团队/小组": team, "LP": "总计"}
        m = team_fc[team_fc["小组"] == team]
        if not m.empty:
            r = m.iloc[0]
            for k in ["首通_新生数", "首通_拨通数", "首通_一家多娃新生数",
                      "首通_跟进率", "首通_接通率", "首通_及时跟进率",
                      "首通_生均接通时长", "首通_企微绑定率", "首通_48小时企微绑定率",
                      "首通_秒挂占比", "首通_follow率", "首通_生均外呼次数"]:
                rec[k] = r.get(k)
        m = team_fclass[team_fclass["小组_首课"] == team]
        if not m.empty:
            r = m.iloc[0]
            for k in ["首课_学员数", "首课_跟进率", "首课_及时跟进率", "首课_作业完成率"]:
                rec[k] = r.get(k)
        m = team_fspec[team_fspec["小组_首专"] == team]
        if not m.empty:
            r = m.iloc[0]
            for k in ["首专_学员数", "首专_跟进率", "首专_及时跟进率", "首专_作业完成率"]:
                rec[k] = r.get(k)
        m = team_sop[team_sop["小组"] == team]
        if not m.empty:
            r = m.iloc[0]
            for k in ["命中首通新生数", "首通_邀请添加企微执行率", "首通_一家多娃问询执行率",
                      "首通_转介绍执行率", "首通_执行率加和", "首课_执行率加和", "服务池_语义点执行率加和"]:
                v = r.get(k)
                rec[("SOP_" + k) if k != "命中首通新生数" else k] = v
        m = arch_agg[arch_agg["小组"] == team]
        if not m.empty:
            r = m.iloc[0]
            rec["在职数"] = int(r["在职数"])
            rec["平均入职月"] = round(r["平均入职月"], 1) if pd.notna(r["平均入职月"]) else None
        m = ai_agg[ai_agg["小组"] == team]
        if not m.empty:
            r = m.iloc[0]
            rec["AI_首课干预中占比"] = r["AI_首课干预中占比"]
            rec["AI_首专干预中占比"] = r["AI_首专干预中占比"]
        team_rows.append(rec)

    summary = pd.DataFrame(team_rows)

    # ── 个人级明细（按 02_lp_first_class 的团队/LP 为基准）──
    base = fclass[fclass["LP"] != "总计"].rename(columns={"小组_首课": "小组"}).copy()
    base = base[base["小组"].isin(TEAM_ORDER)]
    spec_lp = fspec[fspec["LP"] != "总计"].rename(columns={"小组_首专": "小组"})
    base = base.merge(spec_lp, on=["小组", "LP"], how="left", suffixes=("", "_y"))
    fc_lp = fc[fc["LP"] != "总计"].copy()
    base = base.merge(fc_lp[["小组", "LP", "首通_新生数", "首通_拨通数",
                               "首通_一家多娃新生数", "首通_跟进率", "首通_接通率",
                               "首通_及时跟进率", "首通_生均接通时长", "首通_企微绑定率",
                               "首通_48小时企微绑定率", "首通_秒挂占比", "首通_follow率",
                               "首通_生均外呼次数"]],
                      on=["小组", "LP"], how="left")
    sop_lp = sop[(sop["LP"] != "总计")][["小组", "LP", "命中首通新生数",
                                          "首通_邀请添加企微执行率", "首通_一家多娃问询执行率",
                                          "首通_转介绍执行率", "首通_执行率加和",
                                          "首课_执行率加和", "服务池_语义点执行率加和"]].copy()
    sop_lp = sop_lp.rename(columns={
        "首通_邀请添加企微执行率": "SOP_首通_邀请添加企微执行率",
        "首通_一家多娃问询执行率": "SOP_首通_一家多娃问询执行率",
        "首通_转介绍执行率": "SOP_首通_转介绍执行率",
        "首通_执行率加和": "SOP_首通_执行率加和",
        "首课_执行率加和": "SOP_首课_执行率加和",
        "服务池_语义点执行率加和": "SOP_服务池_语义点执行率加和",
    })
    base = base.merge(sop_lp, on=["小组", "LP"], how="left")
    base = base.merge(arch[["小组", "LP", "入职时间", "入职月份", "入职时长", "状态", "职位1"]],
                      on=["小组", "LP"], how="left")
    # AI 个人
    base = base.merge(ai[["小组", "LP", "AI_首课干预中占比", "AI_首专干预中占比"]],
                      on=["小组", "LP"], how="left")

    base.insert(0, "层级", "个人")
    base = base.rename(columns={"小组": "团队/小组"})
    base["_order"] = base["团队/小组"].map({t: i for i, t in enumerate(TEAM_ORDER)})
    base = base.sort_values(["_order", "入职月份"], ascending=[True, False], na_position="last").drop(columns=["_order"])

    merged = pd.concat([summary, base], ignore_index=True)

    # ── 列顺序：严格对齐图片版式 ──
    col_order = [
        "层级", "团队/小组", "LP",
        # 到岗（团队级=在职数）
        "在职数",
        # 首通语义分析执行
        "命中首通新生数",
        "SOP_首通_邀请添加企微执行率",
        "SOP_首通_一家多娃问询执行率",
        "SOP_首通_转介绍执行率",
        "SOP_首通_执行率加和",
        # 首通
        "首通_新生数", "首通_拨通数",
        "首通_及时跟进率", "首通_生均接通时长",
        "首通_企微绑定率", "首通_秒挂占比",
        "首通_follow率", "首通_生均外呼次数",
        # 首课SOP执行
        "SOP_首课_执行率加和",
        # 首课
        "首课_学员数", "首课_跟进率", "首课_及时跟进率", "首课_作业完成率",
        # 首专
        "首专_学员数", "首专_跟进率", "首专_及时跟进率", "首专_作业完成率",
        # LP架构
        "入职时间", "入职月份", "状态",
        # 结论辅助列（不在格式化 Excel 中渲染，仅供 conclusions 使用）
        "首通_跟进率", "首通_接通率", "首通_48小时企微绑定率",
        "AI_首课干预中占比", "AI_首专干预中占比",
    ]
    cols = [c for c in col_order if c in merged.columns]
    merged = merged[cols]
    return merged


def main():
    import sys, io
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
    merged = build_merged()
    EXPORT_BASE.mkdir(parents=True, exist_ok=True)
    merged.to_excel(OUT_PATH, index=False)
    print(f"[OK] 输出: {OUT_PATH}")
    print(f"  rows={len(merged)}, cols={len(merged.columns)}")
    print()
    print("列名:", list(merged.columns))
    print()
    print("=== 团队级汇总（前 10 行） ===")
    pd.set_option("display.max_columns", 30)
    pd.set_option("display.width", 250)
    print(merged.head(10).to_string(index=False, max_colwidth=12))


if __name__ == "__main__":
    main()
