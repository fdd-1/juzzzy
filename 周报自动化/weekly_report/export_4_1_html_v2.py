"""4.1 HTML 周报 v2 — 五月第三周
数据源：BI 导出 exports/4_1_w3/
新增：首通语义点执行(5列)、首课语义点执行(2列)、LP架构(2列)
"""
from __future__ import annotations
import sys, io, math, html
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8")
import pandas as pd
import numpy as np
from pathlib import Path

ROOT = Path(__file__).parent
EXPORT_DIR = ROOT.parent / "exports" / "4_1_w3"
FIRST_CALL_FILE = EXPORT_DIR / "04_first_call" / "益智海外新生首通监控.xlsx"
SERVICE_FILE = EXPORT_DIR / "02_lp_service" / "海外思维学管服务指标统计表.xlsx"
SOP_FILE = EXPORT_DIR / "05_sop" / "海外思维服务SOP执行情况.xlsx"
LP_ARCH_FILE = EXPORT_DIR / "06_lp_arch" / "海外思维LP架构表.xlsx"
AI_FILE = EXPORT_DIR / "07_ai_summary" / "AI学情助手完成情况汇总（验收中）.xlsx"
OUT_HTML = EXPORT_DIR / "4_1_周报_v3.html"

TEAM_ORDER = ["港澳1组", "港澳2组", "港澳组", "美澳1组", "美澳2组", "美澳3组", "美澳4组", "美澳5组"]
EXCLUDE_TEAMS = {"台湾组"}
TARGETS = {
    "首通_及时跟进率": 0.95,
    "首课_及时跟进率": 0.85,
    "首专_及时跟进率": 0.85,
}

# ── 数据加载 ──

def load_first_call():
    df = pd.read_excel(FIRST_CALL_FILE, header=2)
    df = df.rename(columns={"Unnamed: 1": "团队", "Unnamed: 2": "主管", "Unnamed: 3": "LP"})
    df = df.drop(columns=["Unnamed: 0"], errors="ignore")
    team = df.iloc[:10].copy()
    team["团队"] = team["团队"].ffill()
    team = team[team["LP"] == "总计"].copy()
    team["层级"] = "汇总"
    personal = df.iloc[11:].copy()
    personal["团队"] = personal["团队"].ffill()
    personal["主管"] = personal["主管"].ffill()
    personal = personal.dropna(subset=["LP"])
    personal = personal[personal["LP"] != "总计"]
    personal["层级"] = "个人"
    cols_keep = ["层级", "团队", "LP", "新生数", "勿扰新生数", "一家多娃新生数",
                 "跟进率", "及时跟进率", "企微绑定率", "秒挂占比"]
    for c in cols_keep:
        for d in [team, personal]:
            if c not in d.columns:
                d[c] = np.nan
    team, personal = team[cols_keep].copy(), personal[cols_keep].copy()
    rn = {"新生数": "首通_新生数", "勿扰新生数": "首通_勿扰人数",
          "一家多娃新生数": "首通_一家多娃数", "跟进率": "首通_跟进率",
          "及时跟进率": "首通_及时跟进率", "企微绑定率": "首通_企微绑定率",
          "秒挂占比": "首通_秒挂占比"}
    return pd.concat([team.rename(columns=rn), personal.rename(columns=rn)], ignore_index=True)


def load_service():
    df = pd.read_excel(SERVICE_FILE, header=1)
    df = df.rename(columns={"团队": "团队", "LP名称": "LP"})
    df = df[~df["团队"].astype(str).str.startswith("口径")].copy()
    df = df[~df["团队"].astype(str).str.match(r"^\d+）")].copy()
    df = df[~df["团队"].astype(str).str.startswith("注意")].copy()
    df = df[~df["团队"].astype(str).str.startswith("【")].copy()
    team = df[df["LP"] == "总计"].copy()
    team["团队"] = team["团队"].ffill()
    team["层级"] = "汇总"
    personal = df[df["LP"] != "总计"].copy()
    personal = personal.dropna(subset=["LP"])
    personal["团队"] = personal["团队"].ffill()
    personal["层级"] = "个人"
    cols_keep = ["层级", "团队", "LP",
                 "首课学员数", "首课勿扰学员数", "首课跟进率", "首课及时跟进率",
                 "首专学员数", "首专勿扰学员数", "首专跟进率", "首专及时跟进率"]
    for c in cols_keep:
        for d in [team, personal]:
            if c not in d.columns:
                d[c] = np.nan
    team, personal = team[cols_keep].copy(), personal[cols_keep].copy()
    rn = {"首课学员数": "首课_新生数", "首课勿扰学员数": "首课_勿扰人数",
          "首课跟进率": "首课_跟进率", "首课及时跟进率": "首课_及时跟进率",
          "首专学员数": "首专_新生数", "首专勿扰学员数": "首专_勿扰人数",
          "首专跟进率": "首专_跟进率", "首专及时跟进率": "首专_及时跟进率"}
    return pd.concat([team.rename(columns=rn), personal.rename(columns=rn)], ignore_index=True)


def load_sop():
    raw = pd.read_excel(SOP_FILE, header=None)
    data_rows = []
    for i in range(4, len(raw)):
        row = raw.iloc[i]
        lp_val = row.iloc[3]
        if pd.isna(lp_val) or str(lp_val).strip() == "":
            continue
        team_val = row.iloc[1]
        team_str = str(team_val) if not pd.isna(team_val) else None
        data_rows.append({
            "团队_raw": team_str,
            "LP": str(lp_val).strip(),
            "SOP首通_执行率加和": row.iloc[13],
            "SOP首通_拨通新生数": row.iloc[4],
            "SOP首通_企微执行率": row.iloc[10],
            "SOP首通_一家多娃执行率": row.iloc[11],
            "SOP首通_转介绍执行率": row.iloc[12],
            "SOP首课_执行率加和": row.iloc[23],
            "SOP首课_上课感受执行率": row.iloc[14],
        })
    sop = pd.DataFrame(data_rows)
    sop["团队_raw"] = sop["团队_raw"].ffill()
    sop["层级"] = sop["LP"].apply(lambda x: "汇总" if x == "总计" else "个人")
    sop["团队"] = sop["团队_raw"].replace({"海外团队": "海外团队（合）"})
    sop = sop.drop(columns=["团队_raw"])
    return sop


def load_lp_arch():
    raw = pd.read_excel(LP_ARCH_FILE, header=None)
    rows = []
    for i in range(4, len(raw)):
        row = raw.iloc[i]
        name = row.iloc[6]
        if pd.isna(name) or str(name).strip() == "":
            continue
        team = str(row.iloc[3]) if not pd.isna(row.iloc[3]) else ""
        hire_date = row.iloc[9]
        tenure_group = row.iloc[13]
        rows.append({
            "团队": team,
            "LP": str(name).strip(),
            "LP_入职时间": hire_date,
            "LP_入职时长分组": tenure_group,
        })
    arch = pd.DataFrame(rows)
    arch["团队"] = arch["团队"].ffill()
    return arch


def load_ai():
    df = pd.read_excel(AI_FILE, header=3)
    df = df.rename(columns={"班主任姓名": "LP", "LP小组": "团队"})
    df["首课_干预中占比"] = np.where(df["首课_任务总数"] > 0, df["首课_干预中"] / df["首课_任务总数"], np.nan)
    df["首专_干预中占比"] = np.where(df["首专_任务总数"] > 0, df["首专_干预中"] / df["首专_任务总数"], np.nan)
    return df


def build_ai_table():
    ai = load_ai()
    teams_in_order = ["海外团队（合）"] + TEAM_ORDER
    rows = []
    rec = {
        "层级": "汇总", "团队": "海外团队（合）", "LP": "总计",
        "AI_任务总数": int(ai["任务总数"].sum()),
        "AI_覆盖学员数": int(ai["覆盖学员数"].sum()),
        "AI_首课_任务总数": int(ai["首课_任务总数"].sum()),
        "AI_首课_干预中": int(ai["首课_干预中"].sum()),
        "AI_首专_任务总数": int(ai["首专_任务总数"].sum()),
        "AI_首专_干预中": int(ai["首专_干预中"].sum()),
    }
    rec["AI_首课_干预中占比"] = rec["AI_首课_干预中"] / rec["AI_首课_任务总数"] if rec["AI_首课_任务总数"] > 0 else np.nan
    rec["AI_首专_干预中占比"] = rec["AI_首专_干预中"] / rec["AI_首专_任务总数"] if rec["AI_首专_任务总数"] > 0 else np.nan
    rec["AI_人均任务数"] = rec["AI_任务总数"] / rec["AI_覆盖学员数"] if rec["AI_覆盖学员数"] > 0 else np.nan
    rows.append(rec)

    for team in TEAM_ORDER:
        sub = ai[ai["团队"] == team]
        if sub.empty:
            continue
        rec = {
            "层级": "汇总", "团队": team, "LP": "总计",
            "AI_任务总数": int(sub["任务总数"].sum()),
            "AI_覆盖学员数": int(sub["覆盖学员数"].sum()),
            "AI_首课_任务总数": int(sub["首课_任务总数"].sum()),
            "AI_首课_干预中": int(sub["首课_干预中"].sum()),
            "AI_首专_任务总数": int(sub["首专_任务总数"].sum()),
            "AI_首专_干预中": int(sub["首专_干预中"].sum()),
        }
        rec["AI_首课_干预中占比"] = rec["AI_首课_干预中"] / rec["AI_首课_任务总数"] if rec["AI_首课_任务总数"] > 0 else np.nan
        rec["AI_首专_干预中占比"] = rec["AI_首专_干预中"] / rec["AI_首专_任务总数"] if rec["AI_首专_任务总数"] > 0 else np.nan
        rec["AI_人均任务数"] = rec["AI_任务总数"] / rec["AI_覆盖学员数"] if rec["AI_覆盖学员数"] > 0 else np.nan
        rows.append(rec)

    detail = ai[ai["团队"].isin(TEAM_ORDER)].copy()
    detail = detail.rename(columns={
        "任务总数": "AI_任务总数", "覆盖学员数": "AI_覆盖学员数", "人均任务数": "AI_人均任务数",
        "首课_任务总数": "AI_首课_任务总数", "首课_干预中": "AI_首课_干预中",
        "首课_干预中占比": "AI_首课_干预中占比",
        "首专_任务总数": "AI_首专_任务总数", "首专_干预中": "AI_首专_干预中",
        "首专_干预中占比": "AI_首专_干预中占比",
    })
    detail["层级"] = "个人"
    cols_keep = ["层级", "团队", "LP",
                 "AI_任务总数", "AI_覆盖学员数", "AI_人均任务数",
                 "AI_首课_任务总数", "AI_首课_干预中", "AI_首课_干预中占比",
                 "AI_首专_任务总数", "AI_首专_干预中", "AI_首专_干预中占比"]
    detail = detail[[c for c in cols_keep if c in detail.columns]]

    summary_df = pd.DataFrame(rows)
    merged = pd.concat([summary_df, detail], ignore_index=True)
    team_map = {t: i for i, t in enumerate(teams_in_order)}
    merged["_to"] = merged["团队"].map(team_map).fillna(99)
    merged["_lo"] = merged["层级"].map({"汇总": 0, "个人": 1})
    merged = merged.sort_values(["_lo", "_to"], ascending=True).drop(columns=["_to", "_lo"])
    merged = merged.reset_index(drop=True)
    return merged


def build_merged():
    fc = load_first_call()
    svc = load_service()
    sop = load_sop()
    arch = load_lp_arch()

    fc["团队"] = fc["团队"].replace({"海外教学服务部": "海外团队（合）"})
    svc["团队"] = svc["团队"].replace({"海外团队": "海外团队（合）"})

    merged = fc.merge(svc, on=["层级", "团队", "LP"], how="outer")
    merged = merged.merge(sop, on=["层级", "团队", "LP"], how="left")
    # LP架构只有个人级
    arch_personal = arch.copy()
    arch_personal["层级"] = "个人"
    merged = merged.merge(arch_personal, on=["层级", "团队", "LP"], how="left")

    team_map = {t: i for i, t in enumerate(["海外团队（合）"] + TEAM_ORDER + ["台湾组"])}
    merged["_to"] = merged["团队"].map(team_map).fillna(99)
    merged["_lo"] = merged["层级"].map({"汇总": 0, "个人": 1})
    merged = merged.sort_values(["_lo", "_to"], ascending=True).drop(columns=["_to", "_lo"])
    merged = merged.reset_index(drop=True)
    return merged


# ── 结论生成 ──

def generate_conclusions(df: pd.DataFrame) -> str:
    lines = []
    summary = df[df["层级"] == "汇总"]
    detail = df[df["层级"] == "个人"]
    overall = summary[summary["团队"] == "海外团队（合）"].iloc[0] if len(summary[summary["团队"] == "海外团队（合）"]) > 0 else None

    indicators = [
        ("首通", "首通_跟进率", "首通_及时跟进率", "首通_勿扰人数", "首通_新生数", 0.95),
        ("首课", "首课_跟进率", "首课_及时跟进率", "首课_勿扰人数", "首课_新生数", 0.85),
        ("首专", "首专_跟进率", "首专_及时跟进率", "首专_勿扰人数", "首专_新生数", 0.85),
    ]
    for label, follow_col, timely_col, dnd_col, total_col, target in indicators:
        if overall is not None and not pd.isna(overall.get(timely_col)):
            overall_val = overall[timely_col]
            overall_follow = overall.get(follow_col, np.nan)
            reached = overall_val >= target
            follow_str = f"总体跟进率 {overall_follow*100:.2f}%；" if not pd.isna(overall_follow) else ""
            status = "达到" if reached else "未达"
            lines.append(f'<p><b>{label}：{follow_str}及时跟进率 {overall_val*100:.2f}%，'
                        f'<span class="{"text-green" if reached else "text-red"}">{status}本月目标（{int(target*100)}%）</span></b></p>')
            teams_below = []
            for _, row in summary.iterrows():
                team = row["团队"]
                if team == "海外团队（合）" or team in EXCLUDE_TEAMS:
                    continue
                val = row.get(timely_col)
                if pd.isna(val):
                    continue
                if val < target:
                    teams_below.append((team, val))
            teams_below.sort(key=lambda x: x[1])
            for team, val in teams_below:
                lines.append(f'<p><span class="hl-pink">{team}</span> 及时跟进率 '
                           f'<span class="hl-pink">{val*100:.1f}%</span>，未达目标</p>')
                team_detail = detail[detail["团队"] == team]
                lp_below = []
                for _, lp_row in team_detail.iterrows():
                    lp_val = lp_row.get(timely_col)
                    if pd.isna(lp_val):
                        continue
                    if lp_val < target:
                        dnd = lp_row.get(dnd_col)
                        total = lp_row.get(total_col)
                        lp_below.append((lp_row["LP"], lp_val, dnd, total))
                lp_below.sort(key=lambda x: x[1])
                if lp_below:
                    parts = []
                    for name, v, dnd, total in lp_below[:5]:
                        s = f"{name}({v*100:.0f}%"
                        if not pd.isna(dnd) and float(dnd) > 0:
                            s += f"，含勿扰{int(dnd)}人"
                        s += ")"
                        parts.append(s)
                    lines.append(f'<p class="indent">└ 需关注：{" / ".join(parts)}</p>')
                    has_dnd = any((not pd.isna(d)) and float(d) > 0 for _, _, d, _ in lp_below[:5])
                    if has_dnd:
                        lines.append(f'<p class="indent" style="color:#888;">注：勿扰学员可在评估中剔除其干预归属，数据未做剔除</p>')
        lines.append("<p>&nbsp;</p>")

    if overall is not None and not pd.isna(overall.get("首通_企微绑定率")):
        lines.append(f'<p><b>企微绑定率：整体 {overall["首通_企微绑定率"]*100:.1f}%</b></p>')
        teams_low = []
        for _, row in summary.iterrows():
            team = row["团队"]
            if team == "海外团队（合）" or team in EXCLUDE_TEAMS:
                continue
            val = row.get("首通_企微绑定率")
            if pd.isna(val):
                continue
            if val < overall["首通_企微绑定率"]:
                teams_low.append((team, val))
        teams_low.sort(key=lambda x: x[1])
        if teams_low:
            for team, val in teams_low[:3]:
                lines.append(f'<p><span class="hl-pink">{team}</span> 企微绑定率 '
                           f'<span class="hl-pink">{val*100:.1f}%</span>，低于整体</p>')
    return "\n".join(lines)


def generate_ai_conclusions(ai_df: pd.DataFrame) -> str:
    lines = []
    summary = ai_df[ai_df["层级"] == "汇总"]
    detail = ai_df[ai_df["层级"] == "个人"]
    overall = summary[summary["团队"] == "海外团队（合）"].iloc[0] if len(summary[summary["团队"] == "海外团队（合）"]) > 0 else None

    if overall is None:
        return "<p>无 AI 数据</p>"

    total_tasks = overall.get("AI_任务总数")
    cover = overall.get("AI_覆盖学员数")
    per_capita = overall.get("AI_人均任务数")
    fc_rate = overall.get("AI_首课_干预中占比")
    fz_rate = overall.get("AI_首专_干预中占比")

    parts = [f'AI 任务总数 <b>{int(total_tasks)}</b>',
             f'覆盖学员 <b>{int(cover)}</b> 人',
             f'人均 <b>{per_capita:.2f}</b> 任务']
    lines.append(f'<p><b>整体：{"，".join(parts)}</b></p>')

    fc_str = f'{fc_rate*100:.1f}%' if not pd.isna(fc_rate) else 'N/A'
    fz_str = f'{fz_rate*100:.1f}%' if not pd.isna(fz_rate) else 'N/A'
    lines.append(f'<p>首课干预中占比 <b>{fc_str}</b>；首专干预中占比 <b>{fz_str}</b></p>')

    for label, col in [("首课", "AI_首课_干预中占比"), ("首专", "AI_首专_干预中占比")]:
        overall_val = overall.get(col)
        if pd.isna(overall_val):
            continue
        teams_high = []
        for _, row in summary.iterrows():
            team = row["团队"]
            if team == "海外团队（合）" or team in EXCLUDE_TEAMS:
                continue
            val = row.get(col)
            if pd.isna(val):
                continue
            if val > overall_val:
                teams_high.append((team, val))
        teams_high.sort(key=lambda x: -x[1])
        if teams_high:
            top = teams_high[:3]
            text = " / ".join([f'{t}({v*100:.1f}%)' for t, v in top])
            lines.append(f'<p>{label}干预中占比高于整体的小组：<span class="hl-pink">{text}</span></p>')

    no_task = []
    for _, row in detail.iterrows():
        if pd.isna(row.get("AI_任务总数")) or float(row["AI_任务总数"]) == 0:
            no_task.append((row["团队"], row["LP"]))
    if no_task:
        sample = " / ".join([f"{t}-{n}" for t, n in no_task[:8]])
        lines.append(f'<p>本周无 AI 任务的 LP（共 {len(no_task)} 人）：{sample}{"..." if len(no_task) > 8 else ""}</p>')

    return "\n".join(lines)


# ── HTML 渲染 ──

GROUPS = [
    ("基础", ["层级", "团队", "LP"], "D9D9D9", False),
    ("首通语义点执行", ["SOP首通_执行率加和", "SOP首通_拨通新生数", "SOP首通_企微执行率",
                  "SOP首通_一家多娃执行率", "SOP首通_转介绍执行率"], "E2EFDA", True),
    ("首通", ["首通_新生数", "首通_勿扰人数", "首通_一家多娃数",
              "首通_跟进率", "首通_及时跟进率", "首通_企微绑定率", "首通_秒挂占比"], "BDD7EE", True),
    ("首课语义点执行", ["SOP首课_执行率加和", "SOP首课_上课感受执行率"], "FFF2CC", True),
    ("首课", ["首课_新生数", "首课_勿扰人数", "首课_跟进率", "首课_及时跟进率"], "C6E0B4", True),
    ("首专", ["首专_新生数", "首专_勿扰人数", "首专_跟进率", "首专_及时跟进率"], "F4B084", True),
    ("LP架构", ["LP_入职时间", "LP_入职时长分组"], "D6DCE4", True),
]

AI_GROUPS = [
    ("基础", ["层级", "团队", "LP"], "D9D9D9", False),
    ("总览", ["AI_任务总数", "AI_覆盖学员数", "AI_人均任务数"], "BDD7EE", True),
    ("首课AI干预", ["AI_首课_任务总数", "AI_首课_干预中", "AI_首课_干预中占比"], "C6E0B4", True),
    ("首专AI干预", ["AI_首专_任务总数", "AI_首专_干预中", "AI_首专_干预中占比"], "F4B084", True),
]

COL_LABELS = {
    "层级": "层级", "团队": "团队/小组", "LP": "LP姓名",
    "SOP首通_执行率加和": "执行率加和", "SOP首通_拨通新生数": "拨通新生数",
    "SOP首通_企微执行率": "企微/WS/Line执行率",
    "SOP首通_一家多娃执行率": "一家多娃执行率", "SOP首通_转介绍执行率": "转介绍执行率",
    "首通_新生数": "新生数", "首通_勿扰人数": "勿扰人数", "首通_一家多娃数": "一家多娃数",
    "首通_跟进率": "跟进率", "首通_及时跟进率": "及时跟进率",
    "首通_企微绑定率": "企微绑定率", "首通_秒挂占比": "秒挂占比",
    "SOP首课_执行率加和": "执行率加和", "SOP首课_上课感受执行率": "上课感受执行率",
    "首课_新生数": "新生数", "首课_勿扰人数": "勿扰人数",
    "首课_跟进率": "跟进率", "首课_及时跟进率": "及时跟进率",
    "首专_新生数": "新生数", "首专_勿扰人数": "勿扰人数",
    "首专_跟进率": "跟进率", "首专_及时跟进率": "及时跟进率",
    "LP_入职时间": "入职时间", "LP_入职时长分组": "入职时长分组",
    "AI_任务总数": "任务总数", "AI_覆盖学员数": "覆盖学员数", "AI_人均任务数": "人均任务数",
    "AI_首课_任务总数": "任务总数", "AI_首课_干预中": "干预中",
    "AI_首课_干预中占比": "干预中占比",
    "AI_首专_任务总数": "任务总数", "AI_首专_干预中": "干预中",
    "AI_首专_干预中占比": "干预中占比",
}

PCT_COLS = {
    "SOP首通_企微执行率", "SOP首通_一家多娃执行率", "SOP首通_转介绍执行率",
    "SOP首课_上课感受执行率",
    "首通_跟进率", "首通_及时跟进率", "首通_企微绑定率", "首通_秒挂占比",
    "首课_跟进率", "首课_及时跟进率",
    "首专_跟进率", "首专_及时跟进率",
    "AI_首课_干预中占比", "AI_首专_干预中占比",
}
INT_COLS = {
    "SOP首通_拨通新生数",
    "首通_新生数", "首通_勿扰人数", "首通_一家多娃数",
    "首课_新生数", "首课_勿扰人数",
    "首专_新生数", "首专_勿扰人数",
    "AI_任务总数", "AI_覆盖学员数",
    "AI_首课_任务总数", "AI_首课_干预中",
    "AI_首专_任务总数", "AI_首专_干预中",
}
FLOAT1_COLS = {"SOP首通_执行率加和", "SOP首课_执行率加和"}
FLOAT2_COLS = {"AI_人均任务数"}
DATE_COLS = {"LP_入职时间"}


def fmt_cell(val, col):
    if val is None or (isinstance(val, float) and math.isnan(val)) or pd.isna(val):
        return ""
    if col in PCT_COLS:
        try:
            return f"{float(val)*100:.1f}%"
        except Exception:
            return html.escape(str(val))
    if col in INT_COLS:
        try:
            return f"{int(float(val))}"
        except Exception:
            return html.escape(str(val))
    if col in FLOAT1_COLS:
        try:
            return f"{float(val):.2f}"
        except Exception:
            return html.escape(str(val))
    if col in FLOAT2_COLS:
        try:
            return f"{float(val):.2f}"
        except Exception:
            return html.escape(str(val))
    if col in DATE_COLS:
        try:
            ts = pd.Timestamp(val)
            return ts.strftime("%Y-%m-%d") if not pd.isna(ts) else ""
        except Exception:
            return html.escape(str(val).split(" ")[0])
    return html.escape(str(val))


def render_table(df: pd.DataFrame, groups=None) -> str:
    if groups is None:
        groups = GROUPS
    cols = []
    for _, gcols, _, _ in groups:
        for c in gcols:
            if c in df.columns:
                cols.append(c)
    df = df[cols].copy()
    freeze_cols = 3

    h1_cells = []
    for gname, gcols, gcolor, is_data in groups:
        present = [c for c in gcols if c in df.columns]
        if not present:
            continue
        bg = "4472C4" if is_data else gcolor
        fg = "#fff" if is_data else "#000"
        h1_cells.append(
            f'<th colspan="{len(present)}" class="gh" '
            f'style="background:#{bg};color:{fg}">{html.escape(gname)}</th>')

    h2_cells = []
    col_to_gc = {}
    for _, gcols, gcolor, _ in groups:
        for c in gcols:
            col_to_gc[c] = gcolor
    for i, c in enumerate(df.columns):
        bg = col_to_gc.get(c, "FFFFFF")
        frozen_cls = " frozen" if i < freeze_cols else ""
        h2_cells.append(
            f'<th class="ch{frozen_cls}" style="background:#{bg}">'
            f'{html.escape(COL_LABELS.get(c, c))}</th>')

    body_rows = []
    for ri in range(len(df)):
        row = df.iloc[ri]
        is_summary = row.get("层级") == "汇总"
        tr_cls = "summary" if is_summary else "detail"
        tds = []
        for i, c in enumerate(df.columns):
            frozen_cls = " frozen" if i < freeze_cols else ""
            text = fmt_cell(row[c], c)
            tds.append(f'<td class="{frozen_cls.strip()}">{text}</td>')
        body_rows.append(f'<tr class="{tr_cls}">{"".join(tds)}</tr>')

    return (
        f'<div class="table-wrap"><table class="report-table">'
        f'<thead>'
        f'<tr class="group-h">{"".join(h1_cells)}</tr>'
        f'<tr class="col-row">{"".join(h2_cells)}</tr>'
        f'</thead>'
        f'<tbody>{"".join(body_rows)}</tbody>'
        f'</table></div>'
    )


def main():
    merged = build_merged()
    conclusions = generate_conclusions(merged)
    table_html = render_table(merged)

    ai_merged = build_ai_table()
    ai_conclusions = generate_ai_conclusions(ai_merged)
    ai_table_html = render_table(ai_merged, AI_GROUPS)

    page = (HTML_TEMPLATE
            .replace("<!--CONCLUSIONS-->", conclusions)
            .replace("<!--TABLE-->", table_html)
            .replace("<!--AI_CONCLUSIONS-->", ai_conclusions)
            .replace("<!--AI_TABLE-->", ai_table_html))
    OUT_HTML.parent.mkdir(parents=True, exist_ok=True)
    OUT_HTML.write_text(page, encoding="utf-8")
    print(f"[OK] HTML: {OUT_HTML}")
    print(f"  service rows={len(merged)}, ai rows={len(ai_merged)}")


HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="zh-CN">
<head>
<meta charset="utf-8">
<title>4.1 服务指标跟进 — 周报（5.19-5.25）</title>
<style>
  :root {
    --border: #cfd5dc;
    --text: #1f2329;
    --muted: #646a73;
    --summary-bg: #FFF2CC;
    --pink: #FBBFBC;
    --red: #D83931;
    --green: #2ea121;
  }
  * { box-sizing: border-box; }
  body {
    font-family: "Microsoft YaHei", "微软雅黑", -apple-system, sans-serif;
    margin: 24px 32px;
    color: var(--text);
    background: #f7f8fa;
  }
  h1 { font-size: 20px; margin: 0 0 6px; }
  h2 { font-size: 16px; margin: 24px 0 10px; padding-left: 8px; border-left: 4px solid #4472C4; }
  .meta { color: var(--muted); font-size: 12px; margin-bottom: 18px; }
  .callout {
    display: flex; gap: 12px;
    background: rgb(255,245,235); border: 1px solid rgb(254,212,164);
    border-radius: 6px; padding: 14px 16px; margin-bottom: 20px;
    line-height: 1.8; font-size: 14px;
  }
  .callout-emoji { font-size: 18px; line-height: 1.6; }
  .callout-body p { margin: 0 0 4px; }
  .callout-body p.indent { padding-left: 20px; color: #555; }
  .text-red { color: var(--red); }
  .text-green { color: var(--green); }
  .hl-pink { background: var(--pink); padding: 0 2px; border-radius: 2px; }
  .table-wrap {
    overflow-x: auto; overflow-y: auto;
    max-height: 80vh;
    background: #fff; border-radius: 4px;
    box-shadow: 0 1px 3px rgba(0,0,0,0.06);
    margin-bottom: 24px;
    position: relative;
  }
  table.report-table {
    border-collapse: separate; border-spacing: 0;
    width: max-content; min-width: 100%; font-size: 12px;
  }
  table.report-table th, table.report-table td {
    border: 1px solid var(--border);
    padding: 6px 8px; text-align: center;
    vertical-align: middle; white-space: nowrap;
  }
  table.report-table thead { position: sticky; top: 0; z-index: 10; }
  table.report-table thead tr.group-h th {
    height: 28px; font-weight: 700; font-size: 13px;
    position: sticky; top: 0; z-index: 10;
  }
  table.report-table thead tr.col-row th {
    height: 42px; font-weight: 700; color: #1f2329;
    white-space: normal; line-height: 1.35; min-width: 64px;
    position: sticky; top: 29px; z-index: 10;
  }
  table.report-table th.frozen, table.report-table td.frozen {
    position: sticky; z-index: 5; background: #fff;
  }
  table.report-table thead th.frozen { z-index: 15; }
  table.report-table th:nth-child(1), table.report-table td:nth-child(1) { left: 0; min-width: 50px; }
  table.report-table th:nth-child(2), table.report-table td:nth-child(2) { left: 50px; min-width: 90px; }
  table.report-table th:nth-child(3), table.report-table td:nth-child(3) { left: 140px; min-width: 70px; }
  table.report-table tbody tr.detail td { background: #fff; }
  table.report-table tbody tr.summary td { background: var(--summary-bg) !important; font-weight: 700; }
  table.report-table tbody tr.summary td.frozen { background: var(--summary-bg) !important; }
  table.report-table tbody tr:hover td { background: #eef4ff; }
  table.report-table tbody tr.summary:hover td { background: #fce8a3; }
</style>
</head>
<body>
  <h1>4.1 服务指标跟进</h1>
  <div class="meta">数据周期：5.19–5.25（第三自然周）｜ 来源：BI Smartbi 导出</div>
  <h2>结论</h2>
  <div class="callout">
    <div class="callout-emoji">❗</div>
    <div class="callout-body"><!--CONCLUSIONS--></div>
  </div>
  <h2>服务指标数据表</h2>
  <!--TABLE-->
  <h2>AI 学情跟进汇总</h2>
  <div class="meta">数据日期：2026-05-24（截止当日累计快照）｜ 报表：AI学情助手完成情况汇总（验收中）</div>
  <div class="callout">
    <div class="callout-emoji">🤖</div>
    <div class="callout-body"><!--AI_CONCLUSIONS--></div>
  </div>
  <!--AI_TABLE-->
</body>
</html>
"""


if __name__ == "__main__":
    main()
