"""4.1 板块结论生成器
基于 _merged_4_1.xlsx 的真实数据,自动生成 callout 格式的 XML 结论。

结论模板(参考飞书周报文档):
- 首通: 跟进率/及时跟进率(目标95%)/企微绑定率/秒挂占比 + 落后组+落后LP
- 首课: 跟进率/及时跟进率(目标85%) + 落后组+落后LP
- 首专: 跟进率/及时跟进率(目标85%) + 落后组+落后LP
- 语义分析: 添加企微/一家多娃问询/转介绍 执行率 + 落后组
"""
from __future__ import annotations
from pathlib import Path
import pandas as pd

# 落后判定阈值
TARGETS = {
    "首通_及时跟进率": 0.95,
    "首课及时跟进率": 0.85,
    "首专及时跟进率": 0.85,
}

# 排除的团队(不进入结论但保留数据)
EXCLUDE_TEAMS = {"台湾组"}

# LP 过滤门槛(新生数过少不算落后)
LP_MIN_NEW_COUNT = 3


def _get_summary_row(df: pd.DataFrame) -> pd.Series | None:
    """获取整体汇总行(LP=='总计' 且 团队 包含 '海外团队' 或 类似)。"""
    if "LP" not in df.columns:
        return None
    # 找到 团队=='海外团队' 且 LP=='总计' 的行
    mask = (df["LP"] == "总计")
    if "团队" in df.columns:
        # 优先海外团队
        ov_mask = mask & (df["团队"] == "海外团队")
        if ov_mask.any():
            return df[ov_mask].iloc[0]
    if mask.any():
        return df[mask].iloc[0]
    return None


def _get_team_rows(df: pd.DataFrame) -> pd.DataFrame:
    """获取分组汇总行(LP=='总计' 但不是海外团队)。"""
    if "LP" not in df.columns or "团队" not in df.columns:
        return pd.DataFrame()
    mask = (df["LP"] == "总计") & (df["团队"] != "海外团队")
    return df[mask].copy()


def _fmt_pct(v) -> str:
    """格式化为百分比字符串。"""
    try:
        return f"{float(v) * 100:.2f}%"
    except Exception:
        return "—"


def _laggard_teams(team_df: pd.DataFrame, col: str, threshold: float, exclude: set = None) -> list:
    """找出某列低于阈值的团队列表 [(团队名, 值), ...]。"""
    exclude = exclude or set()
    laggards = []
    for _, row in team_df.iterrows():
        team = row.get("团队", "")
        if team in exclude:
            continue
        v = row.get(col)
        if pd.isna(v):
            continue
        try:
            v_f = float(v)
            if v_f < threshold:
                laggards.append((team, v_f))
        except Exception:
            pass
    return sorted(laggards, key=lambda x: x[1])


def generate_4_1_callout(merged_path: Path) -> str:
    """生成 4.1 板块的 callout XML 结论。"""
    df = pd.read_excel(merged_path)
    summary = _get_summary_row(df)
    teams = _get_team_rows(df)

    lines = ['<callout emoji="❗">']

    if summary is not None:
        # 首通
        lines.append("<p><b>跟进:</b></p>")

        ft_follow = summary.get("首通_跟进率")
        ft_timely = summary.get("首通_及时跟进率")
        if ft_follow is not None and ft_timely is not None:
            timely_v = float(ft_timely) if not pd.isna(ft_timely) else 0
            target = TARGETS["首通_及时跟进率"]
            achieved = "达本月目标" if timely_v >= target else f'<span text-color="rgb(216,57,49)">未达本月目标</span>'
            lines.append(
                f"<p><b>首通:总体跟进率{_fmt_pct(ft_follow)};及时跟进率{_fmt_pct(ft_timely)},{achieved}(95%)</b></p>"
            )
            # 落后组
            laggards = _laggard_teams(teams, "首通_及时跟进率", target, EXCLUDE_TEAMS)
            for team, v in laggards[:3]:
                lines.append(f'<p><span background-color="rgb(251,191,188)">{team}</span>及时跟进率仅{_fmt_pct(v)},需注意</p>')

        # 企微绑定率
        wq = summary.get("首通_企微绑定率")
        if wq is not None:
            lines.append(f"<p>--<b>企微绑定率</b>整体为<b>{_fmt_pct(wq)}</b></p>")

        # 秒挂占比
        mh = summary.get("首通_秒挂占比")
        if mh is not None:
            lines.append(f"<p>--<b>秒挂</b>占比为{_fmt_pct(mh)}</p>")

        # 首课
        sk_follow = summary.get("首课跟进率")
        sk_timely = summary.get("首课及时跟进率")
        if sk_follow is not None and sk_timely is not None:
            t_v = float(sk_timely) if not pd.isna(sk_timely) else 0
            target = TARGETS["首课及时跟进率"]
            achieved = "达本月及时跟进率目标" if t_v >= target else f'<span text-color="rgb(216,57,49)">未达本月及时跟进率目标</span>'
            lines.append(
                f"<p><b>首课:总体跟进率{_fmt_pct(sk_follow)},及时跟进率{_fmt_pct(sk_timely)},{achieved}(85%)</b></p>"
            )
            laggards = _laggard_teams(teams, "首课及时跟进率", target, EXCLUDE_TEAMS)
            for team, v in laggards[:3]:
                lines.append(f'<p><span background-color="rgb(251,191,188)">{team}</span>及时跟进率仅{_fmt_pct(v)},需注意</p>')

        # 首专
        sz_follow = summary.get("首专跟进率")
        sz_timely = summary.get("首专及时跟进率")
        if sz_follow is not None and sz_timely is not None:
            t_v = float(sz_timely) if not pd.isna(sz_timely) else 0
            target = TARGETS["首专及时跟进率"]
            achieved = "达本月及时跟进率目标" if t_v >= target else f'<span text-color="rgb(216,57,49)">未达本月及时跟进率目标</span>'
            lines.append(
                f"<p><b>首专:总体跟进率{_fmt_pct(sz_follow)},及时跟进率{_fmt_pct(sz_timely)},{achieved}(85%)</b></p>"
            )
            laggards = _laggard_teams(teams, "首专及时跟进率", target, EXCLUDE_TEAMS)
            for team, v in laggards[:3]:
                lines.append(f'<p><span background-color="rgb(251,191,188)">{team}</span>及时跟进率仅{_fmt_pct(v)},需注意</p>')

        # 语义分析
        lines.append("<p></p>")
        lines.append("<p><b>语义分析:</b></p>")
        wx_exec = summary.get("首通语义点执行_邀请添加企微/WS/Line执行率")
        if wx_exec is not None:
            lines.append(f"<p><b>--添加企微/WS/Line的执行率为{_fmt_pct(wx_exec)}</b></p>")
        yj_exec = summary.get("首通语义点执行_一家多娃问询执行率")
        zjs_exec = summary.get("首通语义点执行_转介绍执行率")
        if yj_exec is not None and zjs_exec is not None:
            lines.append(f"<p><b>--一家多娃问询&amp;转介绍执行率为{_fmt_pct(yj_exec)}/{_fmt_pct(zjs_exec)}</b></p>")

    lines.append('</callout>')
    return "".join(lines)


if __name__ == "__main__":
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument("--merged", required=True)
    args = parser.parse_args()

    callout = generate_4_1_callout(Path(args.merged))
    print(callout)
