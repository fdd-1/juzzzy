"""根据参考文档生成各板块的结论 callout

结论模板来自：https://my.feishu.cn/docx/Veyzd0uGtoBKvyxRGgIcp3MznQc
"""
from __future__ import annotations
import sys
from pathlib import Path
import pandas as pd
import warnings
warnings.filterwarnings('ignore')

if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    except (AttributeError, ValueError):
        pass


def parse_pct(val):
    """解析百分比字符串为浮点数 (0-1)"""
    if pd.isna(val):
        return None
    s = str(val).strip()
    if s.endswith('%'):
        try:
            return float(s.rstrip('%')) / 100
        except (ValueError, TypeError):
            return None
    try:
        return float(s)
    except (ValueError, TypeError):
        return None


def fmt_pct(val) -> str:
    """格式化为百分比字符串"""
    v = parse_pct(val)
    if v is None:
        return "—"
    return f"{v * 100:.2f}%"


def gen_4_2_callout(df: pd.DataFrame) -> str:
    """4.2 组班意向结论"""
    # 找到海外团队总计行
    if '团队/小组' not in df.columns:
        return '<callout emoji="❗"><p>组班多意向占比数据待补充</p></callout>'

    overall = df[df['团队/小组'].astype(str).str.contains('海外', na=False)]
    if len(overall) == 0:
        return '<callout emoji="❗"><p>组班多意向占比数据待补充</p></callout>'

    row = overall.iloc[0]
    duoyixiang = row.get('汇总_多意向占比', '—')

    return f'''<callout emoji="❗">
<p><b>整体多意向占比{fmt_pct(duoyixiang)}</b></p>
</callout>'''


def gen_4_3_callout(df: pd.DataFrame) -> str:
    """4.3 群发消息结论"""
    # 找到海外团队总计
    if '小组' not in df.columns:
        return '<callout emoji="❗"><p>群发消息数据待补充</p></callout>'

    overall = df[df['小组'].astype(str).str.contains('海外', na=False)]
    if len(overall) == 0:
        # 尝试找'总计'行
        overall = df[df.iloc[:, 1].astype(str) == '总计']
    if len(overall) == 0:
        return '<callout emoji="❗"><p>群发消息数据待补充</p></callout>'

    row = overall.iloc[0]
    # 找营销类群发消息数列
    msg_count = "—"
    for col in df.columns:
        if '群发消息数' in col and '营销' in str(col):
            msg_count = row[col]
            break

    return f'''<callout emoji="❗">
<p><b>营销类lp群发消息总数{msg_count}条</b></p>
</callout>'''


def gen_4_4_callout(df: pd.DataFrame) -> str:
    """4.4 停课唤醒结论 - 含落后组提示"""
    if 'lp组别' not in df.columns:
        return '<callout emoji="❗"><p>停课唤醒数据待补充</p></callout>'

    # 找海外团队总计
    overall = df[df['lp组别'] == '海外团队']
    if len(overall) == 0:
        return '<callout emoji="❗"><p>停课唤醒数据待补充</p></callout>'

    row = overall.iloc[0]
    tk_zhanbi = row.get('停课占比', '—')

    # 找90天内唤醒率 - 列名"停课90天内唤醒_唤醒率"
    huanxing_rate = '—'
    waihu_col = None
    for col in df.columns:
        col_s = str(col)
        if '90天内唤醒' in col_s and '唤醒率' in col_s:
            huanxing_rate = row[col]
        if '本月停课唤醒目标学员' in col_s and '唤醒率' in col_s:
            waihu_col = col

    waihu_rate = row[waihu_col] if waihu_col else '—'

    # 找落后组：停课90天内唤醒-唤醒率 低的组
    laggards = []
    summary_df = df[df['LP个人'] == '总计']
    summary_df = summary_df[summary_df['lp组别'] != '海外团队']

    for _, r in summary_df.iterrows():
        team = r.get('lp组别', '')
        for col in df.columns:
            col_s = str(col)
            if '本月停课唤醒目标学员' in col_s and '唤醒率' in col_s:
                rate = parse_pct(r[col])
                if rate is not None and rate < 0.5:
                    laggards.append((team, fmt_pct(r[col])))
                    break

    callout_parts = [
        '<callout emoji="❗">',
        f'<p><b>当前整体停课占比{tk_zhanbi}</b></p>',
        f'<p><b>整体停课90天内唤醒率达{huanxing_rate}</b></p>',
        f'<p>整体停课唤醒跟进为{waihu_rate}</p>',
    ]
    for team, rate in laggards[:3]:
        callout_parts.append(f'<p>——{team}外呼跟进率仅{rate}，需注意提醒LP跟进停课唤醒目标学员</p>')
    callout_parts.append('</callout>')

    return '\n'.join(callout_parts)


def gen_4_5_fuwuyue_callout(df: pd.DataFrame) -> str:
    """4.5 服务月跟进结论"""
    if '小组' not in df.columns:
        return '<callout emoji="❗"><p>服务月跟进数据待补充</p></callout>'

    overall = df[df['小组'].astype(str).str.contains('海外', na=False)]
    if len(overall) == 0:
        return '<callout emoji="❗"><p>服务月跟进数据待补充</p></callout>'

    row = overall.iloc[0]
    waihu_rate = '—'
    follow_rate = '—'
    for col in df.columns:
        col_s = str(col)
        # 严格匹配：服务池-外呼跟进率（不含 微信、有效）
        if col_s == '服务池-外呼跟进率':
            waihu_rate = row[col]
        if col_s == '服务池-综合有效跟进率':
            follow_rate = row[col]

    # 找落后组：服务池-外呼跟进率 低
    laggards = []
    summary = df[df['LP'] == '总计']
    summary = summary[~summary['小组'].astype(str).str.contains('海外', na=False)]

    for _, r in summary.iterrows():
        team = r.get('小组', '')
        rate = parse_pct(r.get('服务池-外呼跟进率'))
        if rate is not None and rate < 0.5:
            laggards.append((team, fmt_pct(r['服务池-外呼跟进率'])))

    callout_parts = [
        '<callout emoji="❗">',
        f'<p><b>服务池跟进：</b></p>',
        f'<p>外呼覆盖率{waihu_rate}，综合有效跟进达{follow_rate}</p>',
    ]
    for team, rate in laggards[:3]:
        callout_parts.append(f'<p>——{team}服务池跟进仅{rate}，需注意</p>')
    callout_parts.append('</callout>')

    return '\n'.join(callout_parts)


def gen_4_5_sop_callout(df: pd.DataFrame) -> str:
    """4.5 服务池SOP结论"""
    if '小组' not in df.columns:
        return '<callout emoji="❗"><p>服务池SOP数据待补充</p></callout>'

    overall = df[df['小组'].astype(str).str.contains('海外', na=False)]
    if len(overall) == 0:
        return '<callout emoji="❗"><p>服务池SOP数据待补充</p></callout>'

    row = overall.iloc[0]
    jiahuo = '—'
    for col in df.columns:
        if '执行率加和' in str(col):
            jiahuo = row[col]
            break

    target = 2.4
    try:
        jiahuo_val = float(jiahuo) if not isinstance(jiahuo, str) else float(str(jiahuo).rstrip('%'))
        achieved = '达' if jiahuo_val >= target else '未达'
    except (ValueError, TypeError):
        achieved = '未达'

    # 找落后组
    laggards = []
    summary = df[df['LP'] == '总计']
    summary = summary[~summary['小组'].astype(str).str.contains('海外', na=False)]

    for _, r in summary.iterrows():
        team = r.get('小组', '')
        for col in df.columns:
            if '执行率加和' in str(col):
                try:
                    val = r[col]
                    if isinstance(val, str):
                        val = float(val.rstrip('%'))
                    if val < target:
                        laggards.append((team, val))
                except (ValueError, TypeError):
                    pass
                break

    callout_parts = [
        '<callout emoji="❗">',
        f'<p><b>语义分析：</b></p>',
        f'<p>服务池语义点加和执行为{jiahuo}，{achieved}服务池语义点加和目标（2.4）</p>',
    ]
    for team, val in sorted(laggards, key=lambda x: x[1])[:2]:
        callout_parts.append(f'<p>——{team}语义点执行较低（{val:.2f}）</p>')
    callout_parts.append('</callout>')

    return '\n'.join(callout_parts)


def gen_4_6_waihu_callout(df: pd.DataFrame) -> str:
    """4.6 系统外呼结论"""
    if '小组' not in df.columns:
        return '<callout emoji="❗"><p>系统外呼数据待补充</p></callout>'

    # 找海外团队总计 / 整体
    overall = df[df['小组'].astype(str).str.contains('海外', na=False)]
    if len(overall) == 0:
        overall = df.iloc[[0]]

    row = overall.iloc[0]
    coverage = '—'
    coverage_mom = '—'
    avg_call = '—'
    avg_call_mom = '—'

    for col in df.columns:
        col_s = str(col)
        if '覆盖率' in col_s and '月环比' not in col_s and '整体' in col_s:
            coverage = row[col]
        if '整体' in col_s and '覆盖率月环比' in col_s:
            coverage_mom = row[col]
        if '生均呼次' in col_s and '月环比' not in col_s:
            avg_call = row[col]
        if '生均外呼次数月环比' in col_s:
            avg_call_mom = row[col]

    return f'''<callout emoji="❗">
<p><b>整体系统外呼：</b></p>
<p>全部活跃学员的外呼覆盖率为{coverage}月环比{coverage_mom}，生均呼次{avg_call}月环比{avg_call_mom}</p>
</callout>'''


def gen_4_6_qiwei_callout(df: pd.DataFrame) -> str:
    """4.6 企微回复结论"""
    overall = None
    if '当前小组' in df.columns:
        overall = df[df['当前小组'].astype(str).str.contains('海外', na=False)]
    elif '小组' in df.columns:
        overall = df[df['小组'].astype(str).str.contains('海外', na=False)]

    if overall is None or len(overall) == 0:
        # 尝试找第一行（如果是总计）
        if '当前小组' in df.columns:
            overall = df[df['当前小组'].astype(str) == '总计']
            if len(overall) == 0:
                overall = df.iloc[[0]] if len(df) > 0 else None

    if overall is None or len(overall) == 0:
        return '<callout emoji="❗"><p>整体微信发送&回复比数据待补充</p></callout>'

    row = overall.iloc[0]
    huifu_bi = '—'
    for col in df.columns:
        col_s = str(col)
        if '回复比' in col_s and '整体' in col_s:
            huifu_bi = row[col]
            break

    return f'''<callout emoji="❗">
<p><b>整体微信发送&amp;回复比：</b></p>
<p>整体群发回复比为{huifu_bi}</p>
</callout>'''


def generate_all_callouts(base_dir: Path) -> dict:
    """生成所有板块的结论 callout"""
    callouts = {}

    # 4.2
    df42 = pd.read_excel(base_dir / "4_2" / "_merged_4_2_v2.xlsx")
    callouts['4.2'] = gen_4_2_callout(df42)

    # 4.3
    df43 = pd.read_excel(base_dir / "4_3" / "_merged_4_3_v2.xlsx")
    callouts['4.3'] = gen_4_3_callout(df43)

    # 4.4
    df44 = pd.read_excel(base_dir / "4_4" / "_merged_4_4_v3.xlsx")
    callouts['4.4'] = gen_4_4_callout(df44)

    # 4.5 服务月跟进
    df45f = pd.read_excel(base_dir / "4_5" / "_merged_4_5_fuwuyue_v2.xlsx")
    callouts['4.5_fuwuyue'] = gen_4_5_fuwuyue_callout(df45f)

    # 4.5 服务池SOP
    df45s = pd.read_excel(base_dir / "4_5" / "_merged_4_5_sop_v2.xlsx")
    callouts['4.5_sop'] = gen_4_5_sop_callout(df45s)

    # 4.6 外呼
    df46w = pd.read_excel(base_dir / "4_6" / "_merged_4_6_waihu_v2.xlsx")
    callouts['4.6_waihu'] = gen_4_6_waihu_callout(df46w)

    # 4.6 企微
    df46q = pd.read_excel(base_dir / "4_6" / "_merged_4_6_qiwei_v2.xlsx")
    callouts['4.6_qiwei'] = gen_4_6_qiwei_callout(df46q)

    return callouts


if __name__ == "__main__":
    import argparse
    import sys as _sys

    _sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from _paths import PROJECT_ROOT  # noqa: E402

    parser = argparse.ArgumentParser(description="生成所有板块的 callout")
    parser.add_argument(
        "--base",
        default=str(PROJECT_ROOT / "exports" / "weekly_20260601_20260607"),
        help="输入根目录（包含 4_X / _merged_4_X.xlsx）",
    )
    args = parser.parse_args()

    callouts = generate_all_callouts(Path(args.base))
    for key, callout in callouts.items():
        print(f"\n=== {key} ===")
        print(callout)
