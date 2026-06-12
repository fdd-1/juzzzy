"""按参考文档格式生成各板块结论 callout

参考: https://my.feishu.cn/docx/SbLFdUogiouIErx0zpXcw4krnij

格式:
> **小标题** · 整体描述

**亮点**
- 点1
- 点2

**风险**
- 点1
- 点2

**待办**
- 行动1
- 行动2
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
    v = parse_pct(val)
    if v is None:
        return "—"
    return f"{v * 100:.2f}%"


def fmt_num(val, decimals: int = 2) -> str:
    if pd.isna(val):
        return "—"
    try:
        if isinstance(val, str) and val.endswith('%'):
            return val
        return f"{float(val):.{decimals}f}"
    except (ValueError, TypeError):
        return str(val)


def make_callout(emoji: str, summary: str, highlights: list, risks: list, todos: list) -> str:
    """构造参考文档格式的 callout XML"""
    parts = [f'<callout emoji="{emoji}">']

    # 总结
    parts.append(f'<p><b>整体：</b>{summary}</p>')
    parts.append('<p></p>')

    # 亮点
    if highlights:
        parts.append('<p><b>亮点</b></p>')
        for h in highlights:
            parts.append(f'<p>·{h}</p>')
        parts.append('<p></p>')

    # 风险
    if risks:
        parts.append('<p><b>风险</b></p>')
        for r in risks:
            parts.append(f'<p>·{r}</p>')
        parts.append('<p></p>')

    # 待办
    if todos:
        parts.append('<p><b>待办</b></p>')
        for t in todos:
            parts.append(f'<p>·{t}</p>')

    parts.append('</callout>')
    return '\n'.join(parts)


def gen_4_1_callout(df: pd.DataFrame) -> str:
    """4.1 服务指标跟进 & 语义分析 - 完整结论"""

    # 找海外团队总计
    overall = df[(df['团队'] == '海外团队') & (df['LP'] == '总计')]
    if len(overall) == 0:
        return '<callout emoji="❗"><p>4.1 数据待补充</p></callout>'

    row = overall.iloc[0]

    # 各指标
    ft_follow = parse_pct(row.get('首通_跟进率'))
    ft_timely = parse_pct(row.get('首通_及时跟进率'))
    ft_qiwei = parse_pct(row.get('首通_企微绑定率'))
    ft_miaogua = parse_pct(row.get('首通_秒挂占比'))
    sk_follow = parse_pct(row.get('首课跟进率'))
    sk_timely = parse_pct(row.get('首课及时跟进率'))
    sz_follow = parse_pct(row.get('首专跟进率'))
    sz_timely = parse_pct(row.get('首专及时跟进率'))
    qiwei_exec = parse_pct(row.get('首通语义点执行_邀请添加企微/WS/Line执行率'))
    yjdw_exec = parse_pct(row.get('首通语义点执行_一家多娃问询执行率'))
    zjs_exec = parse_pct(row.get('首通语义点执行_转介绍执行率'))

    # 找小组总计
    teams = df[(df['LP'] == '总计') & (df['团队'] != '海外团队')]

    # 落后组：首通及时跟进率 < 95%
    ft_lag = []
    for _, r in teams.iterrows():
        rate = parse_pct(r.get('首通_及时跟进率'))
        if rate is not None and rate < 0.95 and r['团队'] != '台湾组':
            ft_lag.append((r['团队'], rate))
    ft_lag.sort(key=lambda x: x[1])

    # 首课落后：< 85%
    sk_lag = []
    for _, r in teams.iterrows():
        rate = parse_pct(r.get('首课及时跟进率'))
        if rate is not None and rate < 0.85 and r['团队'] != '台湾组':
            sk_lag.append((r['团队'], rate))
    sk_lag.sort(key=lambda x: x[1])

    # 首专落后：< 85%
    sz_lag = []
    for _, r in teams.iterrows():
        rate = parse_pct(r.get('首专及时跟进率'))
        if rate is not None and rate < 0.85 and r['团队'] != '台湾组':
            sz_lag.append((r['团队'], rate))
    sz_lag.sort(key=lambda x: x[1])

    # 落后LP：首通及时跟进率 < 80% 且新生数 >= 3
    ft_lag_lp = []
    individuals = df[df['LP'] != '总计']
    for _, r in individuals.iterrows():
        new_count = r.get('首通_新生数')
        try:
            n = float(new_count) if not pd.isna(new_count) else 0
        except (ValueError, TypeError):
            n = 0
        rate = parse_pct(r.get('首通_及时跟进率'))
        if rate is not None and rate < 0.8 and n >= 3:
            ft_lag_lp.append((r['团队'], r['LP'], rate))
    ft_lag_lp.sort(key=lambda x: x[2])

    # 首课落后LP
    sk_lag_lp = []
    for _, r in individuals.iterrows():
        n_val = r.get('首课学员数')
        try:
            n = float(n_val) if not pd.isna(n_val) else 0
        except (ValueError, TypeError):
            n = 0
        rate = parse_pct(r.get('首课及时跟进率'))
        if rate is not None and rate < 0.7 and n >= 3:
            sk_lag_lp.append((r['团队'], r['LP'], rate))
    sk_lag_lp.sort(key=lambda x: x[2])

    summary = (
        f'首通跟进率{fmt_pct(ft_follow)}/及时{fmt_pct(ft_timely)}（目标95%），'
        f'首课跟进率{fmt_pct(sk_follow)}/及时{fmt_pct(sk_timely)}（目标85%），'
        f'首专跟进率{fmt_pct(sz_follow)}/及时{fmt_pct(sz_timely)}（目标85%）。'
        f'企微绑定率{fmt_pct(ft_qiwei)}，秒挂占比{fmt_pct(ft_miaogua)}。'
    )

    # 亮点
    highlights = []
    achieve_groups = []
    for _, r in teams.iterrows():
        if r['团队'] == '台湾组':
            continue
        rate = parse_pct(r.get('首通_及时跟进率'))
        if rate is not None and rate >= 0.95:
            achieve_groups.append((r['团队'], rate))
    achieve_groups.sort(key=lambda x: -x[1])
    if achieve_groups:
        top = achieve_groups[:2]
        highlights.append(f'首通及时跟进率达标的小组：' + '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in top))

    sk_top = [(r['团队'], parse_pct(r.get('首课及时跟进率'))) for _, r in teams.iterrows()
              if r['团队'] != '台湾组' and parse_pct(r.get('首课及时跟进率')) is not None]
    sk_top = [t for t in sk_top if t[1] >= 0.85]
    sk_top.sort(key=lambda x: -x[1])
    if sk_top[:2]:
        highlights.append(f'首课及时跟进率领先：' + '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in sk_top[:2]))

    if qiwei_exec is not None and qiwei_exec >= 0.5:
        highlights.append(f'添加企微执行率{fmt_pct(qiwei_exec)}，整体执行较好')

    # 风险
    risks = []
    if ft_timely is not None and ft_timely < 0.95:
        risks.append(f'首通及时跟进率{fmt_pct(ft_timely)}，未达95%目标')
    if ft_lag:
        groups_str = '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in ft_lag[:3])
        risks.append(f'首通及时跟进率落后小组：{groups_str}')
    if sk_lag:
        groups_str = '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in sk_lag[:3])
        risks.append(f'首课及时跟进率落后小组：{groups_str}（目标85%）')
    if sz_lag:
        groups_str = '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in sz_lag[:3])
        risks.append(f'首专及时跟进率落后小组：{groups_str}（目标85%）')
    if ft_lag_lp:
        lps_str = '、'.join(f'{lp[0]}-{lp[1]}({fmt_pct(lp[2])})' for lp in ft_lag_lp[:3])
        risks.append(f'首通及时跟进率落后LP：{lps_str}')
    if ft_qiwei is not None and ft_qiwei < 0.6:
        risks.append(f'企微绑定率{fmt_pct(ft_qiwei)}，整体偏低')
    if qiwei_exec is not None and qiwei_exec < 0.5:
        risks.append(f'添加企微执行率仅{fmt_pct(qiwei_exec)}，需提升')

    # 待办
    todos = []
    if ft_lag:
        todos.append(f'落后组首通及时跟进率本周提升至95%以上：{ft_lag[0][0]}/{ft_lag[1][0] if len(ft_lag)>1 else ""}'.rstrip('/'))
    if ft_lag_lp:
        first_lp = ft_lag_lp[0]
        todos.append(f'TL本周抽查落后LP外呼跟进数据，重点关注：{first_lp[0]}-{first_lp[1]}')
    if sk_lag_lp:
        first_lp = sk_lag_lp[0]
        todos.append(f'每日排查首课未跟进数据，重点关注：{first_lp[0]}-{first_lp[1]}')
    if qiwei_exec is not None and qiwei_exec < 0.5:
        todos.append('排查首通语义点执行情况，提升邀请添加企微执行率')
    if not todos:
        todos.append('保持当前跟进节奏，TL本周抽查录音质量')

    return make_callout('❗', summary, highlights, risks, todos)


def gen_4_2_callout(df: pd.DataFrame) -> str:
    """4.2 组班多意向占比"""
    overall = df[df['团队/小组'].astype(str).str.contains('海外', na=False)]
    if len(overall) == 0:
        return '<callout emoji="❗"><p>4.2 数据待补充</p></callout>'

    row = overall.iloc[0]
    duoyixiang = row.get('汇总_多意向占比')
    waiting = row.get('汇总_当前意向等待学员数')

    summary = f'整体多意向占比{fmt_pct(duoyixiang)}，当前意向等待学员{fmt_num(waiting, 0)}人。'

    # 找各组多意向占比
    teams = df[df['LP'] == '总计']
    teams = teams[~teams['团队/小组'].astype(str).str.contains('海外', na=False)]

    high_groups = []
    low_groups = []
    overall_v = parse_pct(duoyixiang)

    for _, r in teams.iterrows():
        v = parse_pct(r.get('汇总_多意向占比'))
        if v is None:
            continue
        team = r['团队/小组']
        if overall_v and v >= overall_v:
            high_groups.append((team, v))
        elif overall_v and v < overall_v * 0.7:  # 低于整体 70%
            low_groups.append((team, v))

    high_groups.sort(key=lambda x: -x[1])
    low_groups.sort(key=lambda x: x[1])

    highlights = []
    if high_groups:
        groups_str = '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in high_groups[:3])
        highlights.append(f'多意向占比高于整体的小组：{groups_str}')

    risks = []
    if low_groups:
        groups_str = '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in low_groups[:3])
        risks.append(f'多意向占比明显落后的小组：{groups_str}')

    todos = ['关注高意向学员转化进度，盘点多意向学员的报班节奏']
    if low_groups:
        todos.append(f'落后组复盘组班意向跟进话术：{low_groups[0][0]}')

    return make_callout('❗', summary, highlights, risks, todos)


def gen_4_3_callout(df: pd.DataFrame) -> str:
    """4.3 群发跟进"""
    if '小组' not in df.columns:
        return '<callout emoji="❗"><p>4.3 数据待补充</p></callout>'

    overall = df[df['小组'].astype(str).str.contains('海外', na=False)]
    if len(overall) == 0:
        # 尝试找总计行
        if '负责人/LP姓名' in df.columns:
            overall = df[df['负责人/LP姓名'] == '总计'].head(1)

    if len(overall) == 0:
        return '<callout emoji="❗"><p>4.3 数据待补充</p></callout>'

    row = overall.iloc[0]

    # 找营销类总数和服务类总数
    marketing_count = '—'
    service_count = '—'
    grfa_zhanbi = '—'
    for col in df.columns:
        col_s = str(col)
        if '营销' in col_s and '群发消息数' in col_s:
            marketing_count = row[col]
            break
    for col in df.columns:
        col_s = str(col)
        if '服务' in col_s and '群发消息数' in col_s:
            service_count = row[col]
            break
    for col in df.columns:
        col_s = str(col)
        if '个人群发占比' in col_s:
            grfa_zhanbi = row[col]
            break

    summary = f'营销类LP群发{fmt_num(marketing_count, 0)}条，服务类LP群发{fmt_num(service_count, 0)}条，个人群发占比{grfa_zhanbi}。'

    # 找各组个人群发占比
    teams = df[df['负责人/LP姓名'] == '总计'] if '负责人/LP姓名' in df.columns else pd.DataFrame()
    if len(teams) > 0:
        teams = teams[~teams['小组'].astype(str).str.contains('海外', na=False)]

    high_groups = []
    low_groups = []
    overall_grfa = parse_pct(grfa_zhanbi)

    for _, r in teams.iterrows():
        v = parse_pct(r.get('营销类LP群发-汇总_个人群发占比') if '营销类LP群发-汇总_个人群发占比' in df.columns else None)
        if v is None:
            continue
        team = r['小组']
        if overall_grfa and v >= overall_grfa:
            high_groups.append((team, v))
        elif overall_grfa and v < overall_grfa * 0.5:
            low_groups.append((team, v))

    highlights = []
    if high_groups:
        high_groups.sort(key=lambda x: -x[1])
        highlights.append(f'个人群发占比表现较好：' + '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in high_groups[:2]))

    risks = []
    if low_groups:
        low_groups.sort(key=lambda x: x[1])
        risks.append(f'个人群发不足的小组：' + '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in low_groups[:2]))

    todos = ['持续关注营销类群发与服务类群发的覆盖与质量']
    if low_groups:
        todos.append(f'落后组提升LP个人群发频次：{low_groups[0][0]}')

    return make_callout('❗', summary, highlights, risks, todos)


def gen_4_4_callout(df: pd.DataFrame) -> str:
    """4.4 停课唤醒"""
    overall = df[(df['lp组别'] == '海外团队') & (df['LP个人'] == '总计')]
    if len(overall) == 0:
        return '<callout emoji="❗"><p>4.4 数据待补充</p></callout>'

    row = overall.iloc[0]
    tk_zhanbi = row.get('停课占比')

    huanxing_rate = '—'
    waihu_rate = '—'
    for col in df.columns:
        col_s = str(col)
        if '90天内唤醒' in col_s and '唤醒率' in col_s:
            huanxing_rate = row[col]
        if '本月停课唤醒目标学员' in col_s and '唤醒率' in col_s:
            waihu_rate = row[col]

    summary = f'停课占比{tk_zhanbi}（目标6%），停课90天内唤醒率{huanxing_rate}，整体唤醒跟进{waihu_rate}。'

    # 各组数据
    teams = df[(df['LP个人'] == '总计') & (df['lp组别'] != '海外团队')]

    high_groups = []
    low_groups = []
    overall_huanxing = parse_pct(huanxing_rate)

    for _, r in teams.iterrows():
        v = None
        for col in df.columns:
            if '90天内唤醒' in str(col) and '唤醒率' in str(col):
                v = parse_pct(r[col])
                break
        if v is None:
            continue
        team = r['lp组别']
        if overall_huanxing and v >= overall_huanxing * 1.2:
            high_groups.append((team, v))
        elif v < (overall_huanxing or 0) * 0.5:
            low_groups.append((team, v))

    # 停课占比落后（>目标）
    target = 0.06
    over_target = []
    for _, r in teams.iterrows():
        v = parse_pct(r.get('停课占比'))
        if v is not None and v > target:
            over_target.append((r['lp组别'], v))
    over_target.sort(key=lambda x: -x[1])

    # 落后LP
    individuals = df[df['LP个人'] != '总计']
    lag_lps = []
    for _, r in individuals.iterrows():
        v = parse_pct(r.get('停课占比'))
        if v is not None and v > 0.1:  # 停课占比 > 10%
            lag_lps.append((r.get('lp组别'), r.get('LP个人'), v))
    lag_lps.sort(key=lambda x: -x[2])

    highlights = []
    if high_groups:
        high_groups.sort(key=lambda x: -x[1])
        highlights.append(f'90天内唤醒率领先：' + '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in high_groups[:2]))

    risks = []
    if over_target:
        groups_str = '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in over_target[:3])
        risks.append(f'停课占比超目标(6%)的小组：{groups_str}')
    if low_groups:
        low_groups.sort(key=lambda x: x[1])
        groups_str = '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in low_groups[:3])
        risks.append(f'90天内唤醒率落后小组：{groups_str}')
    if lag_lps:
        lps_str = '、'.join(f'{lp[0]}-{lp[1]}({fmt_pct(lp[2])})' for lp in lag_lps[:3])
        risks.append(f'停课占比偏高的LP：{lps_str}')

    todos = []
    if over_target:
        todos.append(f'停课占比超目标的小组本周制定停课唤醒计划：{over_target[0][0]}')
    if lag_lps:
        first = lag_lps[0]
        todos.append(f'TL重点关注高停课占比LP：{first[0]}-{first[1]}')
    if low_groups:
        todos.append(f'唤醒率落后组提升外呼跟进密度：{low_groups[0][0]}')
    if not todos:
        todos.append('维持现有唤醒跟进节奏')

    return make_callout('❗', summary, highlights, risks, todos)


def gen_4_5_fuwuyue_callout(df: pd.DataFrame) -> str:
    """4.5 服务月跟进"""
    overall = df[(df['小组'] == '海外团队') & (df['LP'] == '总计')]
    if len(overall) == 0:
        return '<callout emoji="❗"><p>4.5 数据待补充</p></callout>'

    row = overall.iloc[0]
    waihu = row.get('服务池-外呼跟进率')
    follow = row.get('服务池-综合有效跟进率')

    summary = f'服务池外呼跟进率{waihu}，综合有效跟进率{follow}。'

    # 各组数据
    teams = df[(df['LP'] == '总计') & (df['小组'] != '海外团队')]

    high_groups = []
    low_groups = []
    overall_waihu = parse_pct(waihu)

    for _, r in teams.iterrows():
        v = parse_pct(r.get('服务池-外呼跟进率'))
        if v is None:
            continue
        team = r['小组']
        if overall_waihu and v >= overall_waihu * 1.1:
            high_groups.append((team, v))
        elif overall_waihu and v < overall_waihu * 0.7:
            low_groups.append((team, v))

    # 落后LP：服务池-外呼跟进率 < 30% 且 学员数 >= 5
    individuals = df[df['LP'] != '总计']
    lag_lps = []
    for _, r in individuals.iterrows():
        v = parse_pct(r.get('服务池-外呼跟进率'))
        try:
            n = float(r.get('服务池-学员数', 0))
        except (ValueError, TypeError):
            n = 0
        if v is not None and v < 0.3 and n >= 5:
            lag_lps.append((r['小组'], r['LP'], v))
    lag_lps.sort(key=lambda x: x[2])

    highlights = []
    if high_groups:
        high_groups.sort(key=lambda x: -x[1])
        highlights.append(f'服务池外呼跟进率领先：' + '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in high_groups[:2]))

    risks = []
    if low_groups:
        low_groups.sort(key=lambda x: x[1])
        groups_str = '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in low_groups[:3])
        risks.append(f'服务池外呼跟进率落后小组：{groups_str}')
    if lag_lps:
        lps_str = '、'.join(f'{lp[0]}-{lp[1]}({fmt_pct(lp[2])})' for lp in lag_lps[:3])
        risks.append(f'服务池跟进偏弱的LP：{lps_str}')

    todos = []
    if low_groups:
        todos.append(f'落后组本周外呼跟进率提升至整体均值以上：{low_groups[0][0]}')
    if lag_lps:
        first = lag_lps[0]
        todos.append(f'TL重点抓服务池跟进数据：{first[0]}-{first[1]}')
    todos.append('TL抽查服务池外呼录音，确认是否有效触达学员')

    return make_callout('❗', summary, highlights, risks, todos)


def gen_4_5_sop_callout(df: pd.DataFrame) -> str:
    """4.5 服务池SOP"""
    overall = df[(df['小组'] == '海外团队') & (df['LP'] == '总计')]
    if len(overall) == 0:
        return '<callout emoji="❗"><p>4.5 SOP 数据待补充</p></callout>'

    row = overall.iloc[0]
    jiahuo = '—'
    for col in df.columns:
        if '执行率加和' in str(col):
            jiahuo = row[col]
            break

    target = 2.4
    try:
        jh_val = float(jiahuo) if not isinstance(jiahuo, str) else float(str(jiahuo).rstrip('%'))
    except (ValueError, TypeError):
        jh_val = 0

    summary = f'服务池语义点执行率加和{fmt_num(jh_val, 2)}（目标2.4），{"已达目标" if jh_val >= target else "未达目标"}。'

    teams = df[(df['LP'] == '总计') & (df['小组'] != '海外团队')]

    high_groups = []
    low_groups = []

    for _, r in teams.iterrows():
        v = None
        for col in df.columns:
            if '执行率加和' in str(col):
                try:
                    val = r[col]
                    if isinstance(val, str):
                        v = float(val.rstrip('%'))
                    else:
                        v = float(val)
                except (ValueError, TypeError):
                    v = None
                break
        if v is None:
            continue
        team = r['小组']
        if v >= target:
            high_groups.append((team, v))
        elif v < target * 0.6:
            low_groups.append((team, v))

    # 落后LP
    individuals = df[df['LP'] != '总计']
    lag_lps = []
    for _, r in individuals.iterrows():
        v = None
        for col in df.columns:
            if '执行率加和' in str(col):
                try:
                    val = r[col]
                    if isinstance(val, str):
                        v = float(val.rstrip('%'))
                    else:
                        v = float(val)
                except (ValueError, TypeError):
                    v = None
                break
        try:
            n = float(r.get('服务池-命中服务池学员数', 0))
        except (ValueError, TypeError):
            n = 0
        if v is not None and v < 1.5 and n >= 5:
            lag_lps.append((r['小组'], r['LP'], v))
    lag_lps.sort(key=lambda x: x[2])

    highlights = []
    if high_groups:
        high_groups.sort(key=lambda x: -x[1])
        highlights.append(f'语义点加和达标的小组：' + '、'.join(f'{g[0]}({fmt_num(g[1], 2)})' for g in high_groups[:2]))

    risks = []
    if low_groups:
        low_groups.sort(key=lambda x: x[1])
        groups_str = '、'.join(f'{g[0]}({fmt_num(g[1], 2)})' for g in low_groups[:3])
        risks.append(f'语义点加和明显偏低的小组：{groups_str}')
    if lag_lps:
        lps_str = '、'.join(f'{lp[0]}-{lp[1]}({fmt_num(lp[2], 2)})' for lp in lag_lps[:3])
        risks.append(f'语义点执行偏弱的LP：{lps_str}')

    todos = []
    if low_groups:
        todos.append(f'落后组本周提升语义点执行率加和至2.4以上：{low_groups[0][0]}')
    if lag_lps:
        first = lag_lps[0]
        todos.append(f'TL抽查录音，重点辅导：{first[0]}-{first[1]}')
    todos.append('每周复盘服务池SOP执行情况，输出录音反馈')

    return make_callout('❗', summary, highlights, risks, todos)


def gen_4_6_waihu_callout(df: pd.DataFrame) -> str:
    """4.6 系统外呼监控"""
    # 找"整体"行
    overall = df[df['小组'].astype(str) == '整体']
    if len(overall) == 0:
        overall = df[df['小组'].astype(str).str.contains('海外', na=False)]
    if len(overall) == 0:
        return '<callout emoji="❗"><p>4.6 外呼数据待补充</p></callout>'

    row = overall.iloc[0]
    coverage = '—'
    coverage_mom = '—'
    avg_call = '—'
    avg_call_mom = '—'

    for col in df.columns:
        col_s = str(col)
        if col_s == '整体_覆盖率':
            coverage = row[col]
        if '整体_覆盖率月环比' in col_s:
            coverage_mom = row[col]
        if col_s == '整体_生均呼次':
            avg_call = row[col]
        if '生均外呼次数月环比' in col_s:
            avg_call_mom = row[col]

    summary = f'活跃学员外呼覆盖率{coverage}（月环比{coverage_mom}），生均呼次{avg_call}（月环比{avg_call_mom}）。'

    # 找各组覆盖率（LP姓名='总计' 且 小组 != '整体'）
    teams = df[(df['LP姓名'] == '总计') & (df['小组'] != '整体')]

    high_groups = []
    low_groups = []
    overall_cov = parse_pct(coverage)

    for _, r in teams.iterrows():
        v = parse_pct(r.get('整体_覆盖率'))
        if v is None:
            continue
        team = r['小组']
        if overall_cov and v >= overall_cov * 1.1:
            high_groups.append((team, v))
        elif overall_cov and v < overall_cov * 0.7:
            low_groups.append((team, v))

    # 落后LP
    individuals = df[df['LP姓名'] != '总计']
    lag_lps = []
    for _, r in individuals.iterrows():
        v = parse_pct(r.get('整体_覆盖率'))
        try:
            n = float(r.get('整体_学员量', 0))
        except (ValueError, TypeError):
            n = 0
        if v is not None and v < 0.2 and n >= 50:
            lag_lps.append((r.get('小组'), r['LP姓名'], v))
    lag_lps.sort(key=lambda x: x[2])

    highlights = []
    if high_groups:
        high_groups.sort(key=lambda x: -x[1])
        highlights.append(f'外呼覆盖率领先小组：' + '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in high_groups[:2]))

    risks = []
    if low_groups:
        low_groups.sort(key=lambda x: x[1])
        groups_str = '、'.join(f'{g[0]}({fmt_pct(g[1])})' for g in low_groups[:3])
        risks.append(f'外呼覆盖率落后小组：{groups_str}')
    if lag_lps:
        lps_str = '、'.join(f'{lp[0]}-{lp[1]}({fmt_pct(lp[2])})' for lp in lag_lps[:3])
        risks.append(f'外呼覆盖率偏低的LP：{lps_str}')

    todos = ['关注覆盖率月环比变化，确保跟进节奏稳定']
    if low_groups:
        todos.append(f'落后组本周提升外呼覆盖率：{low_groups[0][0]}')
    if lag_lps:
        first = lag_lps[0]
        todos.append(f'TL本周关注LP外呼数据：{first[0]}-{first[1]}')

    return make_callout('❗', summary, highlights, risks, todos)


def gen_4_6_qiwei_callout(df: pd.DataFrame) -> str:
    """4.6 企微回复"""
    team_col = '当前小组' if '当前小组' in df.columns else '小组'

    # 整体行 - 找 当前小组 = '整体' 或 含'海外'
    overall = df[df[team_col].astype(str) == '整体']
    if len(overall) == 0:
        overall = df[df[team_col].astype(str).str.contains('海外', na=False)]
    if len(overall) == 0:
        return '<callout emoji="❗"><p>4.6 企微数据待补充</p></callout>'

    row = overall.iloc[0]
    huifu_bi = '—'
    fasong_count = '—'
    for col in df.columns:
        col_s = str(col)
        if col_s == '整体_回复比':
            huifu_bi = row[col]
        if col_s == '整体_发送消息条数':
            fasong_count = row[col]

    summary = f'整体微信发送{fmt_num(fasong_count, 0)}条，群发回复比{fmt_num(huifu_bi, 2)}。'

    # 找各组回复比
    teams = df[(df['LP姓名'] == '总计') & (df[team_col].astype(str) != '整体')]

    try:
        overall_huifu = float(huifu_bi) if not isinstance(huifu_bi, str) else float(str(huifu_bi))
    except (ValueError, TypeError):
        overall_huifu = None

    high_groups = []
    low_groups = []

    for _, r in teams.iterrows():
        try:
            v = float(r.get('整体_回复比', 0))
        except (ValueError, TypeError):
            continue
        team = r[team_col]
        if overall_huifu and v >= overall_huifu * 1.1:
            high_groups.append((team, v))
        elif overall_huifu and v < overall_huifu * 0.7:
            low_groups.append((team, v))

    highlights = []
    if high_groups:
        high_groups.sort(key=lambda x: -x[1])
        highlights.append(f'回复比领先小组：' + '、'.join(f'{g[0]}({fmt_num(g[1], 2)})' for g in high_groups[:2]))

    risks = []
    if low_groups:
        low_groups.sort(key=lambda x: x[1])
        risks.append(f'回复比落后小组：' + '、'.join(f'{g[0]}({fmt_num(g[1], 2)})' for g in low_groups[:2]))

    todos = ['关注LP个人群发回复质量，提升回复比']
    if low_groups:
        todos.append(f'落后组本周复盘群发话术：{low_groups[0][0]}')

    return make_callout('❗', summary, highlights, risks, todos)


def generate_all_callouts(base_dir: Path) -> dict:
    """生成所有板块结论"""
    callouts = {}

    # 4.1
    df41 = pd.read_excel(base_dir / "4_1" / "_merged_4_1.xlsx")
    callouts['4.1'] = gen_4_1_callout(df41)

    # 4.2
    df42 = pd.read_excel(base_dir / "4_2" / "_merged_4_2_v2.xlsx")
    callouts['4.2'] = gen_4_2_callout(df42)

    # 4.3
    df43 = pd.read_excel(base_dir / "4_3" / "_merged_4_3_v2.xlsx")
    callouts['4.3'] = gen_4_3_callout(df43)

    # 4.4
    df44 = pd.read_excel(base_dir / "4_4" / "_merged_4_4_v3.xlsx")
    callouts['4.4'] = gen_4_4_callout(df44)

    # 4.5 fuwuyue
    df45f = pd.read_excel(base_dir / "4_5" / "_merged_4_5_fuwuyue_v2.xlsx")
    callouts['4.5_fuwuyue'] = gen_4_5_fuwuyue_callout(df45f)

    # 4.5 sop
    df45s = pd.read_excel(base_dir / "4_5" / "_merged_4_5_sop_v2.xlsx")
    callouts['4.5_sop'] = gen_4_5_sop_callout(df45s)

    # 4.6 waihu
    df46w = pd.read_excel(base_dir / "4_6" / "_merged_4_6_waihu_v2.xlsx")
    callouts['4.6_waihu'] = gen_4_6_waihu_callout(df46w)

    # 4.6 qiwei
    df46q = pd.read_excel(base_dir / "4_6" / "_merged_4_6_qiwei_v2.xlsx")
    callouts['4.6_qiwei'] = gen_4_6_qiwei_callout(df46q)

    return callouts


if __name__ == "__main__":
    import argparse
    import sys as _sys

    _sys.path.insert(0, str(Path(__file__).resolve().parent.parent))
    from _paths import PROJECT_ROOT  # noqa: E402

    parser = argparse.ArgumentParser(description="生成所有板块的 callout（v2）")
    parser.add_argument(
        "--base",
        default=str(PROJECT_ROOT / "exports" / "weekly_20260601_20260607"),
        help="输入根目录（包含 4_X / _merged_4_X*.xlsx）",
    )
    args = parser.parse_args()

    callouts = generate_all_callouts(Path(args.base))
    for key, c in callouts.items():
        print(f"\n=== {key} ===")
        print(c)
