# -*- coding: utf-8 -*-
"""Generate HTML sections 4.2-4.6 for the weekly report.

Output: writes new HTML to 4_2_to_4_6_section.html in the same directory.
The caller will splice this into the main report.
"""
import json
import os

DATA_DIR = os.path.dirname(os.path.abspath(__file__))


def load_values(fname):
    with open(os.path.join(DATA_DIR, fname), encoding='utf-8') as f:
        d = json.load(f)
    return d['data']['valueRange']['values']


def fmt_pct(v, dec=1):
    if v is None or v == '':
        return ''
    if isinstance(v, str):
        return v
    return f"{v * 100:.{dec}f}%"


def fmt_num(v, dec=0):
    if v is None or v == '':
        return ''
    if isinstance(v, str):
        return v
    if dec == 0:
        try:
            return str(int(round(v)))
        except Exception:
            return str(v)
    return f"{v:.{dec}f}"


def fmt_int(v):
    if v is None or v == '':
        return ''
    try:
        return str(int(round(float(v))))
    except Exception:
        return str(v)


def safe_get(row, idx):
    try:
        return row[idx]
    except IndexError:
        return None


# === 4.2 组班多意向占比 ===
def build_4_2():
    v = load_values('4_2_raw.json')
    # rows: 0 group header, 1 col header, 2 海外团队总计, 3 blank, 4-6 区块汇总, 7 blank, 8-14 小组汇总, 15 blank, 16+ 个人
    # Categories: 汇总(C-E), 复课M1-M3(F-K), 复课M4-M12+(L-Q), 调课M1-M3(R-W), 调课M4-M12+(X-AC)
    sub_cols = ['当前意向等待学员数', '2个意向及以上学员占比', '人均等待天数', '等待0-7天学员数', '等待8-14天学员数', '等待14天以上学员数']

    css_colors = {
        '汇总': '#FFF2CC',
        '复课M1-M3': '#BDD7EE',
        '复课M4-M12+': '#C6E0B4',
        '调课M1-M3': '#F4B084',
        '调课M4-M12+': '#FFD966',
    }

    out = []
    out.append('<h2>4.2 组班多意向占比</h2>')
    out.append('<div class="meta">数据周期：2026/05/01 — 2026/05/25（5月第三个自然周末日）｜ 报表：思维LP组班意向提交播报</div>')

    # ------- conclusion -------
    # Build conclusion from group rows (rows 8..14: 台湾组, 港澳1组, 港澳2组, 港澳组, 美澳1组, 美澳2组, 美澳3组, 美澳4组)
    # Actually looking at structure: rows 8-14 ARE per-group totals
    # 总体多意向占比 from row 2 (海外团队)
    overall = v[2]
    overall_total = overall[2]
    overall_count = overall[3]
    overall_ratio = overall[4]

    conc_lines = []
    conc_lines.append(f'<p><b>整体多意向占比 {fmt_pct(overall_ratio,2)}（当前意向等待学员 {overall_total} 人，其中 2 个意向及以上 {overall_count} 人）</b></p>')
    # Find groups with multi-intent ratio above/below
    group_rows = []
    # rows 4..6 are area-level totals (台湾区/港澳区/欧美澳区) — also include
    for i in range(4, 15):
        row = v[i]
        if row[0] and row[1] == '总计':
            group_rows.append(row)
    # Sort by 多意向占比 desc
    valid = [r for r in group_rows if r[4] is not None and r[2]]
    valid.sort(key=lambda r: r[4] or 0, reverse=True)
    high = [r for r in valid if (r[4] or 0) >= overall_ratio]
    low = [r for r in valid if (r[4] or 0) < overall_ratio]
    if high:
        ss = ' / '.join(f"{r[0]}({fmt_pct(r[4],2)})" for r in high[:6])
        conc_lines.append(f'<p>多意向占比高于整体的小组：<span class="hl-pink">{ss}</span></p>')
    if low:
        ss = ' / '.join(f"{r[0]}({fmt_pct(r[4],2)})" for r in low)
        conc_lines.append(f'<p>多意向占比低于整体的小组：<span class="hl-pink">{ss}</span></p>')
    # Categories breakdown for 海外团队
    conc_lines.append('<p>&nbsp;</p>')
    conc_lines.append('<p><b>分场景占比（海外团队整体）：</b></p>')
    cat_starts = [(5, '复课M1-M3'), (11, '复课M4-M12+'), (17, '调课M1-M3'), (23, '调课M4-M12+')]
    for c_idx, c_name in cat_starts:
        c_total = overall[c_idx]
        c_pct = overall[c_idx + 1]
        if c_total:
            conc_lines.append(f'<p>{c_name}：等待学员 {c_total} 人，2 个意向及以上 {fmt_pct(c_pct,2)}</p>')

    out.append('<h2 style="font-size:14px;border-left:4px solid #4472C4;">结论</h2>')
    out.append('<div class="callout"><div class="callout-emoji">❗</div><div class="callout-body">')
    out.extend(conc_lines)
    out.append('</div></div>')

    # ------- table -------
    out.append('<div class="table-wrap"><table class="report-table"><thead>')
    # group header row
    gh = '<tr class="group-h">'
    gh += '<th colspan="2" class="gh" style="background:#D9D9D9;color:#000">基础</th>'
    gh += f'<th colspan="3" class="gh" style="background:{css_colors["汇总"]};color:#000">汇总</th>'
    gh += f'<th colspan="6" class="gh" style="background:{css_colors["复课M1-M3"]};color:#000">复课：首消M1-M3</th>'
    gh += f'<th colspan="6" class="gh" style="background:{css_colors["复课M4-M12+"]};color:#000">复课：首消M4-M12+</th>'
    gh += f'<th colspan="6" class="gh" style="background:{css_colors["调课M1-M3"]};color:#000">调课：首消M1-M3</th>'
    gh += f'<th colspan="6" class="gh" style="background:{css_colors["调课M4-M12+"]};color:#000">调课：首消M4-M12+</th>'
    gh += '</tr>'
    out.append(gh)

    # column header row
    ch = '<tr class="col-row">'
    ch += '<th class="ch frozen" style="background:#D9D9D9">团队/小组</th>'
    ch += '<th class="ch frozen" style="background:#D9D9D9">LP</th>'
    # 汇总 (3 cols)
    ch += f'<th class="ch" style="background:{css_colors["汇总"]}">当前意向等待学员数</th>'
    ch += f'<th class="ch" style="background:{css_colors["汇总"]}">2个意向及以上学员数</th>'
    ch += f'<th class="ch" style="background:{css_colors["汇总"]}">多意向占比</th>'
    # 4 categories x 6 cols
    for cat_name in ['复课M1-M3', '复课M4-M12+', '调课M1-M3', '调课M4-M12+']:
        for col in sub_cols:
            ch += f'<th class="ch" style="background:{css_colors[cat_name]}">{col}</th>'
    ch += '</tr>'
    out.append(ch)
    out.append('</thead><tbody>')

    # body rows
    def render_row(row, kind='detail'):
        cls = 'summary' if kind == 'summary' else 'detail'
        cells = []
        cells.append(f'<td class="frozen">{row[0] or ""}</td>')
        cells.append(f'<td class="frozen">{row[1] or ""}</td>')
        # 汇总 cols
        cells.append(f'<td>{fmt_int(row[2])}</td>')
        cells.append(f'<td>{fmt_int(row[3])}</td>')
        cells.append(f'<td>{fmt_pct(row[4],2)}</td>')
        # 4 categories: F..K, L..Q, R..W, X..AC
        for start in [5, 11, 17, 23]:
            cells.append(f'<td>{fmt_int(row[start])}</td>')          # 当前意向等待学员数
            cells.append(f'<td>{fmt_pct(row[start+1],2)}</td>')      # 2个意向及以上学员占比
            cells.append(f'<td>{fmt_num(row[start+2],2)}</td>')      # 人均等待天数
            cells.append(f'<td>{fmt_int(row[start+3])}</td>')        # 0-7天
            cells.append(f'<td>{fmt_int(row[start+4])}</td>')        # 8-14天
            cells.append(f'<td>{fmt_int(row[start+5])}</td>')        # 14+
        return f'<tr class="{cls}">' + ''.join(cells) + '</tr>'

    # 海外团队 总计 (row 2)
    out.append(render_row(v[2], 'summary'))
    # 区块汇总 (rows 4-6)
    for i in [4, 5, 6]:
        if v[i][0]:
            out.append(render_row(v[i], 'summary'))
    # 小组汇总 (rows 8-14)
    for i in range(8, 15):
        if v[i][0]:
            out.append(render_row(v[i], 'summary'))
    # 个人 (rows 16+)
    for i in range(16, len(v)):
        if not v[i][1]:
            continue
        out.append(render_row(v[i], 'detail'))

    out.append('</tbody></table></div>')
    return '\n'.join(out)


# === 4.3 群发跟进 ===
def build_4_3():
    v = load_values('4_3_raw.json')
    out = []
    out.append('<h2>4.3 群发跟进</h2>')
    out.append('<div class="meta">数据周期：2026/05/01 — 2026/05/25（5月第三个自然周末日）｜ 报表：思维海外群发消息汇总数据播报</div>')

    overall = v[2]
    # cols (0-indexed): 0=小组, 1=负责人/LP姓名,
    # 营销类汇总: 2 群发消息数, 3 环比上月, 4 个人群发消息数, 5 环比上月, 6 个人群发占比, 7 群发学员数
    # 营销类重复内容: 8 续费, 9 转介绍, 10 拓美术, 11 拓叫叫, 12 拓英语
    # 服务类汇总: 13 群发消息数, 14 环比上月, 15 个人群发消息数, 16 环比上月, 17 个人群发占比, 18 群发学员数
    # 服务类重复内容: 19 通知-学员流转, 20 服务-学员流转, 21 服务-问卷调查, 22 服务-节日祝福, 23 服务-停课通知, 24 学习总结, 25 打卡
    # 26 其他LP群发-重复内容

    overall_market_total = overall[2]
    overall_market_lp = overall[4]
    overall_market_lp_pct = overall[6]
    overall_service_total = overall[13]
    overall_service_lp = overall[15]
    overall_service_lp_pct = overall[17]

    conc_lines = []
    conc_lines.append(f'<p><b>营销类 LP 群发：消息总数 {fmt_int(overall_market_total)} 条（环比上月 {fmt_pct(overall[3],2)}）；个人群发 {fmt_int(overall_market_lp)} 条（环比 {fmt_pct(overall[5],2)}），个人群发占比 {fmt_pct(overall_market_lp_pct,2)}；触达学员 {fmt_int(overall[7])} 人</b></p>')
    # marketing breakdown
    conc_lines.append(f'<p class="indent">└ 内容分布：续费 {fmt_int(overall[8])} 条 / 转介绍 {fmt_int(overall[9])} 条 / 拓美术 {fmt_int(overall[10])} 条 / 拓叫叫 {fmt_int(overall[11])} 条 / 拓英语 {fmt_int(overall[12])} 条</p>')
    conc_lines.append(f'<p><b>服务类 LP 群发：消息总数 {fmt_int(overall_service_total)} 条（环比 {fmt_pct(overall[14],2)}）；个人群发 {fmt_int(overall_service_lp)} 条（环比 {fmt_pct(overall[16],2)}），个人群发占比 {fmt_pct(overall_service_lp_pct,2)}；触达学员 {fmt_int(overall[18])} 人</b></p>')
    conc_lines.append(f'<p class="indent">└ 内容分布：通知-学员流转 {fmt_int(overall[19])} 条 / 服务-学员流转 {fmt_int(overall[20])} 条 / 问卷调查 {fmt_int(overall[21])} 条 / 节日祝福 {fmt_int(overall[22])} 条 / 停课通知 {fmt_int(overall[23])} 条 / 学习总结 {fmt_int(overall[24])} 条 / 打卡 {fmt_int(overall[25])} 条</p>')
    # find groups with high marketing message volume
    groups = []
    for i in range(3, len(v)):
        row = v[i]
        if row[0] and (row[1] == '总计' or row[1] is None) and row[0] != '海外团队':
            if any([row[2], row[13]]):
                groups.append(row)
    conc_lines.append('<p>&nbsp;</p>')
    groups_sorted = sorted([r for r in groups if r[2]], key=lambda r: r[2] or 0, reverse=True)
    if groups_sorted:
        top = ' / '.join(f"{r[0]}({fmt_int(r[2])}条)" for r in groups_sorted[:5])
        conc_lines.append(f'<p>营销类群发消息数 TOP5（小组）：<span class="hl-pink">{top}</span></p>')

    # zero/low marketing groups
    zero = [r for r in groups if (r[2] or 0) == 0]
    if zero:
        conc_lines.append(f'<p>营销类群发为 0 的小组：<span class="hl-pink">{" / ".join(r[0] for r in zero)}</span></p>')

    # service top
    svc_sorted = sorted([r for r in groups if r[13]], key=lambda r: r[13] or 0, reverse=True)
    if svc_sorted:
        top = ' / '.join(f"{r[0]}({fmt_int(r[13])}条)" for r in svc_sorted[:5])
        conc_lines.append(f'<p>服务类群发消息数 TOP5（小组）：<span class="hl-pink">{top}</span></p>')

    out.append('<h2 style="font-size:14px;border-left:4px solid #4472C4;">结论</h2>')
    out.append('<div class="callout"><div class="callout-emoji">❗</div><div class="callout-body">')
    out.extend(conc_lines)
    out.append('</div></div>')

    # table
    out.append('<div class="table-wrap"><table class="report-table"><thead>')
    out.append('<tr class="group-h">'
               '<th colspan="2" class="gh" style="background:#D9D9D9;color:#000">基础</th>'
               '<th colspan="6" class="gh" style="background:#4472C4;color:#fff">营销类LP群发-汇总</th>'
               '<th colspan="5" class="gh" style="background:#BDD7EE;color:#000">营销类LP群发-重复内容</th>'
               '<th colspan="6" class="gh" style="background:#4472C4;color:#fff">服务类LP群发-汇总</th>'
               '<th colspan="7" class="gh" style="background:#C6E0B4;color:#000">服务类LP群发-重复内容</th>'
               '<th colspan="1" class="gh" style="background:#F4B084;color:#000">其他</th>'
               '</tr>')
    headers = ['小组', 'LP姓名',
               '群发消息数', '环比上月', '个人群发消息数', '环比上月', '个人群发占比', '群发学员数',
               '续费', '转介绍', '拓美术', '拓叫叫', '拓英语',
               '群发消息数', '环比上月', '个人群发消息数', '环比上月', '个人群发占比', '群发学员数',
               '通知-学员流转', '服务-学员流转', '服务-问卷调查', '服务-节日祝福', '服务-停课通知', '学习总结', '打卡',
               '其他LP群发-重复内容']
    bg = ['#D9D9D9', '#D9D9D9'] + ['#DCE6F1'] * 6 + ['#BDD7EE'] * 5 + ['#E2EFDA'] * 6 + ['#C6E0B4'] * 7 + ['#F4B084']
    th_row = '<tr class="col-row">'
    for i, (h, c) in enumerate(zip(headers, bg)):
        cls = ' frozen' if i < 2 else ''
        th_row += f'<th class="ch{cls}" style="background:{c}">{h}</th>'
    th_row += '</tr>'
    out.append(th_row)
    out.append('</thead><tbody>')

    def render_row(row, kind='detail'):
        cls = 'summary' if kind == 'summary' else 'detail'
        cells = []
        cells.append(f'<td class="frozen">{row[0] or ""}</td>')
        cells.append(f'<td class="frozen">{row[1] or ""}</td>')
        # 营销类汇总
        cells.append(f'<td>{fmt_int(row[2])}</td>')
        cells.append(f'<td>{fmt_pct(row[3],2)}</td>')
        cells.append(f'<td>{fmt_int(row[4])}</td>')
        cells.append(f'<td>{fmt_pct(row[5],2)}</td>')
        cells.append(f'<td>{fmt_pct(row[6],2)}</td>')
        cells.append(f'<td>{fmt_int(row[7])}</td>')
        # 营销重复
        for ci in range(8, 13):
            cells.append(f'<td>{fmt_int(row[ci])}</td>')
        # 服务汇总
        cells.append(f'<td>{fmt_int(row[13])}</td>')
        cells.append(f'<td>{fmt_pct(row[14],2)}</td>')
        cells.append(f'<td>{fmt_int(row[15])}</td>')
        cells.append(f'<td>{fmt_pct(row[16],2)}</td>')
        cells.append(f'<td>{fmt_pct(row[17],2)}</td>')
        cells.append(f'<td>{fmt_int(row[18])}</td>')
        # 服务重复
        for ci in range(19, 26):
            cells.append(f'<td>{fmt_int(row[ci])}</td>')
        # 其他
        cells.append(f'<td>{fmt_int(safe_get(row, 26))}</td>')
        return f'<tr class="{cls}">' + ''.join(cells) + '</tr>'

    # body
    summary_groups = ['海外团队', '台湾区', '港澳区', '欧美澳区', '台湾组', '港澳1组', '港澳2组', '港澳组', '美澳1组', '美澳2组', '美澳3组', '美澳4组', '美澳5组']
    for i in range(2, len(v)):
        row = v[i]
        if not row[0] and not row[1]:
            continue
        if row[1] == '总计' or (row[0] and not row[1]):
            kind = 'summary'
        else:
            kind = 'detail'
        out.append(render_row(row, kind))

    out.append('</tbody></table></div>')
    return '\n'.join(out)


def safe_get(row, idx):
    try:
        return row[idx]
    except (IndexError, TypeError):
        return None


# === 4.4 停课唤醒 ===
def build_4_4():
    v = load_values('4_4_raw.json')
    out = []
    out.append('<h2>4.4 停课唤醒</h2>')
    out.append('<div class="meta">数据周期：2026/05/01 — 2026/05/25（5月第三个自然周末日）｜ 报表：思维停课学员执行监控</div>')

    overall = v[3]
    # row index for 海外团队 总计?  Let me check by looking at row[0]
    # Actually the first data row should be 海外团队 总计
    # Try row 3
    # cols:
    # 0 lp组别, 1 LP个人, 2 停课占比目标, 3 停课占比, 4 有效在读人数, 5 当前实际停课人数, 6 当前实际停课执行中学员数, 7 执行中学员停课率, 8 停课GAP, 9 本月已排课未复课, 10 次月复课, 11 停课数,
    # 12 衰退学员数, 13 不可唤醒, 14 唤醒学员数, 15 唤醒率, 16-19 外呼指标(本月停课目标),
    # 20 90天内 停课数, 21 停课占比, 22 唤醒数, 23 唤醒率, 24-27 外呼,
    # 28 90天以上 停课数, 29 停课占比, 30 唤醒数, 31 唤醒率, 32-35 外呼,
    # 36 当月新增 停课数, 37-40 外呼,
    # 41 停课预警 停课数, 42-45 外呼

    # Find 海外团队 总计 row
    overall_row = None
    for r in v[3:8]:
        if r[0] in ('海外团队', '海外'):
            overall_row = r
            break
    if not overall_row:
        overall_row = v[3]

    stop_pct = overall_row[3]
    awaken_90_pct = overall_row[23]
    overall_calls = overall_row[16]  # 外呼跟进率 of 本月停课目标

    conc_lines = []
    conc_lines.append(f'<p><b>当前整体停课占比 {fmt_pct(stop_pct,2)}（停课目标 {fmt_pct(overall_row[2],2)}）</b></p>')
    conc_lines.append(f'<p><b>整体停课 90 天内唤醒率 {fmt_pct(awaken_90_pct,2)}</b></p>')
    conc_lines.append(f'<p>整体本月停课目标外呼跟进率 {fmt_pct(overall_calls,2)}</p>')

    # find groups with high 停课占比
    high = []
    for i in range(3, len(v)):
        r = v[i]
        if r[0] and (r[1] == '总计' or r[1] is None) and r[0] not in ('海外团队', '海外', '台湾区', '港澳区', '欧美澳区'):
            if r[3] is not None and r[2] is not None:
                high.append(r)
    over_target = sorted([r for r in high if (r[3] or 0) > (r[2] or 0)], key=lambda r: (r[3] or 0) - (r[2] or 0), reverse=True)
    if over_target:
        ss = ' / '.join(f"{r[0]}（停课占比 {fmt_pct(r[3],2)}/目标 {fmt_pct(r[2],2)}）" for r in over_target[:6])
        conc_lines.append(f'<p>停课占比超过目标的小组：<span class="hl-pink">{ss}</span></p>')

    # 90 days awaken rate low
    low_aw = sorted([r for r in high if r[23] is not None], key=lambda r: r[23] or 0)
    if low_aw:
        ss = ' / '.join(f"{r[0]}({fmt_pct(r[23],2)})" for r in low_aw[:5])
        conc_lines.append(f'<p>90 天内唤醒率较低的小组：<span class="hl-pink">{ss}</span></p>')

    out.append('<h2 style="font-size:14px;border-left:4px solid #4472C4;">结论</h2>')
    out.append('<div class="callout"><div class="callout-emoji">❗</div><div class="callout-body">')
    out.extend(conc_lines)
    out.append('</div></div>')

    # table — full 46 cols
    out.append('<div class="table-wrap"><table class="report-table"><thead>')
    out.append('<tr class="group-h">'
               '<th colspan="2" class="gh" style="background:#D9D9D9;color:#000">基础</th>'
               '<th colspan="9" class="gh" style="background:#4472C4;color:#fff">停课情况</th>'
               '<th colspan="9" class="gh" style="background:#BDD7EE;color:#000">本月停课唤醒目标学员</th>'
               '<th colspan="8" class="gh" style="background:#C6E0B4;color:#000">停课90天内唤醒</th>'
               '<th colspan="8" class="gh" style="background:#FFD966;color:#000">停课90天以上唤醒</th>'
               '<th colspan="5" class="gh" style="background:#F4B084;color:#000">当月新增停课待唤醒</th>'
               '<th colspan="5" class="gh" style="background:#E7E6E6;color:#000">停课预警学员</th>'
               '</tr>')
    headers = ['lp组别', 'LP个人',
               '停课占比目标', '停课占比', '有效在读人数', '当前实际停课人数', '当前实际停课执行中学员数', '执行中学员停课率', '停课GAP', '本月已排课未复课人数', '次月复课人数',
               '停课数', '衰退学员数', '不可唤醒学员', '唤醒学员数(本月签到)', '唤醒率', '外呼跟进率', '生均拨打次数', '外呼接通率', '生均有效通话时长',
               '停课数', '停课占比', '唤醒数', '唤醒率', '外呼跟进率', '生均拨打次数', '外呼接通率', '生均有效通话时长',
               '停课数', '停课占比', '唤醒数', '唤醒率', '外呼跟进率', '生均拨打次数', '外呼接通率', '生均有效通话时长',
               '停课数', '外呼跟进率', '生均拨打次数', '外呼接通率', '生均有效通话时长',
               '停课数', '外呼跟进率', '生均拨打次数', '外呼接通率', '生均有效通话时长']
    bg = ['#D9D9D9', '#D9D9D9'] + ['#DCE6F1'] * 9 + ['#BDD7EE'] * 9 + ['#C6E0B4'] * 8 + ['#FFD966'] * 8 + ['#F4B084'] * 5 + ['#E7E6E6'] * 5
    th_row = '<tr class="col-row">'
    for i, (h, c) in enumerate(zip(headers, bg)):
        cls = ' frozen' if i < 2 else ''
        th_row += f'<th class="ch{cls}" style="background:{c}">{h}</th>'
    th_row += '</tr>'
    out.append(th_row)
    out.append('</thead><tbody>')

    pct_cols = {2, 3, 7, 15, 16, 18, 21, 23, 24, 26, 29, 31, 32, 34, 37, 39, 42, 44}
    num2_cols = {17, 19, 25, 27, 33, 35, 38, 40, 43, 45}

    def render_row(row, kind='detail'):
        cls = 'summary' if kind == 'summary' else 'detail'
        cells = []
        cells.append(f'<td class="frozen">{row[0] or ""}</td>')
        cells.append(f'<td class="frozen">{row[1] or ""}</td>')
        for ci in range(2, 46):
            val = safe_get(row, ci)
            if ci in pct_cols:
                cells.append(f'<td>{fmt_pct(val,2)}</td>')
            elif ci in num2_cols:
                cells.append(f'<td>{fmt_num(val,2)}</td>')
            else:
                cells.append(f'<td>{fmt_int(val)}</td>')
        return f'<tr class="{cls}">' + ''.join(cells) + '</tr>'

    for i in range(2, len(v)):
        row = v[i]
        if not row[0] and not row[1]:
            continue
        if (row[1] == '总计' or (row[0] and not row[1])) and row[0]:
            kind = 'summary'
        else:
            kind = 'detail'
        out.append(render_row(row, kind))

    out.append('</tbody></table></div>')
    return '\n'.join(out)


# === 4.5 服务月跟进 + SOP ===
def build_4_5():
    v = load_values('4_5_raw.json')
    out = []
    out.append('<h2>4.5 服务月跟进</h2>')
    out.append('<div class="meta">数据周期：2026/05/01 — 2026/05/25（5月第三个自然周末日）｜ 报表：思维转介绍过程跟进报表_末次渠道（仅服务池） + 海外思维服务SOP执行情况</div>')

    overall = v[2]
    # cols: 0 小组, 1 LP, 2 学员数, 3 外呼跟进率, 4 外呼有效跟进率, 5 微信外呼跟进率, 6 微信覆盖率, 7 微信有效回复率, 8 综合有效跟进率,
    # 9 总外呼次数, 10 总外呼时长_min, 11 生均外呼次数, 12 带R数, 13 带R效率, 14 秒挂占比

    overall_call_pct = overall[3]
    overall_eff = overall[8]

    conc_lines = []
    conc_lines.append('<p><b>服务池跟进：</b></p>')
    conc_lines.append(f'<p>外呼覆盖率 {fmt_pct(overall_call_pct,2)}，综合有效跟进达 {fmt_pct(overall_eff,2)}；微信覆盖率 {fmt_pct(overall[6],2)}，微信有效回复率 {fmt_pct(overall[7],2)}</p>')
    conc_lines.append(f'<p>总外呼次数 {fmt_int(overall[9])}，总外呼时长 {fmt_int(overall[10])} min，生均外呼次数 {fmt_num(overall[11],2)}，带 R 数 {fmt_int(overall[12])}（带 R 效率 {fmt_pct(overall[13],2)}）</p>')
    # find low-call groups
    groups = []
    for i in range(7, 16):
        r = v[i]
        if r[0] and r[1] == '总计':
            groups.append(r)
    low_call = sorted(groups, key=lambda r: r[3] or 0)
    # report all groups below overall
    below = [r for r in groups if (r[3] or 0) < (overall_call_pct or 0.85)]
    if below:
        below.sort(key=lambda r: r[3] or 0)
        for worst in below[:3]:
            # find low individual within worst group
            low_lps = []
            cur_grp = None
            for j in range(16, len(v)):
                rr = v[j]
                if rr[0]:
                    cur_grp = rr[0]
                if cur_grp == worst[0] and rr[1] and rr[1] != '总计':
                    if rr[3] is not None and rr[3] < (overall_call_pct or 0.85) - 0.1:
                        low_lps.append((rr[1], rr[3]))
            names = ' / '.join(f"{n}({fmt_pct(p,2)})" for n, p in low_lps[:5])
            line = f'<p><span class="hl-pink">{worst[0]}</span> 服务池外呼跟进 <span class="hl-pink">{fmt_pct(worst[3],2)}</span>，低于整体'
            if names:
                line += f'，需关注：{names}'
            line += '</p>'
            conc_lines.append(line)

    out.append('<h2 style="font-size:14px;border-left:4px solid #4472C4;">结论</h2>')
    out.append('<div class="callout"><div class="callout-emoji">❗</div><div class="callout-body">')
    out.extend(conc_lines)
    out.append('</div></div>')

    # === Service pool follow-up table ===
    out.append('<h2 style="font-size:14px;border-left:4px solid #4472C4;">服务池跟进数据表</h2>')
    out.append('<div class="table-wrap"><table class="report-table"><thead>')
    out.append('<tr class="group-h"><th colspan="2" class="gh" style="background:#D9D9D9;color:#000">基础</th>'
               '<th colspan="13" class="gh" style="background:#4472C4;color:#fff">服务池</th></tr>')
    headers = ['小组', 'LP', '学员数', '外呼跟进率', '外呼有效跟进率', '微信外呼跟进率', '微信覆盖率', '微信有效回复率', '综合有效跟进率', '总外呼次数', '总外呼时长_min', '生均外呼次数', '带R数', '带R效率', '秒挂占比']
    bg = ['#D9D9D9', '#D9D9D9'] + ['#BDD7EE'] * 13
    th = '<tr class="col-row">'
    for i, (h, c) in enumerate(zip(headers, bg)):
        cls = ' frozen' if i < 2 else ''
        th += f'<th class="ch{cls}" style="background:{c}">{h}</th>'
    th += '</tr>'
    out.append(th)
    out.append('</thead><tbody>')

    pct_cols = {3, 4, 5, 6, 7, 8, 13, 14}
    num2_cols = {11}

    def render_row(row, kind='detail'):
        cls = 'summary' if kind == 'summary' else 'detail'
        cells = [f'<td class="frozen">{row[0] or ""}</td>', f'<td class="frozen">{row[1] or ""}</td>']
        for ci in range(2, 15):
            val = safe_get(row, ci)
            if ci in pct_cols:
                cells.append(f'<td>{fmt_pct(val,2)}</td>')
            elif ci in num2_cols:
                cells.append(f'<td>{fmt_num(val,2)}</td>')
            else:
                cells.append(f'<td>{fmt_int(val)}</td>')
        return f'<tr class="{cls}">' + ''.join(cells) + '</tr>'

    for i in range(2, len(v)):
        row = v[i]
        if not row[0] and not row[1]:
            continue
        # only keep service-pool data; this sheet is already 服务池-only
        if (row[1] == '总计' or (row[0] and not row[1])):
            kind = 'summary'
        else:
            kind = 'detail'
        out.append(render_row(row, kind))

    out.append('</tbody></table></div>')

    # === SOP table ===
    out.append('<h2 style="font-size:14px;border-left:4px solid #4472C4;">海外思维服务 SOP 执行情况（服务池）</h2>')
    sop = load_values('4_5_sop_uv.json')
    sop_fv = load_values('4_5_sop_fv.json')
    # We use unformatted for math; formatted to display
    # row 3 海外团队 总计; rows 5-13 group totals; rows 15+ individuals
    overall_sop = sop[3]
    w1_total = overall_sop[3]
    w2_total = overall_sop[9]
    if isinstance(w1_total, str):
        w1_total = 0
    if isinstance(w2_total, str):
        w2_total = 0
    sum_exec = (w1_total or 0) + (w2_total or 0)

    # Build sub-conclusion: list groups with low SOP execution
    sop_groups = []
    for i in range(5, 14):
        r = sop[i]
        if r[0] and r[2] == '总计':
            w1 = r[3] if isinstance(r[3], (int, float)) else 0
            w2 = r[9] if isinstance(r[9], (int, float)) else 0
            sop_groups.append((r[0], w1 + w2, w1, w2))
    sop_groups.sort(key=lambda x: x[1])
    target = 2.4  # service pool semantic target

    callout_lines = []
    callout_lines.append(f'<p><b>SOP 语义点执行率加和（服务池 W1+W2）：海外团队 {fmt_num(sum_exec,1)}（W1 {fmt_num(w1_total,1)} / W2 {fmt_num(w2_total,1)}）</b></p>')
    if sum_exec < target:
        callout_lines.append(f'<p>未达服务池语义点加和目标（{target}），整体偏低 {target - sum_exec:.1f}</p>')
    low_sop = [g for g in sop_groups if g[1] < target]
    if low_sop:
        ss = ' / '.join(f"{g[0]}({fmt_num(g[1],1)})" for g in low_sop[:6])
        callout_lines.append(f'<p>语义点执行加和较低的小组：<span class="hl-pink">{ss}</span></p>')

    out.append('<div class="callout" style="background:#fff8e6;border-color:#ffe2a8;"><div class="callout-emoji">📌</div>'
               + '<div class="callout-body">' + ''.join(callout_lines) + '</div></div>')

    out.append('<div class="table-wrap"><table class="report-table"><thead>')
    out.append('<tr class="group-h">'
               '<th colspan="3" class="gh" style="background:#D9D9D9;color:#000">基础</th>'
               '<th colspan="1" class="gh" style="background:#FFF2CC;color:#000">服务池-LP合计</th>'
               '<th colspan="6" class="gh" style="background:#BDD7EE;color:#000">服务池-W1</th>'
               '<th colspan="6" class="gh" style="background:#C6E0B4;color:#000">服务池-W2</th>'
               '</tr>')
    headers = ['小组', '负责人', 'LP',
               '执行率加和(W1+W2)',
               '语义点执行率加和', '命中服务池学员数', '学情反馈执行率', '学习规划执行率', '告知转介绍权益执行率', '告知剩余课时执行率',
               '语义点执行率加和', '命中服务池学员数', '学情反馈执行率', '学习规划执行率', '告知转介绍权益执行率', '告知剩余课时执行率']
    bg = ['#D9D9D9', '#D9D9D9', '#D9D9D9', '#FFF2CC'] + ['#DCE6F1'] * 6 + ['#E2EFDA'] * 6
    th = '<tr class="col-row">'
    for i, (h, c) in enumerate(zip(headers, bg)):
        cls = ' frozen' if i < 3 else ''
        th += f'<th class="ch{cls}" style="background:{c}">{h}</th>'
    th += '</tr>'
    out.append(th)
    out.append('</thead><tbody>')

    pct_cols = {5, 6, 7, 8, 11, 12, 13, 14}  # adjusted indexes after we insert sum column at index 3
    # In source data: cols are 0小组,1负责人,2LP,3 W1语义点加和,4 W1命中,5-8 W1四个执行率,9 W2语义点加和,10 W2命中,11-14 W2四个执行率
    # Output we add 加和 column after 2: so output cols 0,1,2 = 小组/负责人/LP, 3 = sum, 4=W1加和, 5=W1命中, 6-9=W1 pcts, 10=W2加和, 11=W2命中, 12-15=W2 pcts

    def num_or_zero(v):
        if v is None or v == '':
            return 0
        if isinstance(v, str):
            try:
                return float(v)
            except:
                return 0
        return v

    def sop_row(row_uv, kind):
        cls = 'summary' if kind == 'summary' else 'detail'
        w1 = num_or_zero(safe_get(row_uv, 3))
        w2 = num_or_zero(safe_get(row_uv, 9))
        s = w1 + w2
        cells = [f'<td class="frozen">{row_uv[0] or ""}</td>',
                 f'<td class="frozen">{row_uv[1] or ""}</td>',
                 f'<td class="frozen">{row_uv[2] or ""}</td>',
                 f'<td>{fmt_num(s,1)}</td>']
        # W1: 加和(3), 命中(4), 学情反馈(5), 学习规划(6), 告知R(7), 告知课时(8)
        cells.append(f'<td>{fmt_num(safe_get(row_uv,3),1)}</td>')
        cells.append(f'<td>{fmt_int(safe_get(row_uv,4))}</td>')
        for ci in [5, 6, 7, 8]:
            cells.append(f'<td>{fmt_pct(safe_get(row_uv,ci),2)}</td>')
        # W2
        cells.append(f'<td>{fmt_num(safe_get(row_uv,9),1)}</td>')
        cells.append(f'<td>{fmt_int(safe_get(row_uv,10))}</td>')
        for ci in [11, 12, 13, 14]:
            cells.append(f'<td>{fmt_pct(safe_get(row_uv,ci),2)}</td>')
        return f'<tr class="{cls}">' + ''.join(cells) + '</tr>'

    for i in range(3, len(sop)):
        row = sop[i]
        if not any([row[0], row[1], row[2]]):
            continue
        if row[2] == '总计' or (row[0] and row[2] == '总计'):
            kind = 'summary'
        else:
            kind = 'detail'
        out.append(sop_row(row, kind))

    out.append('</tbody></table></div>')

    return '\n'.join(out)


# === 4.6 系统电话外呼 & 微信回复监控 ===
def build_4_6():
    out = []
    out.append('<h2>4.6 系统电话外呼 &amp; 微信回复监控</h2>')
    out.append('<div class="meta">数据周期：2026/05/01 — 2026/05/25（5月第三个自然周末日）｜ 报表：LP系统外呼监控-分池子 / LP企微回复比监控-分池子</div>')

    # ===== Call data =====
    v = load_values('4_6_call_raw.json')
    overall = v[2]
    # 0小组,1LP姓名 then groups: 整体(C-Q) 续费池(R-AF) 服务池(AG-AU) M1-M3(AV-BJ) 上月底(BK-BY) 其他(BZ-CN)
    # Each pool has 15 metrics: 学员量, 外呼人数, 覆盖率, 覆盖率月环比, 外呼次数, 生均呼次, 生均外呼次数月环比, 外呼接通率, 生均通次, 生均通时_min, 有效接通率, 生均有效通次, 生均有效通时_min, 企微-生均通次, 企微-生均通时_min
    # But integer count differs - let me count: 15? row1 has cols 2..16 for "整体" = 15 metrics. Then 续费池 starts at 17. Yes 15 cols per pool.
    # Wait: row 0 group header lengths from indices: 整体(2-16=15), 续费池(17-31=15), 服务池(32-46=15), M1-M3(47-61=15), 上月底(62-76=15), 其他(77-91=15)
    # total = 2 + 15*6 = 92 ✓

    # === Call conclusion ===
    overall_cov = overall[4]  # 整体-覆盖率
    overall_avg = overall[7]  # 整体-生均呼次
    conc_lines = []
    conc_lines.append('<p><b>整体系统外呼：</b></p>')
    conc_lines.append(f'<p>全部活跃学员的外呼覆盖率为 {fmt_pct(overall_cov,2)}（环比 {fmt_pct(overall[5],2)}），生均呼次 {fmt_num(overall_avg,2)}（环比 {fmt_pct(overall[8],2)}）</p>')
    # find groups with low overall coverage
    grp = []
    for i in range(3, 13):
        r = v[i]
        if r[0] and r[1] == '总计':
            grp.append(r)
    low_cov = sorted([r for r in grp if r[4] is not None], key=lambda r: r[4] or 0)
    if low_cov:
        worst = low_cov[:3]
        ss = ' / '.join(f"{r[0]}({fmt_pct(r[4],2)})" for r in worst)
        conc_lines.append(f'<p>整体外呼覆盖率较低：<span class="hl-pink">{ss}</span></p>')
    # 服务池 covers
    svc_cov = overall[33]  # 32 学员量, 33 外呼人数, 34 覆盖率
    svc_avg = overall[37]
    conc_lines.append(f'<p>服务池外呼覆盖率 {fmt_pct(safe_get(overall,34),2)}，生均呼次 {fmt_num(safe_get(overall,37),2)}</p>')

    # ===== Wechat data =====
    w = load_values('4_6_wechat_raw.json')
    woverall = w[2]
    # cols: 0当前小组, 1LP姓名, 2学员量, 3发送方式, 4发送消息条数, 5生均发送条数, 6回复比, 7-10 续费池(学员量,消息数,生均,回复比), 11-14 服务池, 15-18 M1-M3, 19-22 上月底, 23-26 其他
    overall_total_msg = woverall[4]
    overall_avg_msg = woverall[5]
    overall_reply_ratio = woverall[6]
    conc_lines.append('<p>&nbsp;</p>')
    conc_lines.append('<p><b>整体微信发送 &amp; 回复比：</b></p>')
    conc_lines.append(f'<p>整体群发消息数 {fmt_int(overall_total_msg)} 条，生均发送 {fmt_num(overall_avg_msg,2)} 条，回复比 {fmt_num(overall_reply_ratio,2)}</p>')
    conc_lines.append(f'<p>服务池消息数 {fmt_int(woverall[12])} 条，生均 {fmt_num(woverall[13],2)} 条，回复比 {fmt_num(woverall[14],2)}；续费池消息数 {fmt_int(woverall[8])} 条，生均 {fmt_num(woverall[9],2)} 条，回复比 {fmt_num(woverall[10],2)}</p>')

    out.append('<h2 style="font-size:14px;border-left:4px solid #4472C4;">结论</h2>')
    out.append('<div class="callout"><div class="callout-emoji">❗</div><div class="callout-body">')
    out.extend(conc_lines)
    out.append('</div></div>')

    # ===== Call table =====
    out.append('<h2 style="font-size:14px;border-left:4px solid #4472C4;">LP 系统外呼监控（分池子）</h2>')
    out.append('<div class="table-wrap"><table class="report-table"><thead>')
    pool_cols = ['学员量', '外呼人数', '覆盖率', '覆盖率月环比', '外呼次数', '生均呼次', '生均外呼次数月环比', '外呼接通率', '生均通次', '生均通时_min', '有效接通率', '生均有效通次', '生均有效通时_min', '企微-生均通次', '企微-生均通时_min']
    pool_names = ['整体', '续费池', '服务池', 'M1-M3', '上月底续费(上月最后7天续费)', '其他非做工池']
    pool_bg = ['#DCE6F1', '#FFE699', '#C6E0B4', '#F4B084', '#D9D2E9', '#E7E6E6']
    out.append('<tr class="group-h">'
               '<th colspan="2" class="gh" style="background:#D9D9D9;color:#000">基础</th>'
               + ''.join(f'<th colspan="{len(pool_cols)}" class="gh" style="background:{pool_bg[i]};color:#000">{pool_names[i]}</th>' for i in range(6))
               + '</tr>')
    th = '<tr class="col-row">'
    th += '<th class="ch frozen" style="background:#D9D9D9">小组</th>'
    th += '<th class="ch frozen" style="background:#D9D9D9">LP姓名</th>'
    for pi in range(6):
        for c in pool_cols:
            th += f'<th class="ch" style="background:{pool_bg[pi]}">{c}</th>'
    th += '</tr>'
    out.append(th)
    out.append('</thead><tbody>')

    # determine cell types per col offset within pool
    pct_cols = {2, 3, 6, 7, 10}  # 覆盖率, 覆盖率月环比, 生均外呼次数月环比, 外呼接通率, 有效接通率
    num2_cols = {5, 8, 9, 11, 12, 13, 14}

    def render_call_row(row, kind):
        cls = 'summary' if kind == 'summary' else 'detail'
        cells = [f'<td class="frozen">{row[0] or ""}</td>', f'<td class="frozen">{row[1] or ""}</td>']
        # 6 pools × 15 cols starting at col 2
        for pi in range(6):
            base = 2 + pi * 15
            for j in range(15):
                ci = base + j
                val = safe_get(row, ci)
                if j in pct_cols:
                    cells.append(f'<td>{fmt_pct(val,2)}</td>')
                elif j in num2_cols:
                    cells.append(f'<td>{fmt_num(val,2)}</td>')
                else:
                    cells.append(f'<td>{fmt_int(val)}</td>')
        return f'<tr class="{cls}">' + ''.join(cells) + '</tr>'

    for i in range(2, len(v)):
        row = v[i]
        if not row[0] and not row[1]:
            continue
        if row[1] == '总计' or (row[0] and not row[1]):
            kind = 'summary'
        else:
            kind = 'detail'
        out.append(render_call_row(row, kind))

    out.append('</tbody></table></div>')

    # ===== Wechat table =====
    out.append('<h2 style="font-size:14px;border-left:4px solid #4472C4;">LP 企微回复比监控（分池子）</h2>')
    out.append('<div class="table-wrap"><table class="report-table"><thead>')
    out.append('<tr class="group-h">'
               '<th colspan="3" class="gh" style="background:#D9D9D9;color:#000">基础</th>'
               '<th colspan="3" class="gh" style="background:#DCE6F1;color:#000">整体</th>'
               '<th colspan="4" class="gh" style="background:#FFE699;color:#000">续费池</th>'
               '<th colspan="4" class="gh" style="background:#C6E0B4;color:#000">服务池</th>'
               '<th colspan="4" class="gh" style="background:#F4B084;color:#000">M1-M3</th>'
               '<th colspan="4" class="gh" style="background:#D9D2E9;color:#000">上月底续费(上月最后7天续费)</th>'
               '<th colspan="4" class="gh" style="background:#E7E6E6;color:#000">其他非做工池</th>'
               '</tr>')
    th = '<tr class="col-row">'
    th += '<th class="ch frozen" style="background:#D9D9D9">小组</th>'
    th += '<th class="ch frozen" style="background:#D9D9D9">LP姓名</th>'
    th += '<th class="ch frozen" style="background:#D9D9D9">发送方式</th>'
    # 整体: 学员量, 发送消息条数, 生均发送条数, 回复比 (note: 学员量 only on '总计' line)
    th += '<th class="ch" style="background:#DCE6F1">发送消息条数</th>'
    th += '<th class="ch" style="background:#DCE6F1">生均发送条数</th>'
    th += '<th class="ch" style="background:#DCE6F1">回复比</th>'
    pool_bg2 = ['#FFE699', '#C6E0B4', '#F4B084', '#D9D2E9', '#E7E6E6']
    pool_names2 = ['续费池', '服务池', 'M1-M3', '上月底', '其他']
    for pi in range(5):
        for c in ['学员量', '发送消息条数', '生均发送条数', '回复比']:
            th += f'<th class="ch" style="background:{pool_bg2[pi]}">{c}</th>'
    th += '</tr>'
    out.append(th)
    out.append('</thead><tbody>')

    def render_wc_row(row, kind):
        cls = 'summary' if kind == 'summary' else 'detail'
        cells = [f'<td class="frozen">{row[0] or ""}</td>',
                 f'<td class="frozen">{row[1] or ""}</td>',
                 f'<td class="frozen">{row[3] or ""}</td>']  # 发送方式
        # 整体: skip 学员量(2) since shown only on first line; show msg, avg, reply
        cells.append(f'<td>{fmt_int(safe_get(row,4))}</td>')
        cells.append(f'<td>{fmt_num(safe_get(row,5),2)}</td>')
        cells.append(f'<td>{fmt_num(safe_get(row,6),2)}</td>')
        # pools start at col 7 (续费池: 7,8,9,10), 11-14, 15-18, 19-22, 23-26
        for base in [7, 11, 15, 19, 23]:
            cells.append(f'<td>{fmt_int(safe_get(row,base))}</td>')      # 学员量
            cells.append(f'<td>{fmt_int(safe_get(row,base+1))}</td>')    # 消息数
            cells.append(f'<td>{fmt_num(safe_get(row,base+2),2)}</td>')  # 生均
            cells.append(f'<td>{fmt_num(safe_get(row,base+3),2)}</td>')  # 回复比
        return f'<tr class="{cls}">' + ''.join(cells) + '</tr>'

    for i in range(2, len(w)):
        row = w[i]
        if not row[0] and not row[1] and not row[3]:
            continue
        if row[3] == '总计':
            kind = 'summary'
        else:
            kind = 'detail'
        out.append(render_wc_row(row, kind))

    out.append('</tbody></table></div>')

    return '\n'.join(out)


def main():
    parts = []
    parts.append(build_4_2())
    parts.append(build_4_3())
    parts.append(build_4_4())
    parts.append(build_4_5())
    parts.append(build_4_6())

    section_html = '\n\n'.join(parts)

    out_path = os.path.join(DATA_DIR, '4_2_to_4_6_section.html')
    with open(out_path, 'w', encoding='utf-8') as f:
        f.write(section_html)
    print(f'Wrote {len(section_html)} bytes to {out_path}')


if __name__ == '__main__':
    main()
