"""自然周日期工具

约定（与用户对齐的默认规则）：
- 自然周 = 周一 00:00 ~ 周日 23:59:59（ISO 周）
- 「某月第 N 个自然周」= 该月内第 N 个 *完整自然周*；若该月 1 号不是周一，1 号所在那周不计入
- 周报默认以"当月当周"作为参考

提供：
- iso_week_range(d): 给定日期所在自然周的 (周一, 周日)
- nth_week_of_month(year, month, n): 返回该月第 n 个完整自然周的 (周一, 周日)
- week_offset(d, hours=...): 在某个日期/时刻上做小时级偏移
- bi_date_args(...): 一站式生成 bi_skill 所需的 --extra-dates 字符串
"""
from __future__ import annotations

from datetime import date, datetime, time, timedelta


def iso_week_range(d: date) -> tuple[date, date]:
    monday = d - timedelta(days=d.weekday())
    sunday = monday + timedelta(days=6)
    return monday, sunday


def nth_week_of_month(year: int, month: int, n: int) -> tuple[date, date]:
    """该月第 n 个完整自然周的 (周一, 周日)。

    例：2026年5月第3周 → 2026-05-18 ~ 2026-05-24
    """
    if n < 1:
        raise ValueError("n must be >= 1")
    first = date(year, month, 1)
    # 月内第一个周一
    first_monday = first + timedelta(days=(7 - first.weekday()) % 7)
    monday = first_monday + timedelta(weeks=n - 1)
    sunday = monday + timedelta(days=6)
    if monday.month != month and sunday.month != month:
        raise ValueError(f"{year}年{month}月没有第{n}个自然周")
    return monday, sunday


def week_offset(value: date | datetime, *, hours: int = 0) -> datetime:
    """在某个日期/时刻上做小时级偏移。返回 datetime。"""
    if isinstance(value, datetime):
        dt = value
    else:
        dt = datetime.combine(value, time(0, 0))
    return dt + timedelta(hours=hours)


def fmt(value: date | datetime, with_time: bool = False) -> str:
    if isinstance(value, datetime):
        if with_time:
            return value.strftime("%Y-%m-%d %H:%M:%S")
        return value.strftime("%Y-%m-%d")
    return value.strftime("%Y-%m-%d")


def end_of_day(d: date) -> datetime:
    return datetime.combine(d, time(23, 59, 59))


def bi_date_args(pairs: list[tuple[str, date | datetime, bool]]) -> str:
    """生成 bi_skill 的 --extra-dates 字符串。

    pairs: [(字段名, 日期或datetime, 是否带时间)]
    """
    parts = []
    for name, value, with_time in pairs:
        parts.append(f"{name}={fmt(value, with_time=with_time)}")
    return ",".join(parts)


if __name__ == "__main__":
    # 一些快速 sanity check
    monday, sunday = nth_week_of_month(2026, 5, 3)
    assert monday == date(2026, 5, 18), monday
    assert sunday == date(2026, 5, 24), sunday

    # 周日 23:59:59 提前 48 小时
    eod_sunday = end_of_day(sunday)
    minus48 = week_offset(eod_sunday, hours=-48)
    assert minus48 == datetime(2026, 5, 22, 23, 59, 59), minus48

    minus72 = week_offset(eod_sunday, hours=-72)
    assert minus72 == datetime(2026, 5, 21, 23, 59, 59), minus72

    print("[OK] 自然周日期工具自检通过")
    print(f"  2026年5月第3周: {monday} ~ {sunday}")
    print(f"  周日 23:59:59 -48h: {minus48}")
    print(f"  周日 23:59:59 -72h: {minus72}")
