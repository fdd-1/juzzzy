"""日期工具：计算上周一/上周日，处理时间窗口偏移。"""
from __future__ import annotations
from datetime import date, datetime, time, timedelta


def last_week_range(today: date | None = None) -> tuple[date, date]:
    """返回最近一个完整自然周（上周一 ~ 上周日）。

    如果今天是周一 → 返回上周一到上周日（昨天）。
    如果今天是周日 → 返回上周一到上周日（今天）。
    """
    if today is None:
        today = date.today()
    # 本周一
    this_monday = today - timedelta(days=today.weekday())
    # 上周一 = 本周一 - 7
    last_monday = this_monday - timedelta(days=7)
    last_sunday = last_monday + timedelta(days=6)
    return last_monday, last_sunday


def month_start(d: date) -> date:
    """返回该日期所在月份的1号。"""
    return d.replace(day=1)


def last_month_start(d: date) -> date:
    """返回上个月的1号。"""
    first = d.replace(day=1)
    last_day_of_last_month = first - timedelta(days=1)
    return last_day_of_last_month.replace(day=1)


def end_of_day(d: date) -> datetime:
    return datetime.combine(d, time(23, 59, 59))


def fmt_date(d: date | datetime, with_time: bool = False) -> str:
    if isinstance(d, datetime):
        return d.strftime("%Y-%m-%d %H:%M:%S") if with_time else d.strftime("%Y-%m-%d")
    return d.strftime("%Y-%m-%d")


def hours_before(d: datetime, hours: int) -> datetime:
    return d - timedelta(hours=hours)


if __name__ == "__main__":
    # 演示
    today = date.today()
    mon, sun = last_week_range(today)
    print(f"今天: {today} (weekday={today.weekday()})")
    print(f"上周一: {mon}")
    print(f"上周日: {sun}")
    print(f"本月1号: {month_start(today)}")
    print(f"上周日 -72h: {hours_before(end_of_day(sun), 72)}")
    print(f"上周日 -48h: {hours_before(end_of_day(sun), 48)}")
