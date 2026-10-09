"""Reference model of the overtime form's Day / Public holiday / Travel / Work columns.

Python twin of the `Oracle` function in vba/OvertimeStressTest.bas. It re-implements
the rules in minutes instead of copying the Excel formulas, so the two can check
each other. `cross_check_libreoffice.py` compares it with the real formulas.
"""
import datetime as dt

OFFICE_START = 8 * 60        # 'Engineer Name'!Q1
OFFICE_END = 17 * 60 + 30    # 'Engineer Name'!Q2


def _segment(s, e, special, q1, q2):
    """Overtime minutes in [s, e]; weekdays exclude the start day's office hours."""
    if e <= s:
        return 0
    if special:
        return e - s
    off = 1440 if s >= 1440 else 0
    overlap = max(0, min(e, q2 + off) - max(s, q1 + off))
    return (e - s) - overlap


def oracle(date, times, holidays=(), q1=OFFICE_START, q2=OFFICE_END):
    """date: datetime.date or None. times: [From, Until/From, Until/From, Until] in
    minutes after midnight, None when blank. Returns (day, ph, travel_h, work_h),
    with '' where the form shows a blank cell."""
    if date is None:
        return (" ", "", "", "")
    ph = "Y" if date in holidays else "N"
    day = date.strftime("%a")
    if all(t is None for t in times):
        return (day, ph, "", "")
    d, e, f, g = times
    special = date.weekday() >= 5 or ph == "Y"

    # one timeline: each later time is at or after the one before it
    prev = d if d is not None else 0
    u = [prev, None, None, None]
    for i in (1, 2, 3):
        t = times[i]
        if t is None:
            continue
        while t < prev:
            t += 1440
        u[i] = prev = t

    seg = lambda a, b: _segment(a, b, special, q1, q2)
    travel = 0
    if d is not None and e is not None:
        travel += seg(u[0], u[1])
    if f is not None and g is not None:
        travel += seg(u[2], u[3])
    if d is not None and g is not None and e is None and f is None:
        travel += seg(u[0], u[3])

    work = 0
    if e is not None:
        if f is not None:
            work = seg(u[1], u[2])
        elif g is not None:
            work = seg(u[1], u[3])
    elif d is not None and f is not None:
        work = seg(u[0], u[2])

    return (day, ph, round(travel / 60 + 1e-9, 2), round(work / 60 + 1e-9, 2))


if __name__ == "__main__":
    sat = dt.date(2026, 9, 19)
    print(oracle(sat, [0, 60, 390, 570]))   # ('Sat', 'N', 4.0, 5.5)
