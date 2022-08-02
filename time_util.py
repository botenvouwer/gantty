import calendar


def days_left_in_year(year, start_month=1, end_month=None):

    days_in_year = 0

    if end_month is not None:
        for d in range(1, end_month + 1):
            dm = calendar.monthrange(year, d)[1]
            days_in_year += dm
        return days_in_year

    if start_month == 1:
        return 366 if calendar.isleap(year) else 365

    for d in range(start_month, 13):
        dm = calendar.monthrange(year, d)[1]
        days_in_year += dm

    return days_in_year




