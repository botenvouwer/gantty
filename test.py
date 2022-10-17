import datetime

from time_iterate import TimeIterator, TimeIteratorMode
import calendar

start_year = 2022
start_month = 8
months_duration = 12

# t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.MONTHS)
t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.WEEKS)
# t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.DAYS)
# t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.DAYS_NO_WEEKEND)

print(t)
print(len(t))
print(t.get_weeks_in_month(start_year, 9))

for e in t:
    print(e)
