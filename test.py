import datetime


import calendar

from generator.time_iterator import TimeIterator, TimeIteratorMode

start_year = 2024
start_month = 1
months_duration = 24

# t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.MONTHS)
t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.WEEKS)
# t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.DAYS)
# t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.DAYS_NO_WEEKEND)

print(t)
print(len(t))
print(t.get_weeks_in_month(start_year, 9))

for e in t:
    print(e)

# variable t is iterable
# I want to walk over it and print the yielded values


# Print "Hello, World!"

