import datetime


import calendar

from generator.time_iterator import TimeIterator, TimeIteratorMode

start_year = 2024
start_month = 2
months_duration = 200

# t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.MONTHS)
t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.WEEKS)
# t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.DAYS)
# t = TimeIterator(start_year, start_month, months_duration, mode=TimeIteratorMode.DAYS_NO_WEEKEND)
#
# print(t)
# print(len(t))

print(t.get_weeks_in_month(2024, 2))

previous_month = None
for e in t:
    month = e[3]

    if month != previous_month:
        # Print message indicating a new month starts
        print()
    year = e[1]
    month = e[3]

    print(f'{e} {t.get_weeks_in_month(year, month)}')

    previous_month = month

# variable t is iterable
# I want to walk over it and print the yielded values


# Print "Hello, World!"

