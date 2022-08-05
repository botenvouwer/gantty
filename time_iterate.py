import calendar
import datetime
from enum import Enum
from itertools import islice
from math import floor


class TimeIteratorMode(Enum):
    MONTHS = 1
    WEEKS = 2
    DAYS = 7
    DAYS_NO_WEEKEND = 5


class TimeIterator:
    day_names = ('ma', 'di', 'wo', 'do', 'vr', 'za', 'zo')
    month_names = ('Januari', 'Februari', 'Maart', 'April', 'Mei', 'Juni', 'Juli', 'Augustus', 'September', 'Oktober', 'November', 'December')
    calendar = calendar.Calendar()
    __modes = ('months', 'weeks', 'days', 'days_without_weekend')

    def __init__(self, start_year, start_month, months_duration, mode: TimeIteratorMode = TimeIteratorMode.DAYS):
        self.start_year = start_year
        self.start_month = start_month
        self.months_duration = months_duration
        self.mode = mode
        self.end_year = start_year + floor((months_duration - start_month) / 12) + 1
        self.end_month = (months_duration - (12 - start_month)) % 12 - 1

    @staticmethod
    def get_month_name(month):
        return TimeIterator.month_names[month - 1]

    @staticmethod
    def get_day_name(day):
        return TimeIterator.day_names[day]

    def days_in_month(self, year, month):
        return len([x for x in self.calendar.itermonthdays2(year, month) if x[0] !=0 and x[1] < 5]) if self.mode == TimeIteratorMode.DAYS_NO_WEEKEND else calendar.monthrange(year, month)[1]

    def days_left_in_year(self, year, start_month=1, end_month=None):

        days_in_year = 0
        end = 13

        if end_month is not None:
            start_month = 1
            end = end_month + 1

        for d in range(start_month, end):
            dm = self.days_in_month(year, d)
            days_in_year += dm

        return days_in_year

    def iterate_full(self):
        ym_start = 12 * self.start_year + self.start_month - 1
        ym_end = 12 * self.end_year + self.end_month

        passed_months = 0
        i = 0
        ii = 0

        days_left_in_year = self.days_left_in_year(self.start_year, self.start_month)
        previous_year = self.start_year
        for ym in range(ym_start, ym_end):
            y, m = divmod(ym, 12)
            year = y
            month = m + 1
            month_name = self.get_month_name(month)
            passed_months += 1
            days_in_month = self.days_in_month(year, month)

            if previous_year < year:
                days_left_in_year = self.days_left_in_year(year, end_month=self.end_month) if self.end_year == year else self.days_left_in_year(year)
                previous_year = year

            if self.mode == TimeIteratorMode.MONTHS:
                yield passed_months, year, days_left_in_year, month, month_name, days_in_month
                continue

            for day_number in self.calendar.itermonthdays(year, month):
                if day_number == 0:
                    continue
                ii += 1

                week_number = datetime.date(year, month, day_number).isocalendar().week

                if self.mode == TimeIteratorMode.WEEKS:
                    # _last_week_of_year = datetime.date(year, 12, 31).isocalendar().week

                    if ii % 7 == 1:
                        i += 1
                        yield passed_months, year, days_left_in_year, month, month_name, days_in_month, i, week_number

                    continue

                if self.mode in (TimeIteratorMode.DAYS, TimeIteratorMode.DAYS_NO_WEEKEND):
                    day = calendar.weekday(year, month, day_number)

                    if self.mode == TimeIteratorMode.DAYS_NO_WEEKEND and day in (5, 6):
                        continue

                    i += 1
                    day_name = self.get_day_name(day)

                    yield passed_months, year, days_left_in_year, month, month_name, days_in_month, i, week_number, day_number, day, day_name

    def get_weeks_in_month(self, year, month):

        i = 0
        iter_t = self.calendar.itermonthdays(year, month)
        for day_number in iter_t:
            if day_number == 0:
                next(islice(iter_t, 5, 6), None)

            if day_number % 7 == 1:
                i += 1

        return i

    def __len__(self):
        return sum(1 for _ in self.iterate_full())

    def __iter__(self):
        return self.iterate_full()

    def __str__(self):
        s = f"TimeIterator for {self.start_year}-{self.start_month} until {self.end_year}-{self.end_month} taking {self.months_duration} months in {self.mode.name} mode"
        return s
