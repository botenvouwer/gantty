start_year = 2022
start_month = 10
months = 3

print(f"s {start_year} {start_month}")

month = start_month - 1 + months
year = start_year + month // 12
month = month % 12
year = year if month > 0 else year - 1
month = month if month > 0 else 12


print(f"e {year} {month}")
