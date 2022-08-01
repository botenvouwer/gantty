from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Font
from openpyxl.utils import get_column_letter

name = "gantt"
path = Path.home() / (name + '.xlsx')

italic = Font(italic=True)
bold = Font(bold=True)

wb = Workbook()
ws1 = wb.active
ws1.title = name

ws1_headers1 = ('', '', 'sdsd', '', '', '', '', '', 'dfdfdfdf', '', '')
ws1_headers2 = ('', 'provincie', 'aansluitingen', 'opwek ans', 'vermogen', 'diff', 'diff', 'diff', 'aansluitingen', 'opwek ans', 'vermogen')
ws1.append(ws1_headers1)
ws1.append(ws1_headers2)

ws1.column_dimensions['B'].width = 18
for l in ('C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K'):
    ws1.column_dimensions[l].width = 13

ws1['B1'].font = italic
ws1['I1'].font = italic

for c in ws1['A2:K2'][0]:
    c.font = bold

for i in ['', '', '']:
    ws1.append(i)

wb.save(filename=path)
