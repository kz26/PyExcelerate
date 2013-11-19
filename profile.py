import time
from pyexcelerate import Workbook
from pyexcelerate.Style import Style
from pyexcelerate.Font import Font
from pyexcelerate.Fill import Fill
from pyexcelerate.Color import Color
from random import randint

ROWS = 1000
COLUMNS = 100
BOLD = 1
ITALIC = 2
UNDERLINE = 4
RED_BG = 8

formatData = [[1] * COLUMNS] * ROWS
for row in range(ROWS):
	for col in range(COLUMNS):
		formatData[row][col] = randint(1, (1 << 4) - 1)
wb = Workbook()
stime = time.clock()
ws = wb.new_sheet('Test 1')
bold = Style(font=Font(bold=True))
italic = Style(font=Font(italic=True))
underline = Style(font=Font(underline=True))
red = Style(fill=Fill(background=Color(255,0,0,0)))
for row in range(ROWS):
	for col in range(COLUMNS):
		ws.set_cell_value(row + 1, col + 1, 1)
		style = Style()
		if formatData[row][col] & BOLD:
			style.font.bold = True
		if formatData[row][col] & ITALIC:
			style.font.italic = True
		if formatData[row][col] & UNDERLINE:
			style.font.underline = True
		if formatData[row][col] & RED_BG:
			style.fill.background = Color(255, 0, 0)
		ws.set_cell_style(row + 1, col + 1, style)
wb.save('test_pyexcelerate_style_fastest.xlsx')
elapsed = time.clock() - stime
print("pyexcelerate style fastest, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
