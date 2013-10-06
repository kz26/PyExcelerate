from ..Workbook import Workbook
from ..Color import Color
import time
import numpy
from datetime import datetime
from nose.tools import eq_

def test_style():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws[1][1].value = 1
	ws[1][2].value = 1
	ws[1][3].value = 1
	ws[1][1].style.font.bold = True
	ws[1][2].style.font.italic = True
	ws[1][3].style.font.underline = True
	ws[1][1].style.font.strikethrough = True
	ws[1][1].style.fill.background = Color(0, 255, 0, 0)
	ws[1][2].style.fill.background = Color(255, 255, 0, 0)
	wb.save("style-test.xlsx")