from ..Workbook import Workbook
from ..Color import Color
from ..Font import Font
from ..Fill import Fill
from ..Style import Style
import time
import numpy
from datetime import datetime
from nose.tools import eq_
from .utils import get_output_path

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
	ws[2][1].value = "asdf"
	ws.range("A2", "B2").merge()
	eq_(ws[1][2].value, ws[1][1].value)
	ws[2][2].value = "qwer"
	eq_(ws[1][2].value, ws[1][1].value)
	ws[2][1].style.fill.background = Color(0, 255, 0, 0)
	wb.save(get_output_path("style-test.xlsx"))

def test_style_compression():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws.range("A1","C3").value = 1
	ws.range("A1","C1").style.font.bold = True
	ws.range("A2","C3").style.font.italic = True
	ws.range("A3","C3").style.fill.background = Color(255, 0, 0, 0)
	ws.range("C1","C3").style.font.strikethrough = True
	wb.save(get_output_path("style-compression-test.xlsx"))
	
def test_style_reference():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws[1][1].value = 1
	font = Font(bold=True, italic=True, underline=True, strikethrough=True)
	ws[1][1].style.font = font
	wb.save(get_output_path("style-reference-test.xlsx"))

def test_style_row():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws[1].style.fill.background = Color(255, 0, 0)
	ws[1][3].style.fill.background = Color(0, 255, 0)
	wb.save(get_output_path("style-row-test.xlsx"))

def test_and_or_xor():
	bolditalic = Font(bold=True, italic=True)
	italicunderline = Font(italic=True, underline=True)
	eq_(Font(italic=True), bolditalic & italicunderline)
	eq_(Font(bold=True, italic=True, underline=True), bolditalic | italicunderline)
	eq_(Font(bold=True, underline=True), bolditalic ^ italicunderline)
	
	fontstyle = Style(font=Font(bold=True))
	fillstyle = Style(fill=Fill(background=Color(255, 0, 0, 0)))
	eq_(Style(), fontstyle & fillstyle)
	eq_(Style(font=Font(bold=True), fill=Fill(background=Color(255, 0, 0, 0))), fontstyle | fillstyle)
	eq_(Style(font=Font(bold=True), fill=Fill(background=Color(255, 0, 0, 0))), fontstyle ^ fillstyle)