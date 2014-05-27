from ..Workbook import Workbook
from ..Color import Color
from ..Font import Font
from ..Fill import Fill
from ..Style import Style
from ..Alignment import Alignment
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
    ws[1][1].style.font.color = Color(255, 0, 255)
    ws[1][1].style.fill.background = Color(0, 255, 0)
    ws[1][2].style.fill.background = Color(255, 255, 0)
    ws[2][1].value = "asdf"
    ws.range("A2", "B2").merge()
    eq_(ws[1][2].value, ws[1][1].value)
    ws[2][2].value = "qwer"
    eq_(ws[1][2].value, ws[1][1].value)
    ws[2][1].style.fill.background = Color(0, 255, 0)
    ws[1][1].style.alignment.vertical = 'top'
    ws[1][1].style.alignment.horizontal = 'right'
    ws[1][1].style.alignment.rotation = 90
    ws[3][3].style.borders.top.color = Color(255, 0, 0)
    ws[3][3].style.borders.left.color = Color(0, 255, 0)
    ws[3][4].style.borders.right.style = '-.'
    wb.save(get_output_path("style-test.xlsx"))


def test_style_compression():
    wb = Workbook()
    ws = wb.new_sheet("test")
    ws.range("A1", "C3").value = 1
    ws.range("A1", "C1").style.font.bold = True
    ws.range("A2", "C3").style.font.italic = True
    ws.range("A3", "C3").style.fill.background = Color(255, 0, 0)
    ws.range("C1", "C3").style.font.strikethrough = True
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


def test_style_row_col():
    wb = Workbook()
    ws = wb.new_sheet("test")
    ws.range("A1", "D4").value = 'sdfgs5b56seb6se56bse5jsdfljg'
    eq_(Style(), ws.get_row_style(1))
    eq_(Style(), ws.get_col_style(1))
    ws.set_row_style(1, Style(size=-1))
    ws.set_row_style(2, Style(size=0))
    ws.set_row_style(
        3, Style(size=100, fill=Fill(background=Color(0, 255, 0, 0))))
    ws.set_col_style(1, Style(size=-1))
    ws.set_col_style(2, Style(size=0))
    ws.set_col_style(
        3, Style(size=100, fill=Fill(background=Color(255, 0, 0, 0))))
    wb.save(get_output_path("style-auto-row-col-test.xlsx"))


def test_and_or_xor():
    bolditalic = Font(bold=True, italic=True)
    italicunderline = Font(italic=True, underline=True)
    eq_(Font(italic=True), bolditalic & italicunderline)
    eq_(Font(bold=True, italic=True, underline=True),
        bolditalic | italicunderline)
    eq_(Font(bold=True, underline=True), bolditalic ^ italicunderline)

    fontstyle = Style(font=Font(bold=True))
    fillstyle = Style(fill=Fill(background=Color(255, 0, 0, 0)))
    eq_(Style(), fontstyle & fillstyle)
    eq_(Style(font=Font(bold=True), fill=Fill(background=Color(255, 0, 0, 0))),
        fontstyle | fillstyle)
    eq_(Style(font=Font(bold=True), fill=Fill(background=Color(255, 0, 0, 0))),
        fontstyle ^ fillstyle)

    leftstyle = Style(alignment=Alignment('right', 'top'))
    bottomstyle = Style(alignment=Alignment(vertical='top', rotation=15))
    eq_(Style(alignment=Alignment('right', 'top', 15)),
        leftstyle | bottomstyle)
    eq_(Style(alignment=Alignment(vertical='top')), leftstyle & bottomstyle)
    eq_(Style(alignment=Alignment('right', rotation=15)),
        leftstyle ^ bottomstyle)


def test_str_():
    font = Font(bold=True, italic=True, underline=True, strikethrough=True)
    eq_(font.__repr__(), "<Font: Calibri, 11pt b i u s>")