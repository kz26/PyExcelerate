# -*- coding: utf-8 -*-

from __future__ import unicode_literals

from io import BytesIO
from xml.etree import ElementTree
from zipfile import ZipFile

import openpyxl
from nose.tools import eq_

from ..Font import Font
from ..Style import Style
from ..Workbook import Workbook


FORMULAS = [
    '=IF(A2<2,"yes","no")',
    '="left"&"right"',
    '=HYPERLINK("https://example.com/?a=1&b=2","link")',
    '="&amp;"',
    '="<tag/>"',
    '="</f><v>123</v><f>"',
    '=IF(A2>2,"yes","no")',
    '="café"&"<茶>"',
    "=SUM(A2:B2)",
]
NAMESPACE = {"s": "http://schemas.openxmlformats.org/spreadsheetml/2006/main"}


def assert_formula_roundtrip(formula, styled):
    workbook = Workbook()
    worksheet = workbook.new_sheet("Formulas", data=[[formula]])
    if styled:
        worksheet.cell("A1").style = Style(font=Font(bold=True))
    output = BytesIO()
    workbook.save(output)
    output.seek(0)
    with ZipFile(output) as archive:
        tree = ElementTree.fromstring(archive.read("xl/worksheets/sheet1.xml"))
    cell = tree.find("s:sheetData/s:row/s:c", NAMESPACE)
    eq_(len(cell), 1)
    eq_(cell.find("s:f", NAMESPACE).text, formula[1:])
    output.seek(0)
    loaded = openpyxl.load_workbook(output, data_only=False)
    eq_(loaded["Formulas"]["A1"].value, formula)
    eq_(loaded["Formulas"]["A1"].data_type, "f")
    eq_(loaded["Formulas"]["A1"].font.bold, styled)


def test_formula_xml_roundtrip():
    for formula in FORMULAS:
        for styled in (False, True):
            yield assert_formula_roundtrip, formula, styled
