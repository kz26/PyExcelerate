# -*- coding: utf-8 -*-

from ..DataTypes import DataTypes
from nose.tools import eq_
from datetime import datetime, date, time
from ..Workbook import Workbook
from .utils import get_output_path
from decimal import Decimal
import numpy

def test__get_type():
	eq_(DataTypes.get_type(15), DataTypes.NUMBER)
	eq_(DataTypes.get_type(15.0), DataTypes.NUMBER)
	eq_(DataTypes.get_type(Decimal('15.0')), DataTypes.NUMBER)
	eq_(DataTypes.get_type(float('inf')), DataTypes.NUMBER)
	eq_(DataTypes.get_type(float('nan')), DataTypes.NUMBER)
	eq_(DataTypes.get_type(numpy.inf), DataTypes.NUMBER)
	eq_(DataTypes.get_type(numpy.nan), DataTypes.NUMBER)
	eq_(DataTypes.get_type("test"), DataTypes.INLINE_STRING)
	eq_(DataTypes.get_type(datetime.now()), DataTypes.DATE)
	eq_(DataTypes.get_type(True), DataTypes.BOOLEAN)
	
def test_numpy():
	testData = numpy.ones((5, 5), dtype = int)
	wb = Workbook()
	ws = wb.new_sheet("Test 1", data=testData)
	ws[6][1].value = numpy.inf
	ws[6][2].value = numpy.nan
	ws[7][1].value = float('inf')
	ws[7][2].value = float('nan')
	eq_(ws[1][1].value, 1)
	eq_(DataTypes.get_type(ws[1][1].value), DataTypes.NUMBER)
	wb.save(get_output_path("numpy-test.xlsx"))

def test_ampersand_escaping():
	testData = [["http://example.com/?one=1&two=2"]]
	wb = Workbook()
	ws = wb.new_sheet("Test 1", data=testData)
	data = list(ws.get_xml_data())
	assert "http://example.com/?one=1&amp;two=2" in data[0][1][0]

def test_to_excel_date():
	eq_(DataTypes.to_excel_date(datetime(1900, 1, 1, 0, 0, 0)), 1.0)
	eq_(DataTypes.to_excel_date(datetime(1900, 1, 1, 12, 0, 0)), 1.5)
	eq_(DataTypes.to_excel_date(datetime(1900, 1, 1, 12, 0, 0)), 1.5)
	eq_(DataTypes.to_excel_date(datetime(2013, 5, 10, 6, 0, 0)), 41404.25)
	eq_(DataTypes.to_excel_date(date(1900, 1, 1)), 1.0)
	eq_(DataTypes.to_excel_date(date(2013, 5, 10)), 41404.0)
	eq_(DataTypes.to_excel_date(time(6, 0, 0)), 0.25)
	# check excel's improper handling of leap year
	eq_(DataTypes.to_excel_date(datetime(1900, 2, 28, 0, 0, 0)), 59.0)
	eq_(DataTypes.to_excel_date(datetime(1900, 3, 1, 0, 0, 0)), 61.0)

def test_unicode_str():
	wb = Workbook()
	ws = wb.new_sheet("Unicode test")
	ws[1][1].value = 'ಠ_ಠ'
	ws[1][2].value = '(╯°□°）╯︵ ┻━┻'
	ws[1][3].value = 'ᶘ ᵒᴥᵒᶅ'
	ws[1][4].value = 'الله أكبر'
	wb.save(get_output_path("unicode-test.xlsx"))
