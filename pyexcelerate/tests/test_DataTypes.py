from ..DataTypes import DataTypes
from nose.tools import eq_
from datetime import datetime, date, time
from ..Workbook import Workbook
from .utils import get_output_path
import numpy

def test__to_enumeration_value():
	eq_(DataTypes.to_enumeration_value(DataTypes.BOOLEAN), "b")
	eq_(DataTypes.to_enumeration_value(DataTypes.DATE), "d")
	eq_(DataTypes.to_enumeration_value(DataTypes.ERROR), "e")
	eq_(DataTypes.to_enumeration_value(DataTypes.INLINE_STRING), "inlineStr")
	eq_(DataTypes.to_enumeration_value(DataTypes.NUMBER), "n")
	eq_(DataTypes.to_enumeration_value(DataTypes.SHARED_STRING), "s")
	eq_(DataTypes.to_enumeration_value(DataTypes.STRING), "str")

def test__get_type():
	eq_(DataTypes.get_type(15), DataTypes.NUMBER)
	eq_(DataTypes.get_type(15.0), DataTypes.NUMBER)
	eq_(DataTypes.get_type("test"), DataTypes.INLINE_STRING)
	eq_(DataTypes.get_type(datetime.now()), DataTypes.DATE)
	
def test_numpy():
	testData = numpy.ones((5, 5), dtype = int)
	wb = Workbook()
	ws = wb.new_sheet("Test 1", data=testData)
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
