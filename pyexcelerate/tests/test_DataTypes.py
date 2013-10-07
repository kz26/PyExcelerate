from ..DataTypes import DataTypes
from nose.tools import eq_
from datetime import datetime
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
