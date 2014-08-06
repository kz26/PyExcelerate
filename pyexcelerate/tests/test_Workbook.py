from ..Workbook import Workbook
from ..Color import Color
import time
import numpy
import nose
import os
from datetime import datetime
from nose.tools import eq_, raises
from .utils import get_output_path

def test_get_xml_data():
    wb = Workbook()
    ws = wb.new_sheet("Test")
    ws[1][1].value = 1
    eq_(ws[1][1].value, 1)
    ws[1][3].value = 3
    eq_(ws[1][3].value, 3)
    
def test_save():
    ROWS = 65
    COLUMNS = 100
    wb = Workbook()
    testData = [[1] * COLUMNS] * ROWS
    stime = time.clock()
    ws = wb.new_sheet("Test 1", data=testData)
    wb.save(get_output_path("test.xlsx"))
    #print("%s, %s, %s" % (ROWS, COLUMNS, time.clock() - stime))

def test_formulas():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws[1]['A'].value = 1
	ws[1][2].value = 2
	ws[1][3].value = '=SUM(A1,B1)'
	ws[1][4].value = datetime.now()
	ws[1][5].value = datetime(1900,1,1,1,0,0)
	ws[1][6].value = True
	wb.save(get_output_path("formula-test.xlsx"))
	
def test_merge():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws[1][1].value = "asdf"
	ws.range("A1", "B1").merge()
	eq_(ws[1][2].value, ws[1][1].value)
	ws[1][2].value = "qwer"
	eq_(ws[1][2].value, ws[1][1].value)
	wb.save(get_output_path("merge-test.xlsx"))
	
def test_cell():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws.cell("C3").value = "test"
	eq_(ws[3][3].value, "test")

def test_range():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws.range("B2", "D3").value = [[1, 2, 3], [4, 5, 6]]
	eq_(ws[2][2].value, 1)
	eq_(ws[2][3].value, 2)
	eq_(ws[2][4].value, 3)
	eq_(ws[3][2].value, 4)
	eq_(ws[3][3].value, 5)
	eq_(ws[3][4].value, 6)
	
def test_numpy_range():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws.range("A1", "GN13").value = numpy.zeros((13,196))
	wb.save(get_output_path("numpy-range-test.xlsx"))

def test_none():
	testData = [[1,2,None]]
	wb = Workbook()
	ws = wb.new_sheet("Test 1", data=testData)
	ws[1][1].style.font.bold = True
	wb.save(get_output_path("none-test.xlsx"))
	
@raises(Exception)
def test_name_length():
    wb = Workbook()
    ws = wb.new_sheet('12345678901234567890123456789012')

def test_name_length_force():
    wb = Workbook()
    ws = wb.new_sheet('12345678901234567890123456789012', force_name=True)
    
@raises(Exception)
def test_name_duplicate():
    wb = Workbook()
    ws = wb.new_sheet('1234')
    ws = wb.new_sheet('1234')

def test_name_no_duplicate():
    wb = Workbook()
    ws = wb.new_sheet('1234')
    ws = wb.new_sheet('12356')
	

def test_number_precision():
	try:
		import xlrd
	except ImportError:
		raise nose.SkipTest('xlrd not installed')

	filename = get_output_path('precision.xlsx')
	sheetname = 'Sheet1'

	nums = [
		1,
		1.2,
		1.23,
		1.234,
		1.2345,
		1.23456,
		1.234567,
		1.2345678,
		1.23456789,
		1.234567890,
		1.2345678901,
		1.23456789012,
		1.234567890123,
		1.2345678901234,
		1.23456789012345,
	]

	write_workbook = Workbook()
	write_worksheet = write_workbook.new_sheet(sheetname)

	for index, value in enumerate(nums):
		write_worksheet[index + 1][1].value = value

	write_workbook.save(filename)

	read_workbook = xlrd.open_workbook(filename)
	read_worksheet = read_workbook.sheet_by_name(sheetname)

	for row_num in range(len(nums)):
		expected = nums[row_num]
		got = read_worksheet.cell(row_num, 0).value

	if os.path.exists(filename):
		os.remove(filename)

#def test_column_select():
#	wb = Workbook()
#	ws = wb.new_sheet("Test")
#	print(ws[1:3])
#	ws[1:3][1].style.fill.background = Color(255, 0, 0)
