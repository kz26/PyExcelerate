from ..Workbook import Workbook
import time
import numpy
from datetime import datetime
from nose.tools import eq_

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
    wb.save("test.xlsx")
    print("%s, %s, %s" % (ROWS, COLUMNS, time.clock() - stime))

def test_formulas():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws[1][1].value = 1
	ws[1][2].value = 2
	ws[1][3].value = '=SUM(A1,B1)'
	ws[1][4].value = datetime.now()
	wb.save("formula-test.xlsx")
	eq_(1,2)
	
def test_merge():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws[1][1].value = "asdf"
	ws.range("A1", "B1").merge()
	eq_(ws[1][2].value, ws[1][1].value)
	ws[1][2].value = "qwer"
	eq_(ws[1][2].value, ws[1][1].value)
	wb.save("merge-test.xlsx")
	
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
	wb.save("numpy-range-test.xlsx")
