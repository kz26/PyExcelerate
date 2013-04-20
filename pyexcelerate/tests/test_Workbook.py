from ..Workbook import Workbook
import cStringIO as StringIO
import time

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
    print "%s, %s, %s" % (ROWS, COLUMNS, time.clock() - stime)

def test_formulas():
	wb = Workbook()
	ws = wb.new_sheet("test")
	ws[1][1].value = 1
	ws[1][2].value = 2
	ws[1][3].value = '=SUM(A1,B1)'
	wb.save("formula-test.xlsx")