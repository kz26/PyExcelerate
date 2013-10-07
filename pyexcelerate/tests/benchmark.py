from ..Workbook import Workbook
import openpyxl
import xlsxwriter.workbook
import time
from .utils import get_output_path

ROWS = 65000
COLUMNS = 100
testData = [[1] * COLUMNS] * ROWS

def run_pyexcelerate():
	wb = Workbook()
	stime = time.clock()
	ws = wb.new_sheet('Test 1', data=testData)
	wb.save(get_output_path('test_pyexcelerate.xlsx'))
	elapsed = time.clock() - stime
	print("pyexcelerate, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed
	
def run_openpyxl():
	stime = time.clock()
	wb = openpyxl.workbook.Workbook(optimized_write=True) 
	ws = wb.create_sheet()
	ws.title = 'Test 1'
	for row in testData:
		ws.append(row)
	wb.save(get_output_path('test_openpyxl.xlsx'))
	elapsed = time.clock() - stime
	print("openpyxl, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed

def run_xlsxwriter():
	stime = time.clock()
	wb = xlsxwriter.workbook.Workbook(get_output_path('test_xlsxwriter.xlsx'), {'constant_memory': True})
	ws = wb.add_worksheet()
	for row in range(ROWS):
		for col in range(COLUMNS):
			ws.write_number(row, col, 1)
	wb.close()
	elapsed = time.clock() - stime
	print("xlsxwriter, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed

def test_all():
	run_pyexcelerate()
	run_xlsxwriter()
	run_openpyxl()
