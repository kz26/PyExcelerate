from ..Workbook import Workbook
from ..Color import Color
import time
from .utils import get_output_path
from random import randint

ROWS = 65000
COLUMNS = 100
BOLD = 1
ITALIC = 2
UNDERLINE = 4
RED_BG = 8
testData = [[1] * COLUMNS] * ROWS
formatData = [[1] * COLUMNS] * ROWS
def run_pyexcelerate():
	wb = Workbook()
	stime = time.clock()
	ws = wb.new_sheet('Test 1', data=testData)
	wb.save(get_output_path('test_pyexcelerate.xlsx'))
	elapsed = time.clock() - stime
	print("pyexcelerate, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed

def run_openpyxl():
	try:
		import openpyxl
	except ImportError:
		raise Exception('openpyxl not installled')
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
	try:
		import xlsxwriter.workbook
	except ImportError:
		raise Exception('XlsxWriter not installled')
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

def generate_format_data():
	for row in range(ROWS):
		for col in range(COLUMNS):
			formatData[row][col] = randint(1, (1 << 4) - 1)
	

def run_pyexcelerate_optimization():
	wb = Workbook()
	stime = time.clock()
	ws = wb.new_sheet('Test 1')
	for row in range(ROWS):
		for col in range(COLUMNS):
			ws.set_cell_value(row + 1, col + 1, 1)
			if formatData[row][col] & BOLD:
				ws[row + 1][col + 1].style.font.bold = True
			if formatData[row][col] & ITALIC:
				ws[row + 1][col + 1].style.font.italic = True
			if formatData[row][col] & UNDERLINE:
				ws[row + 1][col + 1].style.font.underline = True
			if formatData[row][col] & RED_BG:
				ws[row + 1][col + 1].fill.background = Color(255, 0, 0, 0)
	wb.save(get_output_path('test_pyexcelerate_opt.xlsx'))
	elapsed = time.clock() - stime
	print("pyexcelerate optimization, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed


def run_xlsxwriter_optimization():
	try:
		import xlsxwriter.workbook
	except ImportError:
		raise Exception('XlsxWriter not installled')
	stime = time.clock()
	wb = xlsxwriter.workbook.Workbook(get_output_path('test_xlsxwriter_opt.xlsx'), {'constant_memory': True})
	ws = wb.add_worksheet()

	bold = wb.add_format({'bold': True})
	italic = wb.add_format({'italic': True})
	underline = wb.add_format({'underline': 1})
	bg_color = wb.add_format({'bg_color': 'red'})

	for row in range(ROWS):
		for col in range(COLUMNS):
			if formatData[row][col] & BOLD:
				cell_format = bold
			if formatData[row][col] & ITALIC:
				cell_format = italic
			if formatData[row][col] & UNDERLINE:
				cell_format = underline
			if formatData[row][col] & RED_BG:
				cell_format = bg_color
			ws.write_number(row, col, 1, cell_format)
	wb.close()
	elapsed = time.clock() - stime
	print("xlsxwriter optimization, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed

def run_openpyxl_optimization():
	try:
		import openpyxl
	except ImportError:
		raise Exception('openpyxl not installled')
	stime = time.clock()
	wb = openpyxl.workbook.Workbook(optimized_write=True) 
	ws = wb.create_sheet()
	ws.title = 'Test 1'
	for col_idx in range(COLUMNS):
		col = get_column_letter(col_idx + 1)
		for row in range(ROWS):
			ws.cell('%s%s'%(col, row + 1)).value = 1
			if formatData[row][col] & BOLD:
				ws.cell('%s%s'%(col, row + 1)).style.font.bold = True
			if formatData[row][col] & ITALIC:
				ws.cell('%s%s'%(col, row + 1)).style.font.italic = True
			if formatData[row][col] & UNDERLINE:
				ws.cell('%s%s'%(col, row + 1)).style.font.underline = True
			if formatData[row][col] & RED_BG:
				ws.cell('%s%s'%(col, row + 1)).style.fill.fill_type = openpyxl.style.Fill.FILL_SOLID
				ws.cell('%s%s'%(col, row + 1)).style.fill.start_color = openpyxl.style.Color(Color.RED)
				ws.cell('%s%s'%(col, row + 1)).style.fill.end_color = openpyxl.style.Color(Color.RED)
			ws.write_number(row, col, 1, format)
	wb.save(get_output_path('test_openpyxl_opt.xlsx'))
	elapsed = time.clock() - stime
	print("openpyxl, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed
	
def test_all():
	run_pyexcelerate()
	run_xlsxwriter()
	run_openpyxl()
	ROWS = 650
	COLUMNS = 100
	generate_format_data()
	run_pyexcelerate_optimization()
	run_xlsxwriter_optimization()
	run_openpyxl_optimization()
