from ..Workbook import Workbook
from ..Color import Color
from ..Style import Style
from ..Font import Font
from ..Fill import Fill
import time
from .utils import get_output_path
from random import randint

ROWS = 650
COLUMNS = 100
BOLD = 1
ITALIC = 2
UNDERLINE = 4
RED_BG = 8
testData = [[1] * COLUMNS] * ROWS
formatData = [[1] * COLUMNS] * ROWS
def run_pyexcelerate_value_fastest():
	wb = Workbook()
	stime = time.clock()
	ws = wb.new_sheet('Test 1', data=testData)
	wb.save(get_output_path('test_pyexcelerate_value_fastest.xlsx'))
	elapsed = time.clock() - stime
	print("pyexcelerate value fastest, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed

def run_pyexcelerate_value_faster():
	wb = Workbook()
	stime = time.clock()
	ws = wb.new_sheet('Test 1')
	for row in range(ROWS):
		for col in range(COLUMNS):
			ws.set_cell_value(row + 1, col + 1, 1)
	wb.save(get_output_path('test_pyexcelerate_value_faster.xlsx'))
	elapsed = time.clock() - stime
	print("pyexcelerate value faster, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed
	

def run_pyexcelerate_value_fast():
	wb = Workbook()
	stime = time.clock()
	ws = wb.new_sheet('Test 1')
	for row in range(ROWS):
		for col in range(COLUMNS):
			ws[row + 1][col + 1].value = 1
	wb.save(get_output_path('test_pyexcelerate_value_fast.xlsx'))
	elapsed = time.clock() - stime
	print("pyexcelerate value fast, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
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

def run_xlsxwriter_value():
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
	print("xlsxwriter value, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed
	
def generate_format_data():
	for row in range(ROWS):
		for col in range(COLUMNS):
			formatData[row][col] = randint(1, (1 << 4) - 1)
	
def run_pyexcelerate_style_fastest():
	wb = Workbook()
	stime = time.clock()
	ws = wb.new_sheet('Test 1')
	bold = Style(font=Font(bold=True))
	italic = Style(font=Font(italic=True))
	underline = Style(font=Font(underline=True))
	red = Style(fill=Fill(background=Color(255,0,0,0)))
	for row in range(ROWS):
		for col in range(COLUMNS):
			ws.set_cell_value(row + 1, col + 1, 1)
			style = Style()
			if formatData[row][col] & BOLD:
				style.font.bold = True
			if formatData[row][col] & ITALIC:
				style.font.italic = True
			if formatData[row][col] & UNDERLINE:
				style.font.underline = True
			if formatData[row][col] & RED_BG:
				style.fill.background = Color(255, 0, 0)
			ws.set_cell_style(row + 1, col + 1, style)
	wb.save(get_output_path('test_pyexcelerate_style_fastest.xlsx'))
	elapsed = time.clock() - stime
	print("pyexcelerate style fastest, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed
	
def run_pyexcelerate_style_faster():
	wb = Workbook()
	stime = time.clock()
	ws = wb.new_sheet('Test 1')
	for row in range(ROWS):
		for col in range(COLUMNS):
			ws.set_cell_value(row + 1, col + 1, 1)
			if formatData[row][col] & BOLD:
				ws.get_cell_style(row + 1, col + 1).font.bold = True
			if formatData[row][col] & ITALIC:
				ws.get_cell_style(row + 1, col + 1).font.italic = True
			if formatData[row][col] & UNDERLINE:
				ws.get_cell_style(row + 1, col + 1).font.underline = True
			if formatData[row][col] & RED_BG:
				ws.get_cell_style(row + 1, col + 1).fill.background = Color(255, 0, 0)
	wb.save(get_output_path('test_pyexcelerate_style_faster.xlsx'))
	elapsed = time.clock() - stime
	print("pyexcelerate style faster, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed
	
def run_pyexcelerate_style_fast():
	wb = Workbook()
	stime = time.clock()
	ws = wb.new_sheet('Test 1')
	for row in range(ROWS):
		for col in range(COLUMNS):
			ws[row + 1][col + 1].value = 1
			if formatData[row][col] & BOLD:
				ws[row + 1][col + 1].style.font.bold = True
			if formatData[row][col] & ITALIC:
				ws[row + 1][col + 1].style.font.italic = True
			if formatData[row][col] & UNDERLINE:
				ws[row + 1][col + 1].style.font.underline = True
			if formatData[row][col] & RED_BG:
				ws[row + 1][col + 1].style.fill.background = Color(255, 0, 0, 0)
	wb.save(get_output_path('test_pyexcelerate_style_fast.xlsx'))
	elapsed = time.clock() - stime
	print("pyexcelerate style fast, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed
	
def run_pyexcelerate_style_cheating():
	wb = Workbook()
	stime = time.clock()
	ws = wb.new_sheet('Test 1')

	cell_formats = []

	for i in range(16):
		cell_format = Style()
		if i & BOLD:
			cell_format.font.bold = True
		if i & ITALIC:
			cell_format.font.italic = True
		if i & UNDERLINE:
			cell_format.font.underline = True
		if i & RED_BG:
			cell_format.fill.background = Color(255, 0, 0)
		cell_formats.append(cell_format)

	for row in range(ROWS):
		for col in range(COLUMNS):
			ws.set_cell_value(row + 1, col + 1, 1)
			ws.set_cell_style(row + 1, col + 1, cell_formats[formatData[row][col]])
	wb.save(get_output_path('test_pyexcelerate_style_fastest.xlsx'))
	elapsed = time.clock() - stime
	print("pyexcelerate style cheating, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed

def run_xlsxwriter_style_cheating():
	try:
		import xlsxwriter.workbook
	except ImportError:
		raise Exception('XlsxWriter not installled')
	stime = time.clock()
	wb = xlsxwriter.workbook.Workbook(get_output_path('test_xlsxwriter_style.xlsx'), {'constant_memory': True})
	ws = wb.add_worksheet()
	
	cell_formats = []

	for i in range(16):
		cell_format = wb.add_format()
		if i & BOLD:
			cell_format.set_bold()
		if i & ITALIC:
			cell_format.set_italic()
		if i & UNDERLINE:
			cell_format.set_underline()
		if i & RED_BG:
			cell_format.set_bg_color('red')
		cell_formats.append(cell_format)

	for row in range(ROWS):
		for col in range(COLUMNS):
			ws.write_number(row, col, 1, cell_formats[formatData[row][col]]) 
	wb.close()
	elapsed = time.clock() - stime
	print("xlsxwriter style cheating, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
	return elapsed
	
def run_xlsxwriter_style():
	try:
		import xlsxwriter.workbook
	except ImportError:
		raise Exception('XlsxWriter not installled')
	stime = time.clock()
	wb = xlsxwriter.workbook.Workbook(get_output_path('test_xlsxwriter_style.xlsx'), {'constant_memory': True})
	ws = wb.add_worksheet()
	
	for row in range(ROWS):
		for col in range(COLUMNS):
			format = wb.add_format()
			if formatData[row][col] & BOLD:
				format.set_bold()
			if formatData[row][col] & ITALIC:
				format.set_italic()
			if formatData[row][col] & UNDERLINE:
				format.set_underline()
			if formatData[row][col] & RED_BG:
				format.set_bg_color('red')
			ws.write_number(row, col, 1, format) 
	wb.close()
	elapsed = time.clock() - stime
	print("xlsxwriter style, %s, %s, %s" % (ROWS, COLUMNS, elapsed))
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
	run_pyexcelerate_value_fastest()
	run_pyexcelerate_value_faster()
	run_pyexcelerate_value_fast()
	run_xlsxwriter_value()
	#run_openpyxl()
	generate_format_data()
	run_pyexcelerate_style_cheating()
	run_pyexcelerate_style_fastest()
	run_pyexcelerate_style_faster()
	run_pyexcelerate_style_fast()
	run_xlsxwriter_style_cheating()
	run_xlsxwriter_style()
	#run_openpyxl_optimization()