from ..Workbook import Workbook
import openpyxl
import xlsxwriter.workbook
import time


def test_benchmark():
    ROWS = 65000
    COLUMNS = 100
    wb = Workbook()
    testData = [[1] * COLUMNS] * ROWS
    stime = time.clock()
    ws = wb.new_sheet("Test 1", data=testData)
    wb.save("test.xlsx")
    print "pyexcelerate, %s, %s, %s" % (ROWS, COLUMNS, time.clock() - stime)
    
    stime = time.clock()
    wb = openpyxl.workbook.Workbook(optimized_write=True) 
    ws = wb.create_sheet()
    ws.title ="Test 1"
    for row in testData:
        ws.append(row)
    wb.save("test.xlsx")
    print "openpyxl, %s, %s, %s" % (ROWS, COLUMNS, time.clock() - stime)

    stime = time.clock()
    wb = xlsxwriter.workbook.Workbook('test.xlsx', {'reduce_memory': 1})
    ws = wb.add_worksheet()
    for row in range(ROWS):
        for col in range(COLUMNS):
            ws.write_number(row, col, 1)
    wb.close()
    print "xlsxwriter, %s, %s, %s" % (ROWS, COLUMNS, time.clock() - stime)
