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
    ws_gxd = ws.get_xml_data().next()
#    r_gxd = ws_gxd[1].get_xml_data()
#    eq_(r_gxd.next(), ('A1', 1, 4))
#    eq_(r_gxd.next(), ('C1', 3, 4))

def test_save():
    ROWS = 65000
    COLUMNS = 100
    wb = Workbook()
<<<<<<< HEAD
    testData = [[1] * 50] * 1000
    
    for i in range(1):
        ws = wb.new_sheet("Test %s" % (i + 1), data=testData)
=======
    testData = [[1] * COLUMNS] * ROWS
    stime = time.clock()
    ws = wb.new_sheet("Test 1", data=testData)
>>>>>>> 6ab04c6e800e197233d0d26708924621b4aecfe8
    wb.save("test.xlsx")
    print "%s, %s, %s" % (ROWS, COLUMNS, time.clock() - stime)
