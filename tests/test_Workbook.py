from Workbook import Workbook
from nose.tools import eq_
import cStringIO as StringIO

def test_get_xml_data():
    wb = Workbook()
    ws = wb.new_sheet("Test")
    ws[1][1].value = 1
    eq_(ws[1][1].value, 1)
    ws[1][3].value = 3
    eq_(ws[1][3].value, 3)
    ws_gxd = ws.get_xml_data().next()
    r_gxd = ws_gxd[1].get_xml_data()
    eq_(r_gxd.next(), ('A1', 1, 4))
    eq_(r_gxd.next(), ('C1', 3, 4))

def test_save():
    wb = Workbook()
    testData = [[1] * 50] * 1000
    for i in range(1):
        ws = wb.new_sheet("Test %s" % (i + 1), data=testData)
    wb.save("test.xlsx")

