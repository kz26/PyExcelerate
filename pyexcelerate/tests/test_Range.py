from ..Workbook import Workbook
from ..Range import Range
from nose.tools import eq_

def test__string_to_coordinate():
    stc = Range.string_to_coordinate
    eq_(stc("A1"), (1, 1))
    eq_(stc("A2"), (2, 1))
    eq_(stc("A3"), (3, 1))
    eq_(stc("Z1"), (1, 26))
    eq_(stc("B10"), (10, 2))
    eq_(stc("B24"), (24, 2))
    eq_(stc("B39"), (39, 2))
    eq_(stc("AA1"), (1, 27))
    eq_(stc("AB1"), (1, 28))

def test__coordinate_to_string():
    cts = Range.coordinate_to_string
    eq_(cts((1, 1)), "A1")
    eq_(cts((2, 1)), "A2")
    eq_(cts((3, 1)), "A3")
    eq_(cts((1, 26)), "Z1")
    eq_(cts((1, 2)), "B1")
    eq_(cts((10, 2)), "B10")
    eq_(cts((24, 2)), "B24")
    eq_(cts((39, 2)), "B39")
    eq_(cts((1, 27)), "AA1")
    eq_(cts((1, 28)), "AB1")

def test_merge():
     wb = Workbook()
     ws = wb.new_sheet("Test")
     r1 =  Range("A1", "A5", ws)
     r1.merge()
     r2 =  Range("B1", "B5", ws)
     r2.merge()
     eq_(len(ws.merges), 2)

def test_horizontal_intersection():
    wb = Workbook()
    ws = wb.new_sheet("Test")
    r1 =  Range("A1", "A3", ws)
    r2 =  Range("A2", "A4", ws)
    eq_(r1.intersects(r2), True)
    eq_(r1.intersection(r2), Range("A2", "A3", ws))

def test_vertical_intersection():
    wb = Workbook()
    ws = wb.new_sheet("Test")
    r1 =  Range("A1", "C1", ws)
    r2 =  Range("B1", "D1", ws)
    eq_(r1.intersects(r2), True)
    eq_(r1.intersection(r2), Range("B1", "C1", ws))

def test_rectangular_intersection():
    wb = Workbook()
    ws = wb.new_sheet("Test")
    r1 =  Range("A1", "C3", ws)
    r2 =  Range("B2", "D4", ws)
    eq_(r1.intersects(r2), True)
    eq_(r1.intersection(r2), Range("B2", "C3", ws))

def test_no_intersection():
    wb = Workbook()
    ws = wb.new_sheet("Test")
    r1 =  Range("A1", "B2", ws)
    r2 =  Range("C3", "D4", ws)
    eq_(r1.intersects(r2), False)
    eq_(r1.intersection(r2), None)

def test_range_equal_to_none():
    wb = Workbook()
    ws = wb.new_sheet("Test")
    r1 =  Range("A1", "C3", ws)
    r2 =  Range("B2", "D4", ws)
    eq_(r1.intersection(r2) == None, False)

def test_range_equal_to_itself():
    wb = Workbook()
    ws = wb.new_sheet("Test")
    r1 =  Range("A1", "C3", ws)
    eq_(r1 == r1, True)

"""
def test_get_xml_data():
    wb = Workbook()
    ws = wb.new_sheet("Test")
    ws[1][1].value = 1
    eq_(ws[1][1].value, 1) 
    ws[1][3].value = 3
    eq_(ws[1][3].value, 3)
    gxd = ws[1].get_xml_data()
    eq_(gxd.next(), ('A1', 1, 4))
    eq_(gxd.next(), ('C1', 3, 4))
"""
