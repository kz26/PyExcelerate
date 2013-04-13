from Range import Range
from nose.tools import eq_

def test__string_to_coordinate():
    stc = Range._Range__string_to_coordinate
    eq_(stc("A1"), (1, 1))
    eq_(stc("A2"), (2, 1))
    eq_(stc("A3"), (3, 1))
    eq_(stc("B10"), (10, 2))
    eq_(stc("B24"), (24, 2))
    eq_(stc("B39"), (39, 2))
    eq_(stc("AA1"), (1, 27))
    eq_(stc("AB1"), (1, 28))

def test__cordinate_to_string():
    cts = Range._Range__coordinate_to_string
    eq_(cts((1, 1)), "A1")
    eq_(cts((2, 1)), "A2")
    eq_(cts((3, 1)), "A3")
    eq_(cts((10, 2)), "B10")
    eq_(cts((24, 2)), "B24")
    eq_(cts((39, 2)), "B39")
    eq_(cts((1, 27)), "AA1")
    eq_(cts((1, 28)), "AB1")

