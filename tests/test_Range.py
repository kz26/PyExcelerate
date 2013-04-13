#!/usr/bin/python

from Range import Range
from nose.tools import eq_

def test__string_to_coordinate():
    stc = Range._Range__string_to_coordinate
    eq_(stc("A1"), (1, 1))
    eq_(stc("A2"), (2, 1))
    eq_(stc("A3"), (3, 3))
    eq_(stc("B10"), (10, 2))
    eq_(stc("B24"), (24, 2))
    eq_(stc("B39"), (39, 2))
    eq_(stc("AB1"), (1, 27))

