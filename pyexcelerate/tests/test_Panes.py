from ..Panes import Panes
from ..Workbook import Workbook

from nose.tools import eq_, ok_, assert_raises
from pyexcelerate.Utility import to_unicode

import sys


def test_panes_boolean():
	ok_(not Panes())
	ok_(Panes(x=1))
	ok_(Panes(y=1))


def test_panes_output():
	eq_(to_unicode('<pane state="frozen" topLeftCell="A2" ySplit="1"/>'), Panes(y=1).get_xml())
	eq_(to_unicode('<pane topLeftCell="B1" xSplit="1"/>'), Panes(x=1, freeze=False).get_xml())
	eq_(to_unicode('<pane state="frozen" topLeftCell="C3" xSplit="2" ySplit="2"/>'), Panes(x=2, y=2).get_xml())


def test_worksheet_property():
	wb = Workbook()
	ws = wb.new_sheet("Test 1")
	ws.panes = Panes(x=1, y=2)
	eq_(ws.panes, Panes(x=1, y=2))

	if sys.version_info >= (2, 7):
		# assert_raises not supported before Python 2.7
		with assert_raises(TypeError):
			ws.panes = "banana"
