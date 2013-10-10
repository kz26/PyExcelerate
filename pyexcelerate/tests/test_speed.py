from .benchmark import run_pyexcelerate_value_fastest, run_xlsxwriter_value, run_openpyxl
from nose.tools import ok_
import nose

def test_vs_xlsxwriter():
	raise nose.SkipTest('Skipping speed test')
	ours = run_pyexcelerate_value_fastest()
	theirs = run_xlsxwriter_value()
	ok_(ours / theirs <= 0.67, msg='PyExcelerate is too slow! Better Excelerate it some more!')
