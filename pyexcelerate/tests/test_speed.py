from .benchmark import run_pyexcelerate, run_xlsxwriter, run_openpyxl
from nose.tools import ok_

def test_vs_xlsxwriter():
	ours = run_pyexcelerate()
	theirs = run_xlsxwriter()
	ok_(ours / theirs <= 0.67, msg='PyExcelerate is too slow! Better Excelerate it some more!')
