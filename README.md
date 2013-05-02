# PyExcelerate

Accelerated Excel XLSX writing library for Python 2/3

* Current version: 0.2.4
* Authors: [Kevin Wang](https://github.com/kevmo314) and [Kevin Zhang](https://github.com/whitehat2k9)
* License: Simplified BSD License
* [Source repository](https://github.com/whitehat2k9/PyExcelerate)
* [PyPI page](https://pypi.python.org/pypi/PyExcelerate)

## Description
PyExcelerate is a Python 2/3 library for writing Excel-compatible XLSX spreadsheet files, with an emphasis
on speed.

### Benchmarks
Benchmark code located in pyexcelerate/tests/benchmark.py   
65000 rows x 100 columns of the number 1  
Ubuntu 12.04 LTS, Core i7-2600 3.4GHz, 16GB DDR3, Python 2.7.3

* PyExcelerate 25.5s
* openpyxl (1.6.2): 38.62s (1.5x slower)
* xlsxwriter (0.3.2): 367.26s (14.4x slower)


## Installation

    pip install pyexcelerate

## Usage

### Writing bulk data

```python
from pyexcelerate import Workbook

data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]] # data is a 2D array

wb = Workbook()
wb.new_sheet("sheet name", data=data)
wb.save("output.xlsx")

```

### Writing bulk data to a range

PyExcelerate also permits you to write data to ranges directly, which is faster than writing cell-by-cell.

```python
from pyexcelerate import Workbook

wb = Workbook()
ws = wb.new_sheet("test")
ws.range("B2", "C3").value = [[1, 2], [3, 4]]
wb.save("output.xlsx")

```

### Writing cell data

```python
from pyexcelerate import Workbook

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws[1][1].value = 15 # a number
ws[1][2].value = 20
ws[1][3].value = "=SUM(A1,B1)" # a formula
ws[1][4].value = datetime.now() # a date
wb.save("output.xlsx")

```

### Selecting cells by name

```python
from pyexcelerate import Workbook

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws.cell("A1").value = 12
wb.save("output.xlsx")

```

### Merging cells

```python
from pyexcelerate import Workbook

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws[1][1].value = 15
ws.range("A1", "B1").merge()
wb.save("output.xlsx")

```

## Support
Please use the GitHub Issue Tracker and pull request system to report bugs/issues and submit improvements/changes, respectively.  
All changes to code must be accompanied by a test. We use the Nose testing framework.
