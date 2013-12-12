# PyExcelerate

Accelerated Excel XLSX writing library for Python 2/3

[![Build Status](https://travis-ci.org/kz26/PyExcelerate.png?branch=master)](https://travis-ci.org/kz26/PyExcelerate)
[![Coverage Status](https://coveralls.io/repos/kz26/PyExcelerate/badge.png)](https://coveralls.io/r/kz26/PyExcelerate)

* Current version: 0.4.1
* Authors: [Kevin Wang](https://github.com/kevmo314) and [Kevin Zhang](https://github.com/kz26)
* License: Simplified BSD License
* [Source repository](https://github.com/kz26/PyExcelerate)
* [PyPI page](https://pypi.python.org/pypi/PyExcelerate)

#### Donate to the Developers

Did you find PyExcelerate useful? Then consider a donation of bitcoins to the developers and help fund these poor hungry students.
When donating, please split 50/50.     

`1M9MAvqn6eD41FC4bfLaWaKJ9qhaDUCDVb` Kevin Wang

`16jBgGqW6x945545a8nso4bhhDrqmZCQZq` Kevin Zhang

## Description
PyExcelerate is a Python 2/3 library for writing Excel-compatible XLSX spreadsheet files, with an emphasis
on speed.

### Benchmarks
Benchmark code located in pyexcelerate/tests/benchmark.py   
65000 rows x 100 columns of the number 1  
Ubuntu 12.04 LTS, Core i7-2600 3.4GHz, 16GB DDR3, Python 2.7.3

* PyExcelerate 25.5s
* xlsxwriter (0.3.2): 38.62s (1.5x slower)
* openpyxl (1.6.2): 367.26s (14.4x slower)


## Installation

    pip install pyexcelerate

## Usage

### Writing bulk data

#### Fastest

```python
from pyexcelerate import Workbook

data = [[1, 2, 3], [4, 5, 6], [7, 8, 9]] # data is a 2D array

wb = Workbook()
wb.new_sheet("sheet name", data=data)
wb.save("output.xlsx")

```

### Writing bulk data to a range

PyExcelerate also permits you to write data to ranges directly, which is faster than writing cell-by-cell.

#### Fastest

```python
from pyexcelerate import Workbook

wb = Workbook()
ws = wb.new_sheet("test")
ws.range("B2", "C3").value = [[1, 2], [3, 4]]
wb.save("output.xlsx")

```

### Writing cell data

#### Faster

```python
from datetime import datetime
from pyexcelerate import Workbook

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws.set_cell_value(1, 1, 15) # a number
ws.set_cell_value(1, 2, 20)
ws.set_cell_value(1, 3, "=SUM(A1,B1)") # a formula
ws.set_cell_value(1, 4, datetime.now()) # a date
wb.save("output.xlsx")

```

#### Fast

```python
from datetime import datetime
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

### Styling cells

Styling cells causes **non-negligible** overhead. It **will** increase your execution time (up to 10x longer if done improperly!). Only style cells if absolutely necessary.

#### Fastest

```python
from pyexcelerate import Workbook, Color, Style, Font, Fill
from datetime import datetime

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws.set_cell_value(1, 1, 1)
ws.set_cell_style(1, 1, Style(font=Font(bold=True)))
ws.set_cell_style(1, 1, Style(font=Font(italic=True)))
ws.set_cell_style(1, 1, Style(font=Font(underline=True)))
ws.set_cell_style(1, 1, Style(font=Font(strikethrough=True)))
ws.set_cell_style(1, 1, Style(fill=Fill(background=Color(255,0,0,0))))
ws.set_cell_value(1, 2, datetime.now())
ws.set_cell_style(1, 1, Style(format=Format('mm/dd/yy')))
wb.save("output.xlsx")

```

#### Faster

```python
from pyexcelerate import Workbook, Color
from datetime import datetime

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws.set_cell_value(1, 1, 1)
ws.get_cell_style(1, 1).font.bold = True
ws.get_cell_style(1, 1).font.italic = True
ws.get_cell_style(1, 1).font.underline = True
ws.get_cell_style(1, 1).font.strikethrough = True
ws.get_cell_style(1, 1).fill.background = Color(0, 255, 0, 0)
ws.set_cell_value(1, 2, datetime.now())
ws.get_cell_style(1, 1).format.format = 'mm/dd/yy'
wb.save("output.xlsx")

```
#### Fast

```python
from pyexcelerate import Workbook, Color
from datetime import datetime

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws[1][1].value = 1
ws[1][1].style.font.bold = True
ws[1][1].style.font.italic = True
ws[1][1].style.font.underline = True
ws[1][1].style.font.strikethrough = True
ws[1][1].style.fill.background = Color(0, 255, 0, 0)
ws[1][2].value = datetime.now()
ws[1][2].style.format.format = 'mm/dd/yy'
wb.save("output.xlsx")

```

**Note** that `.style.format.format`'s repetition is due to planned support for conditional formatting and other related features. The formatting syntax may be improved in the future.

### Styling ranges

```python
from pyexcelerate import Workbook, Color
from datetime import datetime

wb = Workbook()
ws = wb.new_sheet("test")
ws.range("A1","C3").value = 1
ws.range("A1","C1").style.font.bold = True
ws.range("A2","C3").style.font.italic = True
ws.range("A3","C3").style.fill.background = Color(255, 0, 0, 0)
ws.range("C1","C3").style.font.strikethrough = True
```

### Styling rows

A simpler (and faster) way to style an entire row.

#### Fastest

```python
from pyexcelerate import Workbook, Color, Style, Fill
from datetime import datetime

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws.set_row_style(1, Style(fill=Fill(background=Color(255,0,0,0))))
wb.save("output.xlsx")

```

#### Faster

```python
from pyexcelerate import Workbook, Color
from datetime import datetime

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws.get_row_style(1).fill.background = Color(255, 0, 0)
wb.save("output.xlsx")

```

#### Fast

```python
from pyexcelerate import Workbook, Color
from datetime import datetime

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws[1].style.fill.background = Color(255, 0, 0)
wb.save("output.xlsx")

```

### Linked styles

PyExcelerate supports using style objects instead manually setting each attribute as well. This permits you to modify the style at a later time.

```python
from pyexcelerate import Workbook, Font

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws[1][1].value = 1
font = Font(bold=True, italic=True, underline=True, strikethrough=True)
ws[1][1].style.font = font
wb.save("output.xlsx")

```

## Packaging with PyInstaller, cx\_freeze, py2exe, etc.

PyExelerate has been successfully tested with PyInstaller and cx\_freeze; py2exe should work as well.
**However, you will need to manually copy the templates directory from PyExcelerate into your app's distribution.**

## Support
Please use the GitHub Issue Tracker and pull request system to report bugs/issues and submit improvements/changes, respectively.  
All changes to code must be accompanied by a test. We use the Nose testing framework.
