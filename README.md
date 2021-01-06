# PyExcelerate

Accelerated Excel XLSX writing library for Python

master: [![build-status-master](https://travis-ci.org/kz26/PyExcelerate.svg?branch=master)](https://travis-ci.org/kz26/PyExcelerate)
dev: [![build-status-dev](https://travis-ci.org/kz26/PyExcelerate.svg?branch=dev)](https://travis-ci.org/kz26/PyExcelerate)
test coverage: [![coverage-status](https://coveralls.io/repos/kz26/PyExcelerate/badge.svg)](https://coveralls.io/r/kz26/PyExcelerate)

* Authors: [Kevin Wang](https://github.com/kevmo314) and [Kevin Zhang](https://github.com/kz26)
* Copyright 2015 Kevin Wang, Kevin Zhang. Portions copyright Google, Inc.
* License: Simplified BSD License
* [Source repository](https://github.com/kz26/PyExcelerate)
* [PyPI page](https://pypi.python.org/pypi/PyExcelerate)

## Description

PyExcelerate is a Python for writing Excel-compatible XLSX spreadsheet files, with an emphasis
on speed.

### Benchmarks
Benchmark code located in pyexcelerate/tests/benchmark.py   
Ubuntu 12.04 LTS, Core i5-3450, 8GB DDR3, Python 2.7.3

```
|          TEST_NAME          | NUM_ROWS | NUM_COLS | TIME_IN_SECONDS |
|-----------------------------|----------|----------|-----------------|
| pyexcelerate value fastest  |     1000 |      100 |            0.47 |
| pyexcelerate value faster   |     1000 |      100 |            0.51 |
| pyexcelerate value fast     |     1000 |      100 |            1.53 |
| xlsxwriter value            |     1000 |      100 |            0.84 |
| openpyxl                    |     1000 |      100 |            2.74 |
| pyexcelerate style cheating |     1000 |      100 |            1.23 |
| pyexcelerate style fastest  |     1000 |      100 |            2.40 |
| pyexcelerate style faster   |     1000 |      100 |            2.75 |
| pyexcelerate style fast     |     1000 |      100 |            6.15 |
| xlsxwriter style cheating   |     1000 |      100 |            1.21 |
| xlsxwriter style            |     1000 |      100 |            4.85 |
| openpyxl                    |     1000 |      100 |            6.32 |

* cheating refers to pregeneration of styles

```

## Installation

PyExcelerate is supported on Python 2.7, 3.4, 3.5, 3.6, 3.7, and 3.8.

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

PyExcelerate also permits you to write data to ranges directly, which is faster than writing cell-by-cell. If writing a Pandas DataFrame, see the [note on compatibility](#Pandas-DataFrames).

#### Fastest

```python
from pyexcelerate import Workbook

wb = Workbook()
ws = wb.new_sheet("test", data=[[1, 2], [3, 4]])
wb.save("output.xlsx")
```

#### Fast

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
from pyexcelerate import Workbook, Color, Style, Font, Fill, Format
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

Note that in this example, subsequent calls to `set_cell_style()` override the previous styles and they are not merged in. To have a combined style, create a single `Style` object with multiple properties, for example

```
ws.set_cell_style(1, 1, Style(font=Font(bold=True), format=Format('mm/dd/yy')))
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

### Styling columns

#### Fastest

```python
from pyexcelerate import Workbook, Color, Style, Fill
from datetime import datetime

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws.set_col_style(1, Style(fill=Fill(background=Color(255,0,0,0))))
wb.save("output.xlsx")
```

### Available style attributes

Consistent with the implementation patterns above, the following style parameters are available:

```python
ws[1][1].style.font.bold = True
ws[1][1].style.font.italic = True
ws[1][1].style.font.underline = True
ws[1][1].style.font.strikethrough = True
ws[1][1].style.font.color = Color(255, 0, 255)
ws[1][1].style.fill.background = Color(0, 255, 0)
ws[1][1].style.alignment.vertical = 'top'
ws[1][1].style.alignment.horizontal = 'right'
ws[1][1].style.alignment.rotation = 90
ws[1][1].style.alignment.wrap_text = True
ws[1][1].style.borders.top.color = Color(255, 0, 0)
ws[1][1].style.borders.right.style = '-.'
```

Each attribute also has constructors for implementing via `set_cell_style()`.

The following border styles are available: `.-`, `..-`, `--`, `..`, `=`, `.`, `medium -.`, `medium -..`, `medium --`, `/-.`, `_`

### Setting row heights and column widths

Row heights and column widths are set using the `size` attribute in `Style`. Appropriate values are:
*	`-1` for auto-fit
*	`0` for hidden
*	Any other value for the appropriate size.

For example, to hide column B:

```python
from pyexcelerate import Workbook, Color, Style, Fill
from datetime import datetime

wb = Workbook()
ws = wb.new_sheet("sheet name")
ws.set_col_style(2, Style(size=0))
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

## Pandas DataFrames

PyExcelerate does not support directly passing a Pandas DataFrame as the data argument to a new worksheet. If the sheet does not require having the headers rendered, the most efficient solution is:

```python
ws = wb.new_sheet("sheet name", data=df.values.tolist())
```

Note that the conversion `.tolist()` is faster as PyExcelerate has some optimizations surrounding data that's provided in lists. If the sheet needs to have headers rendered, consider [asking the Pandas maintainers](https://github.com/pandas-dev/pandas/issues/4517) to integrate PyExcelerate, [use a transformation function](https://gist.github.com/mapa17/bc04be36e447cab0746a0ec8903cc49f), or convert your DataFrame to a list with the headers included.

## Packaging with PyInstaller

PyInstaller is the only packager officially supported by PyExcelerate. Copy hook-pyexcelerate.Writer.py to your PyInstaller hooks directory.


## Support
Please use the GitHub Issue Tracker and pull request system to report bugs/issues and submit improvements/changes, respectively.  **Pull requests *must* be based against the dev branch - if not, we will reject the PR and ask you to rebase against the correct branch.** All nontrivial changes to code should be accompanied by a test when appropriate. We use the Nose testing framework.
