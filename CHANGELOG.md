0.6.8
* Add requirement for six >=1.4.0 

0.6.7
* Fix UnicodeDecodeError with unicode cell content (#35) and add unicode test case

0.6.6
* enhancement: empty stylesheet behavior for openpyxl compatibility and improved performance (issue #34)
* Bugfix: zero values deleted when styles are set (issue #30)

0.6.5
* Add missing Color import in init (PR #29)

0.6.4
* Add check for duplicate worksheet names
* Add check for worksheet name length (for Excel compatibility)
* Autofit bug fix (issue #28)
* other minor fixes

0.6.3
* Max rows fix (issues #25 and #27)

0.6.2
* Python 3.4 added to list of supported Python versions

0.6.1 - May 11, 2014
* Add PyInstaller hook
* Indicate that PyInstaller is the only supported packager

0.6.0 - May 10, 2014
* Support for setting row/column widths
* Fix issue #22 - Decimal data type (thanks rhyek)

0.5.0 - February 7, 2014 - "YOLO"
* Implement new "YOLO" mode for increased PyExceleration
* Fix tests/benchmark.py and update benchmark data

0.4.1 - December 4, 2013
* Fix import errors caused by missing templates
* Remove bundled six.py and use system-wide six.py instead

0.4.0 - November 15 , 2013
* Add basic style support (see README.md for usage and examples)
* Merge and intersection logic bugfixes (thanks morty)
* Fixed float precision bug (thanks jmcnamara)

0.3.0 - July 29, 2013
* Fixed bug in \_\_coordinate\_to\_string function
* Compatibility fixes for Python 3 and numpy datatypes
* Updated test suite to work with Python 3

0.2.6 - July 10, 2013
* Better PyInstaller compatibility (sys.executable in addition to sys.\_MEIPASS)

0.2.5 - May 14, 2013
* Fixed int/float bug (thanks Redoubts)

0.2.4 - May 2, 2013
* Increased exceleration (in response to issue #1)

0.2.3 - April 23, 2013
* #3 - fix numpy and other subclasses support (thanks Redoubts)
* #4 - fix issue with range writer (thanks Redoubts)

0.2.2 - April 22, 2013
* #2 - template path detection fix for PyInstaller (thanks sh0375)

0.2.1 - April 21, 2013
* datetime bugfixes
* optimization for type-checking
* UTF-8 encoding fix (thanks Genmutant from Reddit)

0.2.0 - April 20, 2013
* Add datetime support
* Add Workbook.cell() call

0.1.0 - April 20, 2013
* Initial release
