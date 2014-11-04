from .Workbook import Workbook
from .Style import Style
from .Fill import Fill
from .Font import Font
from .Format import Format
from .Alignment import Alignment
from .Color import Color

try:
	import pkg_resources
	__version__ = pkg_resources.require('PyExcelerate')[0].version
except:
	__version__ = 'unknown'
