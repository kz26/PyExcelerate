from .Workbook import Workbook  # noqa
from .Style import Style  # noqa
from .Fill import Fill  # noqa
from .Font import Font  # noqa
from .Format import Format  # noqa
from .Alignment import Alignment  # noqa

try:
    import pkg_resources
    __version__ = pkg_resources.require('PyExcelerate')[0].version
except:
    __version__ = 'unknown'
