# -*- coding: utf-8 -*-

import collections
import itertools
import math
import re
import sys
from datetime import datetime
from xml.sax.saxutils import escape

import six

from . import Format, Panes, Range, Style
from .DataTypes import DataTypes
from .Utility import to_unicode

# From https://stackoverflow.com/a/22273639/86433
_illegal_unichrs = [
    (0x00, 0x08),
    (0x0B, 0x0C),
    (0x0E, 0x1F),
    (0x7F, 0x84),
    (0x86, 0x9F),
    (0xFDD0, 0xFDDF),
    (0xFFFE, 0xFFFF),
]
if sys.maxunicode >= 0x10000:  # not narrow build
    _illegal_unichrs.extend(
        [
            (0x1FFFE, 0x1FFFF),
            (0x2FFFE, 0x2FFFF),
            (0x3FFFE, 0x3FFFF),
            (0x4FFFE, 0x4FFFF),
            (0x5FFFE, 0x5FFFF),
            (0x6FFFE, 0x6FFFF),
            (0x7FFFE, 0x7FFFF),
            (0x8FFFE, 0x8FFFF),
            (0x9FFFE, 0x9FFFF),
            (0xAFFFE, 0xAFFFF),
            (0xBFFFE, 0xBFFFF),
            (0xCFFFE, 0xCFFFF),
            (0xDFFFE, 0xDFFFF),
            (0xEFFFE, 0xEFFFF),
            (0xFFFFE, 0xFFFFF),
            (0x10FFFE, 0x10FFFF),
        ]
    )
_illegal_ranges = [
    "%s-%s" % (six.unichr(low), six.unichr(high)) for (low, high) in _illegal_unichrs
]
_illegal_xml_chars_RE = re.compile(u"[%s]" % u"".join(_illegal_ranges))


class Worksheet(object):
    __slots__ = (
        "_columns",
        "_name",
        "_dense_cells",
        "_sparse_cells",
        "_styles",
        "_row_styles",
        "_col_styles",
        "_parent",
        "_merges",
        "_attributes",
        "_panes",
        "_show_grid_lines",
        "auto_filter",
    )

    def __init__(self, name, workbook, data=None, force_name=False):
        self._columns = 0  # cache this for speed
        if len(name) > 31 and not force_name:
            # http://stackoverflow.com/questions/3681868/is-there-a-limit-on-an-excel-worksheets-name-length
            raise Exception(
                "Excel does not permit worksheet names longer than 31 characters. Set force_name=True to disable this restriction."
            )
        self._name = name
        self._dense_cells = [[]]
        self._sparse_cells = collections.defaultdict(dict)
        self._styles = collections.defaultdict(dict)
        self._row_styles = {}
        self._col_styles = {}
        self._parent = workbook
        self._merges = []  # list of Range objects
        self._attributes = {}
        self._panes = Panes.Panes()
        self._show_grid_lines = True
        self.auto_filter = False
        if data is not None:
            # Iterate over the data to ensure we receive a copy of immutables.
            if isinstance(data, list):
                self._dense_cells = [[] for i in range(len(data) + 1)]
            for x, row in enumerate(data, 1):
                if isinstance(row, list) and x < len(self._dense_cells):
                    self._dense_cells[x] = [None] + row[:]
                    self._columns = max(self._columns, len(row))
                else:
                    for y, cell in enumerate(row, 1):
                        self._sparse_cells[x][y] = cell
                        self._columns = max(self._columns, y)

    def __getitem__(self, key):
        if isinstance(key, slice):
            if key.step is not None and key.step > 1:
                raise Exception("PyExcelerate doesn't support slicing with steps")
            else:
                return Range.Range(
                    (key.start or 1, 1), (key.stop or float("inf"), float("inf")), self
                )
        else:
            return Range.Range(
                (key, 1), (key, float("inf")), self
            )  # return a row range

    @property
    def panes(self):
        return self._panes

    @panes.setter
    def panes(self, panes):
        if not isinstance(panes, Panes.Panes):
            raise TypeError("Worksheet.panes must be of type Panes")
        self._panes = panes

    @property
    def col_styles(self):
        return self._col_styles.items()

    @property
    def name(self):
        return self._name

    @property
    def merges(self):
        return self._merges

    @property
    def num_rows(self):
        return max(
            len(self._dense_cells) - 1,
            max(six.iterkeys(self._sparse_cells)) if len(self._sparse_cells) > 0 else 0,
        )

    @property
    def num_columns(self):
        return self._columns

    @property
    def show_grid_lines(self):
        return self._show_grid_lines

    @show_grid_lines.setter
    def show_grid_lines(self, show_grid_lines):
        self._show_grid_lines = show_grid_lines

    def cell(self, name):
        # convenience method
        return self.range(name, name)

    def range(self, start, end):
        # convenience method
        return Range.Range(start, end, self)

    def add_merge(self, range):
        for merge in self._merges:
            if range.intersects(merge):
                raise Exception("Invalid merge, intersects existing")
        self._merges.append(range)

    def get_cell_value(self, x, y):
        if x < len(self._dense_cells) and y < len(self._dense_cells[x]):
            return self._dense_cells[x][y]
        # Fallback to sparse cells
        return self._sparse_cells[x].get(y)

    def set_cell_value(self, x, y, value):
        if DataTypes.get_type(value) == DataTypes.DATE:
            self.get_cell_style(x, y).format = Format.Format("yyyy-mm-dd")
        if x < len(self._dense_cells) and y < len(self._dense_cells[x]):
            self._dense_cells[x][y] = value
        else:
            self._sparse_cells[x][y] = value
        self._columns = max(self._columns, y)

    def get_cell_style(self, x, y):
        if y not in self._styles[x]:
            self.set_cell_style(x, y, Style.Style())
        return self._styles[x][y]

    def set_cell_style(self, x, y, value):
        self._styles[x][y] = value
        self._parent.add_style(value)
        if self.get_cell_value(x, y) is None:
            self.set_cell_value(x, y, "")

    def get_row_style(self, row):
        if row not in self._row_styles:
            self.set_row_style(row, Style.Style())
        return self._row_styles[row]

    def set_row_style(self, row, value):
        if hasattr(row, "__iter__"):
            for r in row:
                self.set_row_style(r, value)
        else:
            self._row_styles[row] = value
            self.workbook.add_style(value)

    def get_col_style(self, col):
        if col not in self._col_styles:
            self.set_col_style(col, Style.Style())
        return self._col_styles[col]

    def set_col_style(self, col, value):
        if hasattr(col, "__iter__"):
            for c in col:
                self.set_col_style(c, value)
        else:
            self._col_styles[col] = value
            self.workbook.add_style(value)

    @property
    def workbook(self):
        return self._parent

    def __get_cell_data(self, cell, x, y, style):
        if cell is None:
            return ""  # no cell data
        # boolean values are treated oddly in dictionaries, manually override
        type = DataTypes.get_type(cell)

        if type == DataTypes.NUMBER:
            if math.isnan(cell):
                z = '" t="e"><v>#NUM!</v></c>'
            elif math.isinf(cell):
                z = '" t="e"><v>#DIV/0!</v></c>'
            else:
                z = '"><v>%.15g</v></c>' % (cell)
        elif type == DataTypes.INLINE_STRING or type == DataTypes.ERROR:
            # Also serialize errors to string, we'll try our best...
            z = '" t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>' % escape(
                _illegal_xml_chars_RE.sub(u"\uFFFD", to_unicode(cell if isinstance(cell, six.string_types) else str(cell)))
            )
        elif type == DataTypes.DATE:
            z = '"><v>%s</v></c>' % (DataTypes.to_excel_date(cell))
        elif type == DataTypes.FORMULA:
            z = '"><f>%s</f></c>' % (cell[1:])  # Remove equals sign.
        elif type == DataTypes.BOOLEAN:
            z = '" t="b"><v>%d</v></c>' % (cell)

        if style and hasattr(style, "id"):
            return '<c r="%s" s="%d%s' % (
                Range.Range.coordinate_to_string((x, y)),
                style.id,
                z,
            )
        else:
            return '<c r="%s%s' % (Range.Range.coordinate_to_string((x, y)), z)

    def get_col_xml_string(self, col):
        if col not in self._col_styles or self._col_styles[col].is_default:
            return '<col min="%d" max="%d">' % (col, col)
        style = self._col_styles[col]
        if style.size == -1:
            size = 0

            def get_size(v):
                if isinstance(v, six.string_types):
                    v = to_unicode(v)
                else:
                    v = six.text_type(v)
                return (len(v) * 7 + 5) / 7

            for row in self._dense_cells[1:]:
                if col < len(row):
                    size = max(get_size(row[col]), size)
            for row in six.itervalues(self._sparse_cells):
                if col in row:
                    size = max(get_size(row[col]), size)
        elif DataTypes.get_type(style.size) == DataTypes.NUMBER:
            size = style.size
        else:
            return '<col min="%d" max="%d" hidden="0" width="9.2" style="%d">' % (
                col,
                col,
                style.id,
            )
        return (
            '<col min="%d" max="%d" hidden="%d" bestFit="%d" customWidth="%d" width="%f" style="%d">'
            % (
                col,
                col,
                1 if style.size == 0 else 0,  # hidden
                1 if style.size == -1 else 0,  # best fit
                1 if style.size is not None else 0,  # customWidth
                size,
                style.id,
            )
        )

    def get_row_xml_string(self, row):
        if row in self._row_styles and not self._row_styles[row].is_default:
            style = self._row_styles[row]
            if style.size == -1:
                size = 0
                dense_rows = (
                    enumerate(self._dense_cells[row][1:])
                    if row < len(self._dense_cells)
                    else []
                )
                for y, cell in itertools.chain(
                    dense_rows,
                    six.iteritems(self._sparse_cells[row])
                    if row in self._sparse_cells
                    else [],
                ):
                    try:
                        font_size = self._styles[row][y].font.size
                    except:
                        font_size = 11
                    lines = (
                        cell.count("\n")
                        if DataTypes.get_type(style.size) == DataTypes.STRING
                        else 1
                    )
                    size = max(font_size * (lines + 1) * 4 / 3, size)
            else:
                size = style.size if style.size else 15
            return (
                '<row r="%d" s="%d" customFormat="1" hidden="%d" customHeight="%d" ht="%f">'
                % (
                    row,
                    style.id,
                    1 if style.size == 0 else 0,  # hidden
                    1 if style.size is not None else 0,  # customHeight
                    size,
                )
            )
        else:
            return '<row r="%d">' % row

    def get_xml_data(self):
        # Precondition: styles are aligned. if not, then :v
        # check if we have any row styles that don't have data
        sparse_rows = sorted(
            filter(
                lambda x: x[0] >= len(self._dense_cells),
                six.iteritems(self._sparse_cells),
            )
        )
        for x, row in itertools.chain(enumerate(self._dense_cells[1:], 1), sparse_rows):
            row_data = []
            dense_columns = enumerate(row[1:], 1) if x < len(self._dense_cells) else []
            sparse_columns = sorted(
                six.iteritems(self._sparse_cells[x]) if x in self._sparse_cells else []
            )
            for y, cell in itertools.chain(dense_columns, sparse_columns):
                if x in self._styles and y in self._styles[x]:
                    style = self._styles[x][y]
                elif x in self._row_styles:
                    style = self._row_styles[x]
                elif y in self._col_styles:
                    style = self._col_styles[y]
                else:
                    style = None
                row_data.append(self.__get_cell_data(cell, x, y, style))
            yield x, row_data

    def get_auto_filter_xml_string(self):
        if self.auto_filter:
            return '<autoFilter ref="{}"/>'.format(
                Range.Range((1, 1), (self.num_rows, self.num_columns), self)
            )
