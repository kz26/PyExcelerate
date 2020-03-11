from . import DataTypes
import six
from . import Font, Fill, Format, Style
from six.moves import reduce

#
# Kevin and Kevin's fair warning: this class has been insanely optimized for speed. It is intended
# to be immutable. Please don't modify attributes after instantiation. :)
#


class Range(object):
    A = ord("A")
    Z = ord("Z")
    __slots__ = (
        "_start",
        "_end",
        "worksheet",
        "is_cell",
        "is_row",
        "is_column",
        "x",
        "y",
        "height",
        "width",
    )

    def __init__(self, start, end, worksheet, validate=True):
        self._start = (
            Range.string_to_coordinate(start)
            if validate and isinstance(start, six.string_types)
            else start
        )
        self._end = (
            Range.string_to_coordinate(end)
            if validate and isinstance(end, six.string_types)
            else end
        )
        # Following http://office.microsoft.com/en-ca/excel-help/excel-specifications-and-limits-HA103980614.aspx
        if (
            not (1 <= self._start[0] <= 1048576) and self._start[0] != float("inf")
        ) or (not (1 <= self._end[0] <= 1048576) and self._end[0] != float("inf")):
            raise IndexError("Row index out of bounds")
        if (not (1 <= self._start[1] <= 16384) and self._start[1] != float("inf")) or (
            not (1 <= self._end[1] <= 16384) and self._end[1] != float("inf")
        ):
            raise IndexError("Column index out of bounds")
        self.worksheet = worksheet
        self.is_cell = self._start == self._end

        self.is_row = (
            self._end[1] == float("inf")
            and self._start[0] == self._end[0]
            and self._start[1] == 1
        )
        self.is_column = (
            self._end[0] == float("inf")
            and self._start[1] == self._end[1]
            and self._start[0] == 1
        )

        self.x = self._start[0] if self.is_row or self.is_cell else None
        self.y = self._start[1] if self.is_column or self.is_cell else None
        self.height = self._end[0] - self._start[0] + 1
        self.width = self._end[1] - self._start[1] + 1
        if self.is_cell:
            worksheet._columns = max(worksheet._columns, self.y)

    @property
    def coordinate(self):
        if self.is_cell:
            return self._start
        else:
            raise Exception("Non-singleton range selected")

    @property
    def style(self):
        if self.is_row:
            return self.__get_attr(
                self.worksheet.get_cell_style,
                Range.AttributeInterceptor(self.worksheet.get_row_style(self.x), ""),
            )
        return self.__get_attr(
            self.worksheet.get_cell_style, Range.AttributeInterceptor(self, "style")
        )

    @style.setter
    def style(self, data):
        self.__set_attr(self.worksheet.set_cell_style, data)

    @property
    def value(self):
        return self.__get_attr(self.worksheet.get_cell_value)

    @value.setter
    def value(self, data):
        self.__set_attr(self.worksheet.set_cell_value, data)

    # this class permits doing things like range().style.font.bold = True
    class AttributeInterceptor(object):
        def __init__(self, parent, attribute=""):
            self.__dict__["_parent"] = parent
            self.__dict__["_attribute"] = attribute

        def __getattr__(self, name):
            if self._attribute == "":
                return Range.AttributeInterceptor(self._parent, name)
            return Range.AttributeInterceptor(
                self._parent, "%s.%s" % (self._attribute, name)
            )

        def __setattr__(self, name, value):
            if isinstance(self._parent, Style.Style):
                setattr(
                    reduce(getattr, self._attribute.split("."), self._parent),
                    name,
                    value,
                )
            else:
                for cell in self._parent:
                    setattr(
                        reduce(getattr, self._attribute.split("."), cell), name, value
                    )

    # note that these are not the python __getattr__/__setattr__
    def __get_attr(self, method, default=None):
        if self.is_cell:
            for merge in self.worksheet.merges:
                if self in merge:
                    return method(merge._start[0], merge._start[1])
            return method(self.x, self.y)
        elif default:
            return default
        else:
            raise Exception("Non-singleton range selected")

    def __set_attr(self, method, data):
        if self.is_cell:
            for merge in self.worksheet.merges:
                if self in merge:
                    method(merge._start[0], merge._start[1], data)
                    return
            method(self.x, self.y, data)
        elif self.is_row and isinstance(data, Style.Style):
            # Applying a row style
            self.worksheet.set_row_style(self.x, data)
        elif DataTypes.DataTypes.get_type(data) != DataTypes.DataTypes.ERROR:
            # Attempt to apply in batch
            for cell in self:
                cell.__set_attr(method, data)
        else:
            if len(data) <= self.height:
                for row in data:
                    if len(row) > self.width:
                        raise Exception(
                            "Row too large for range, row has %s columns, but range only has %s"
                            % (len(row), self.width)
                        )
                for x, row in enumerate(data):
                    for y, value in enumerate(row):
                        method(x + self._start[0], y + self._start[1], value)
            else:
                raise Exception(
                    "Too many rows for range, data has %s rows, but range only has %s"
                    % (len(data), self.height)
                )

    def intersection(self, range):
        """
		Calculates the intersection with another range object
		"""
        if self.worksheet != range.worksheet:
            # Different worksheet
            return None
        start = (
            max(self._start[0], range._start[0]),
            max(self._start[1], range._start[1]),
        )
        end = (min(self._end[0], range._end[0]), min(self._end[1], range._end[1]))
        if end[0] < start[0] or end[1] < start[1]:
            return None
        return Range(start, end, self.worksheet, validate=False)

    __and__ = intersection

    def intersects(self, range):
        return self.intersection(range) is not None

    def merge(self):
        self.worksheet.add_merge(self)

    def __iter__(self):
        if self.is_row or self.is_column:
            raise Exception("Can't iterate over an infinite row/column")
        for x in range(self._start[0], self._end[0] + 1):
            for y in range(self._start[1], self._end[1] + 1):
                yield Range((x, y), (x, y), self.worksheet, validate=False)

    def __contains__(self, item):
        return self.intersection(item) == item

    def __hash__(self):
        def hash(val):
            return val[0] << 8 + val[1]

        return hash(self._start) << 24 + hash(self._end)

    def __str__(self):
        return (
            Range.coordinate_to_string(self._start)
            + ":"
            + Range.coordinate_to_string(self._end)
        )

    def __len__(self):
        if self._start[0] == self._end[0]:
            return self.width
        else:
            return self.height

    def __eq__(self, other):
        if other is None:
            return False
        return self._start == other._start and self._end == other._end

    def __ne__(self, other):
        return not (self == other)

    def __getitem__(self, key):
        if self.is_row:
            # return the key'th column
            if isinstance(key, six.string_types):
                key = Range.string_to_coordinate(key)
            return Range((self.x, key), (self.x, key), self.worksheet, validate=False)
        elif self.is_column:
            # return the key'th row
            return Range((key, self.y), (key, self.y), self.worksheet, validate=False)
        else:
            raise Exception("Selection not valid")

    def __setitem__(self, key, value):
        if self.is_row:
            self.worksheet.set_cell_value(self.x, key, value)
        else:
            raise Exception("Couldn't set that")

    @staticmethod
    def string_to_coordinate(s):
        # Convert a base-26 name to integer
        y = 0
        l = len(s)
        for index, c in enumerate(s):
            if ord(c) < Range.A or ord(c) > Range.Z:
                s = s[index:]
                break
            y *= 26
            y += ord(c) - Range.A + 1
        if len(s) == l:
            return y
        else:
            return (int(s), y)

    @staticmethod
    def coordinate_to_string(coord):
        if coord[1] == float("inf"):
            return "IV%s" % str(coord[0])

        # convert an integer to base-26 name
        y = coord[1] - 1
        s = ""
        while y >= 0:
            d, m = divmod(y, 26)
            s = chr(m + Range.A) + s
            y = d - 1
        return s + str(coord[0])
