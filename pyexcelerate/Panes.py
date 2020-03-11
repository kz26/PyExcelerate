from . import Range
from .Utility import to_unicode


class Panes(object):
    __slots__ = ("x", "y", "freeze")

    def __init__(self, x=None, y=None, freeze=True):
        self.x = x or 0
        self.y = y or 0
        self.freeze = freeze

    def __bool__(self):
        return any((self.x, self.y))

    def __nonzero__(self):
        return self.__bool__()

    def __eq__(self, other):
        return self.x == other.x and self.y == other.y and self.freeze == other.freeze

    def get_xml(self):
        attrs = {
            "topLeftCell": Range.Range.coordinate_to_string((self.y + 1, self.x + 1))
        }
        if self.freeze:
            attrs["state"] = "frozen"
        if self.x:
            attrs["xSplit"] = self.x
        if self.y:
            attrs["ySplit"] = self.y
        attr_str = " ".join('%s="%s"' % item for item in sorted(attrs.items()))
        return to_unicode("<pane %s/>" % attr_str)
