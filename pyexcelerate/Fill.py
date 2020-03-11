from . import Utility
from . import Color


class Fill(object):
    __slots__ = ("_background", "id")

    def __init__(self, background=None):
        self._background = background

    @property
    def background(self):
        return Utility.lazy_get(self, "_background", Color.Color())

    @background.setter
    def background(self, value):
        Utility.lazy_set(self, "_background", None, value)

    @property
    def is_default(self):
        return self == Fill()

    def __eq__(self, other):
        if other is None:
            return self.is_default
        return self._background == other._background

    def __hash__(self):
        return hash(self.background)

    def get_xml_string(self):
        if not self.background:
            return '<fill><patternFill patternType="none"/></fill>'
        else:
            return (
                '<fill><patternFill patternType="solid"><fgColor rgb="%s"/></patternFill></fill>'
                % self.background.hex
            )

    def __or__(self, other):
        return Fill(
            background=Utility.nonboolean_or(self._background, other._background, None)
        )

    def __and__(self, other):
        return Fill(
            background=Utility.nonboolean_and(self._background, other._background, None)
        )

    def __xor__(self, other):
        return Fill(
            background=Utility.nonboolean_xor(self._background, other._background, None)
        )

    def __str__(self):
        return "Fill: #%s" % self.background.hex

    def __repr__(self):
        return "<%s>" % self.__str__()
