from . import Worksheet
from .Writer import Writer
from . import Utility
import six
import time


class Workbook(object):
    # map for attribute sets => style attribute id's
    STYLE_ATTRIBUTE_MAP = {
        "fonts": "_font",
        "fills": "_fill",
        "num_fmts": "_format",
        "borders": "_borders",
    }
    STYLE_ID_ATTRIBUTE = "id"
    __slots__ = ("_worksheets", "_styles", "_items", "_has_macros", "_encoding", "_writer")

    def __init__(self, encoding="utf-8"):
        self._worksheets = []
        self._styles = []
        self._items = {}  # dictionary containing lists of fonts, fills, etc.
        self._has_macros = False
        self._encoding = encoding
        self._writer = Writer(self)

    def add_sheet(self, worksheet):
        for sheet in self._worksheets:
            if sheet.name == worksheet.name:
                raise Exception(
                    "There is already a worksheet with the name '%s'. Duplicate worksheet names are not permitted."
                    % worksheet.name
                )
        self._worksheets.append(worksheet)

    def new_sheet(self, sheet_name, data=None, force_name=False):
        worksheet = Worksheet.Worksheet(sheet_name, self, data, force_name)
        self.add_sheet(worksheet)
        return worksheet

    def add_style(self, style):
        # keep them all, even if they're deleted. compress later.
        self._styles.append(style)

    @property
    def has_styles(self):
        return len(self._styles) > 0

    @property
    def styles(self):
        self._align_styles()
        return self._styles
    
    @property
    def has_macros(self):
        return self._has_macros

    def get_xml_data(self):
        self._align_styles()  # because it will be used by the worksheets later
        for index, ws in enumerate(self._worksheets, start=1):
            yield (index, ws)

    def _align_styles(self):
        Utility.YOLO = True
        items = dict([(x, {}) for x in Workbook.STYLE_ATTRIBUTE_MAP.keys()])
        styles = {}
        for index, style in enumerate(self._styles):
            # compress style
            if not style.is_default:
                styles[style] = styles.get(style, len(styles) + 1)
                setattr(style, Workbook.STYLE_ID_ATTRIBUTE, styles[style])
        for style in styles.keys():
            # compress individual attributes
            for attr, attr_id in Workbook.STYLE_ATTRIBUTE_MAP.items():
                obj = getattr(style, attr_id)
                if (
                    obj and not obj.is_default
                ):  # we only care about it if it's not default
                    items[attr][obj] = items[attr].get(obj, len(items[attr]) + 1)
                    obj.id = items[attr][obj]  # apply
        for k, v in items.items():
            # ensure it's sorted properly
            items[k] = [tup[0] for tup in sorted(v.items(), key=lambda x: x[1])]
        self._items = items
        self._styles = [tup[0] for tup in sorted(styles.items(), key=lambda x: x[1])]
        Utility.YOLO = False

    def __getattr__(self, name):
        self._align_styles()
        return self._items[name]

    def __len__(self):
        return len(self._worksheets)

    def _save(self, file_handle):
        self._align_styles()
        self._writer.save(file_handle)

    def save(self, filename_or_filehandle, has_macros=False):
        self._has_macros = has_macros
        if isinstance(filename_or_filehandle, six.string_types):
            with open(filename_or_filehandle, "wb") as fp:
                self._save(fp)
        else:
            self._save(filename_or_filehandle)
