from . import Worksheet
from .Writer import Writer
import time

class Workbook(object):
	# map for attribute sets => style attribute id's
	STYLE_ATTRIBUTE_MAP = {'fonts':'_font', 'fills':'_fill', 'num_fmts':'_format'}
	STYLE_ID_ATTRIBUTE = 'id'
	alignment = None
	def __init__(self, encoding='utf-8'):
		self._worksheets = []
		self._styles = []
		self._items = {} #dictionary containing lists of fonts, fills, etc.
		self._encoding = encoding
		self._writer = Writer(self)

	def add_sheet(self, worksheet):
		self._worksheets.append(worksheet)
		
	def new_sheet(self, sheet_name, data=None):
		worksheet = Worksheet.Worksheet(sheet_name, self, data)
		self._worksheets.append(worksheet)
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

	def get_xml_data(self):
		if Workbook.alignment != self:
			self._align_styles() # because it will be used by the worksheets later
		for index, ws in enumerate(self._worksheets, start=1):
			yield (index, ws)

	def _align_styles(self):
		if Workbook.alignment != self or len(self._items) == 0:
			Workbook.alignment = self
			items = dict([(x, {}) for x in Workbook.STYLE_ATTRIBUTE_MAP.keys()])
			styles = {}
			for index, style in enumerate(self._styles):
				# compress style
				if not style.is_default:
					if style not in styles:
						styles[style] = len(styles) + 1
						setattr(style, Workbook.STYLE_ID_ATTRIBUTE, styles[style])
						# compress individual attributes
						for attr, attr_id in Workbook.STYLE_ATTRIBUTE_MAP.items():
							obj = getattr(style, attr_id)
							if obj and not obj.is_default: # we only care about it if it's not default
								if obj not in items[attr]:
									items[attr][obj] = len(items[attr]) + 1 # insert it
								obj.id = items[attr][obj] # apply
					else:
						setattr(style, Workbook.STYLE_ID_ATTRIBUTE, styles[style])
			for k, v in items.items():
				# ensure it's sorted properly
				items[k] = [tup[0] for tup in sorted(v.items(), key=lambda x: x[1])]
			self._items = items
			self._styles = [tup[0] for tup in sorted(styles.items(), key=lambda x: x[1])]
	def __getattr__(self, name):
		if Workbook.alignment != self:
			self._align_styles()
		return self._items[name]

	def __len__(self):
		return len(self._worksheets)

	def _save(self, file_handle):
		self._align_styles()
		self._writer.save(file_handle)

	def save(self, filename):
		self._save(open(filename, 'wb'))
