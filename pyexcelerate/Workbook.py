from . import Worksheet
from .Writer import Writer

class Workbook(object):
	def __init__(self, encoding='utf-8'):
		self._worksheets = []
		self._styles = []
		self._encoding = encoding
		self._writer = Writer(self)

	def add_sheet(self, worksheet):
		self._worksheets.append(worksheet)
		
	def new_sheet(self, sheet_name, data=None):
		worksheet = Worksheet.Worksheet(sheet_name, self, data)
		self._worksheets.append(worksheet)
		return worksheet

	def add_style(self, style):
		if style not in self._styles:
			self._styles.append(style)
	
	@property
	def has_styles(self):
		return len(self._styles) > 0

	@property
	def styles(self):
		self.align_styles()
		return self._styles

	def align_styles(self):
		for index, style in enumerate(self._styles):
			style.id = index + 1
	
	def get_xml_data(self):
		for index, ws in enumerate(self._worksheets, 1):
			yield (index, ws)

	def __len__(self):
		return len(self._worksheets)

	def _save(self, file_handle):
		self._writer.save(file_handle)

	def save(self, filename):
		self._save(open(filename, 'wb'))
