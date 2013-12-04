from . import Range
from . import Style
from . import Format
from .DataTypes import DataTypes
import six
from datetime import datetime
from xml.sax.saxutils import escape

class Worksheet(object):
	def __init__(self, name, workbook, data=None):
		self._columns = 0 # cache this for speed
		self._name = name
		self._cells = {}
		self._cell_cache = {}
		self._styles = {}
		self._row_styles = {}
		self._parent = workbook
		self._merges = [] # list of Range objects
		self._attributes = {}
		if data != None:
			for x, row in enumerate(data, 1):
				for y, cell in enumerate(row, 1):
					if x not in self._cells:
						self._cells[x] = {}
					self._cells[x][y] = cell
					self._columns = max(self._columns, y)

	def __getitem__(self, key):
		if key not in self._cells:
			self._cells[key] = {}
		return Range.Range((key, 1), (key, float('inf')), self) # return a row range

	@property
	def stylesheet(self):
		return self._stylesheet

	@property
	def name(self):
		return self._name
	
	@property
	def merges(self):
		return self._merges
	
	@property
	def num_rows(self):
		if len(self._cells) > 0:
			return max(self._cells.keys())
		else:
			return 1
	
	@property
	def num_columns(self):
		return max(1, self._columns)
	
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
		if x not in self._cells:
			self._cells[x] = {}
		if y not in self._cells[x]:
			return None
		type = DataTypes.get_type(self._cells[x][y])
		if type == DataTypes.FORMULA:
			# remove the equals sign
			return self._cells[x][y][:1]
		elif type == DataTypes.INLINE_STRING and self._cells[x][y][2:] == '\'=':
			return self._cells[x][y][:1]
		else:
			return self._cells[x][y]
	
	def set_cell_value(self, x, y, value):
		if x not in self._cells:
			self._cells[x] = {}
		if DataTypes.get_type(value) == DataTypes.DATE:
			self.get_cell_style(x, y).format = Format.Format('yyyy-mm-dd')
		self._cells[x][y] = value
	
	def get_cell_style(self, x, y):
		if x not in self._styles:
			self._styles[x] = {}
		if y not in self._styles[x]:
			self.set_cell_style(x, y, Style.Style())
		return self._styles[x][y]
	
	def set_cell_style(self, x, y, value):
		if x not in self._styles:
			self._styles[x] = {}
		self._styles[x][y] = value
		self._parent.add_style(value)
		if not self.get_cell_value(x, y):
			self.set_cell_value(x, y, '')
	
	def get_row_style(self, row):
		if row not in self._row_styles:
			self.set_row_style(row, Style.Style())
		return self._row_styles[row]
		
	def set_row_style(self, row, value):
		self._row_styles[row] = value
		self._parent.add_style(value)
	
	@property
	def workbook(self):
			return self._parent

	def __get_cell_data(self, cell, x, y, style):
		if cell not in self._cell_cache:
			type = DataTypes.get_type(cell)
			
			if type == DataTypes.NUMBER:
				self._cell_cache[cell] = '"><v>%.15g</v></c>' % (cell)
			elif type == DataTypes.INLINE_STRING:
				self._cell_cache[cell] = '" t="inlineStr"><is><t>%s</t></is></c>' % escape(cell)
			elif type == DataTypes.DATE:
				self._cell_cache[cell] = '"><v>%s</v></c>' % (DataTypes.to_excel_date(cell))
			elif type == DataTypes.FORMULA:
				self._cell_cache[cell] = '"><f>%s</f></c>' % (cell)
		
		if style:
			return "<c r=\"%s\" s=\"%d%s" % (Range.Range.coordinate_to_string((x, y)), style.id, self._cell_cache[cell])
		else:
			return "<c r=\"%s%s" % (Range.Range.coordinate_to_string((x, y)), self._cell_cache[cell])
			
	def get_row_xml_string(self, row):
		if row in self._row_styles:
			return "<row r=\"%d\" s=\"%d\" customFormat=\"1\">" % (row, self._row_styles[row].id)
		else:
			return "<row r=\"%d\">" % row
		
	def get_xml_data(self):
		# Precondition: styles are aligned. if not, then :v
		for x, row in six.iteritems(self._cells):
			row_data = []
			for y, cell in six.iteritems(self._cells[x]):
				if x not in self._styles or y not in self._styles[x]:
					style = None
				else:
					style = self._styles[x][y]
				row_data.append(self.__get_cell_data(cell, x, y, style))
			yield x, row_data
