from . import Range
from .DataTypes import DataTypes
from . import six
from datetime import datetime

class Worksheet(object):
	def __init__(self, name, workbook, data=None):
		self._columns = 0 # cache this for speed
		self._name = name
		self._cells = {}
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
	
	def report_column(self, column):
		# listener for column additions
		self._columns = max(self._columns, column)
	
	def add_merge(self, range):
		for merge in self._merges:
			if range.intersects(merge):
				raise Exception("Invalid merge, intersects existing")
		self._merges.append(range)
	
	def get_cell_value(self, x, y):
		if x not in self._cells:
			self._cells[x] = {}
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
		self._cells[x][y] = value
	
	@property
	def workbook(self):
			return self._parent

	_cell_cache = {}
	def __get_cell_data(self, cell, x, y):
		if cell not in self._cell_cache:
			type = DataTypes.get_type(cell)
			if type == DataTypes.NUMBER:
				self._cell_cache[cell] = '" t="n"><v>%.15g</v></c>' % (cell)
			elif type == DataTypes.INLINE_STRING:
				self._cell_cache[cell] = '" t="inlineStr"><is><t>%s</t></is></c>' % (cell)
			elif type == DataTypes.DATE:
				self._cell_cache[cell] = '" t="d"><v>%s</v></c>' % (cell.strftime("%Y-%m-%dT%H:%M:%S.%f"))
			elif type == DataTypes.FORMULA:
				self._cell_cache[cell] = '"><f>%s</f></c>' % (cell)
		# Don't cache the coordinate location
		return '<c r="' + Range.Range.coordinate_to_string((x, y)) + self._cell_cache[cell]
	
	def get_xml_data(self):
		# initialize the shared string hashtable
		# self.shared_strings = SharedStrings.SharedStrings(self)
		for x, row in six.iteritems(self._cells):
			row_data = []
			for y, cell in six.iteritems(self._cells[x]):
				row_data.append(self.__get_cell_data(cell, x, y))
			yield x, row_data
