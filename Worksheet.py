import Range
import SharedStrings

class Worksheet(object):
	def __init__(self, name, workbook):
		self._columns = 0 # cache this for speed
		self._name = name
		self._cells = {}
		self._parent = workbook
		self._merges = [] # list of Range objects
		self._attributes = {}

	def __getitem__(self, key):
		if key not in self._cells:
			self._cells[key] = {}
		return Range.Range((key, 1), (key, float('inf')), self) # return a row range

	@property
	def name(self):
		return self._name
		
	@property
	def num_rows(self):
		if len(self._cells) > 0:
			return max(self._cells.keys())
		else:
			return 1
	
	@property
	def num_columns(self):
		return max(1, self._columns)
	
	def report_column(self, column):
		# listener for column additions
		self._columns = max(self._columns, column)
	
	def add_merge(self, range):
		for merge in self._merges:
			if range.intersects(merge):
				raise Exception("Invalid merge, intersects existing")
		self._merges.append(range)
	
	def get_cell_value(self, x, y):
		return self._cells[x][y]
	
	def set_cell_value(self, x, y, value):
		self._cells[x][y] = value
	
	@property
	def workbook(self):
			return self._parent
			
	def get_xml_data(self):
		# initialize the shared string hashtable
		# self.shared_strings = SharedStrings.SharedStrings(self)
		for row in self._cells.keys():
			yield (row, Range.Range((row, 1), (row, float('inf')), self))
