import Range
import SharedStrings

class Worksheet(object):
	def __init__(self, name, workbook):
		self._columns = 0 # cache this for speed
		self._name = name
		self._cells = []
		self._parent = workbook

	def __getitem__(self, key):
		if key not in self._cells:
			raise Exception("Not implemented")
		return Range.Range((key, 1), (key, None), self) # return a row range

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
	
	def report_columns(self, column):
		# listener for column additions
		self._columns = max(self._columns, column)
	
	@property
	def workbook(self):
			return self._parent
			
	def get_xml_data(self):
		# initialize the shared string hashtable
		# self.shared_strings = SharedStrings.SharedStrings(self)
		rows = []
		for row in self._cells.keys():
			rows.append((row, Range.Range((row, 1), (row, None), self)))
		return rows