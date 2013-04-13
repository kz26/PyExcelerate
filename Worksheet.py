import Range
import SharedStrings

class Worksheet(object):
	def __init__(self, name, workbook):
		self._name = name
		self._cells = []
		self._parent = workbook

	def __getitem__(self, key):
		if key not in self._cells:
			raise Exception("Not implemented")
		return Range.Range(key, key, self) # return a row range

	@property
	def num_rows(self):
		if len(self._cells) > 0:
			return max(self._cells.keys())
		else:
			return 1
	
	@property
	def num_columns(self):
		raise Exception("Not implemented")
	
	@property
	def workbook(self):
			return self._parent
			
	def get_xml_data(self):
		# initialize the shared string hashtable
		# self.shared_strings = SharedStrings.SharedStrings(self)
		rows = []
		for row in self._cells.keys():
			rows.append((row, Range.Range(row, row, self)))
		return rows