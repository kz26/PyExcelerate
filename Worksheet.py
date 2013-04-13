import Range

class Worksheet(object):
	def __init__(self, name):
		self._name = name
		self._cells = []

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
		
	def write(self, xml_stream):
		pass
	