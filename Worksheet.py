import Range

class Worksheet(object):
	def __init__(self, name):
		self._name = name
		self._cells = []

	def cell(self, x, y=None): # 1-indexed
		if y == None:
			return Range.Range(x, x, self) # parse a cell
		else:
			return Range.range(x, y, self)
	@property
	def num_rows(self):
		if len(self._cells) > 0:
			return max(self._cells.keys())
		else:
			return 1
		
	@property
	def num_columns(self):
		raise Exception("Not implemented")
	
	def data(self, x, y): # 1-indexed
		return self._cells[x][y]
		
	def write(self, xml_stream):
		pass
	