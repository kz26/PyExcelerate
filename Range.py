class Range(object):
	A = ord('A')
	Z = ord('Z')
	def __init__(self, start, end, worksheet):
		self._start = Range.to_coordinate(start)
		self._end = Range.to_coordinate(end)
		self._parent = worksheet
	
	@property
	def x(self):
		return self.coordinate[0]
	
	@property
	def y(self):
		return self.coordinate[1]
	
	@property
	def coordinate(self):
		if self._start == self._end:
			return self._start
		else:
			raise Exception("Non-singleton range selected")
	
	def is_cell(self):
		return self._start == self._end
	
	def is_row(self):
		return self._start[0] == self._end[0]
			and self._start[1] == 1
			and self._end[1] == self._parent.num_columns
		
	def is_column(self):
		return self._start[1] == self._end[1]
			and self._start[0] == 1
			and self._end[1] == self._parent.num_rows
	
	def __getitem__(self, key):
		if self.is_row():
			# return the key'th column
			newStart = (self._start[0], key)
			newEnd = (self._end[0], key)
			return Range(newStart, newEnd, self._parent)
		elif self.is_column():
			#return the key'th row
			newStart = (key, self._start[1])
			newEnd = (key, self._end[1])
			return Range(newStart, newEnd, self._parent)			
		else:
			raise Exception("Selection not valid")
	
	def __setitem__(self, key, value):
		
	
	@property
	def value(self):
		return self._parent[x][y]
		
	@value.setter
	def value(self, value):
		for i in range(self._start[0], self._end[0] + 1):
			for j in range(self._start[1], self._end[1] + 1):
				self.
		return self
	
	@staticmethod
	def __name_to_integer(str):
		# Convert a base-26 name to integer
		int = 0
		for i in range(len(str)):
			int *= 26
			if ord(str[i]) < A or ord(str[i]) > Z:
				break # done, hopefully
			int += ord(str[i]) - A + 1
		return int

	@staticmethod
	def to_coordinate(value):
		if isinstance(value, basestring):
			value = Range.__name_to_integer(value)
		return value
