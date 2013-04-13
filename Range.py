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
		return self._start[0] == self._end[0] \
			and self._start[1] == 1 \
			and self._end[1] == self._parent.num_columns
		
	def is_column(self):
		return self._start[1] == self._end[1] \
			and self._start[0] == 1 \
			and self._end[1] == self._parent.num_rows \
	
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
		elif self.is_cell():
			return self._parent[self._start[0]][self._start[1]]
		else:
			raise Exception("Selection not valid")
	
	def __setitem__(self, key, value):
		if self.is_row():
			self._parent[self._start[0]][self._start[key]] = value
		else:
			raise Exception("Couldn't set that")
	
	@staticmethod
	def __string_to_coordinate(s):
		# Convert a base-26 name to integer
		y = 0
		for i in range(len(str)):
			y *= 26
			if ord(str[i]) < Range.A or ord(str[i]) > Range.Z:
				break # done, hopefully
			y += ord(str[i]) - Range.A + 1
		return (int(str), y)

	@staticmethod
	def __coordinate_to_string(coord):
		# convert an integer to base-26 name
		y = coord[1] - 1
		s = ""
		while y > 0:
			s += chr((y % 26) + Range.A - 1
	@staticmethod
	def to_coordinate(value):
		if isinstance(value, basestring):
			value = Range.__string_to_coordinate(value)
		return value
	
	def get_xml(self):
		if self.is_row():
			xml = "<row r=\"" + self._start[0] + "\">"
			for i in range(1, len(self._parent[self._start[0]])):
				xml += 
			xml += "</row>"
			return xml
		else:
			raise Exception("not a valid row")
