import DataTypes

class Range(object):
	A = ord('A')
	Z = ord('Z')
	def __init__(self, start, end, worksheet):
		self._start = Range.to_coordinate(start)
		self._end = Range.to_coordinate(end)
		self._parent = worksheet
		if self.is_cell():
			# Report the column as active so the worksheet knows
			self.worksheet.report_column(self.y)
	
	@property
	def x(self):
		if self.is_row():
			return self._start[0]
		else:
			return self.coordinate[0]
	
	@property
	def y(self):
		if self.is_column():
			return self._start[1]
		else:
			return self.coordinate[1]
			
	@property
	def coordinate(self):
		if self.is_cell():
			return self._start
		else:
			raise Exception("Non-singleton range selected")
	
	@property
	def value(self):
		if self.is_cell():
			return self.worksheet.get_cell_value(self.x, self.y)
		else:
			raise Exception("Not a cell")

	@value.setter
	def value(self, value):
		if self.is_cell():
			self.worksheet.set_cell_value(self.x, self.y, value)
		else:
			raise Exception("Not a cell")
			
	@property
	def key(self):
		if self.is_cell():
			return self.workbook.worksheet.shared_strings.get_key(self.value)
		else:
			raise Exception("Not a cell")
	
	@property
	def worksheet(self):
		return self._parent
	
	def is_cell(self):
		return self._start == self._end
	
	def is_row(self):
		return self._start[0] == self._end[0] \
			and self._start[1] == 1 \
			and self._end[1] == float('inf')
		
	def is_column(self):
		return self._start[1] == self._end[1] \
			and self._start[0] == 1 \
			and self._end[1] == float('inf') \
	
	def intersection(self, range):
		"""
		Calculates the intersection with another range object
		"""
		if self.worksheet != range.worksheet:
			# Different worksheet
			return None
		start = (max(self._start[0], range._start[0]), max(self._start[1], range._start[1]))
		end = (min(self._end[0], range._end[0]), min(self._end[1], range._end[1]))
		return Range(start, end, self.worksheet)
		
	def intersects(self, range):
		return intersection(range) == None
	
	def merge(self):
		self.worksheet.add_merge(self)
	
	def __getitem__(self, key):
		if self.is_row():
			# return the key'th column
			newStart = (self.x, key)
			newEnd = (self.x, key)
			return Range(newStart, newEnd, self.worksheet)
		elif self.is_column():
			#return the key'th row
			newStart = (key, self.y)
			newEnd = (key, self.y)
			return Range(newStart, newEnd, self.worksheet)			
		else:
			raise Exception("Selection not valid")
	
	def __setitem__(self, key, value):
		if self.is_row():
			self.worksheet.set_cell_value(self.x, key, value)
		else:
			raise Exception("Couldn't set that")
	
	@staticmethod
	def __string_to_coordinate(s):
		# Convert a base-26 name to integer
		y = 0
		for i in range(len(s)):
			if ord(s[i]) < Range.A or ord(s[i]) > Range.Z:
				s = s[i:]
				break
			y *= 26
			y += ord(s[i]) - Range.A + 1
		return (int(s), y)

	@staticmethod
	def __coordinate_to_string(coord):
		# convert an integer to base-26 name
		y = coord[1]
		s = ""
		while y > 0:
			s = chr((y % 26) + Range.A - 1) + s
			y -= (y % 26)
			y /= 26
		
		return s + str(coord[0])
	
	@staticmethod
	def to_coordinate(value):
		if isinstance(value, basestring):
			value = Range.__string_to_coordinate(value)
		return value
	
	def get_xml_data(self):
		if self.is_row():
			for index, cell in self.worksheet._cells[self.x].iteritems():
				yield (Range.__coordinate_to_string((self.x, index)), cell, DataTypes.DataTypes.get_type(cell))
		else:
			raise Exception("not a valid row")
