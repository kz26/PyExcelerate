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
		return self.coordinate[0]
	
	@property
	def y(self):
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
			and self._end[1] == None
		
	def is_column(self):
		return self._start[1] == self._end[1] \
			and self._start[0] == 1 \
			and self._end[1] == None \
	
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
		if self.is_row():
			self._parent[self._start[0]][self._start[key]] = value
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
			for index, cell in self._parent[self._start[0]]:
				yield (Range.__coordinate_to_string((self._start[0], index)), cell.value, DataTypes.DataTypes.get_type(cell.value))
		else:
			raise Exception("not a valid row")
