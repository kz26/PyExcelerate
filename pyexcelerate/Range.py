from . import DataTypes
from . import six
from . import Font, Fill, Format
from six.moves import reduce

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
	def height(self):
		return self._end[0] - self._start[0] + 1
	
	@property
	def width(self):
		return self._end[1] - self._start[1] + 1

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
	def style(self):
		return self.__get_attr(self.worksheet.get_cell_style)
		
	@style.setter
	def style(self, data):
		self.__set_attr(self.worksheet.set_cell_style, data)

	@property
	def value(self):
		return self.__get_attr(self.worksheet.get_cell_value)
		
	@value.setter
	def value(self, data):
		self.__set_attr(self.worksheet.set_cell_value, data)
	
	# this class permits doing things like range().font.bold = True
	class AttributeInterceptor(object):
		def __init__(self, parent, attribute):
			self.__dict__['_parent_range'] = parent
			self.__dict__['_attribute'] = attribute
		def __setattr__(self, name, value):
			for cell in self._parent_range:
				setattr(reduce(getattr, self._attribute.split('.'), cell), name, value)

	def __getattr__(self, name):
		# handles .format, .font, .fill
		return Range.AttributeInterceptor(self, "style.%s" % name)

	# note that these are not the python __getattr__/__setattr__
	def __get_attr(self, method):
		if self.is_cell():
			for merge in self.worksheet.merges:
				if self in merge:
					return method(merge._start[0], merge._start[1])
			return method(self.x, self.y)
		else:
			raise Exception("Not a cell")
	
	def __set_attr(self, method, data):
		if self.is_cell():
			for merge in self.worksheet.merges:
				if self in merge:
					method(merge._start[0], merge._start[1], data)
					return
			method(self.x, self.y, data)
		elif DataTypes.DataTypes.get_type(data) != DataTypes.DataTypes.ERROR:
			for cell in self:
				cell.__set_attr(method, data)
		else:
			if len(data) <= self.height:
				for row in data:
					if len(row) > self.width:
						raise Exception("Row too large for range, row has %s columns, but range only has %s" % (len(row), self.width))
				for x, row in enumerate(data):
					for y, value in enumerate(row):
						method(x + self._start[0], y + self._start[1], value)
			else:
				raise Exception("Too many rows for range, data has %s rows, but range only has %s" % (len(data), self.height))
	
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

	def __iter__(self):
		for x in range(self._start[0], self._end[0] + 1):
			for y in range(self._start[1], self._end[1] + 1):
				yield Range((x, y), (x, y), self.worksheet)

	def __contains__(self, item):
		return self.intersection(item) == item

	def __hash__(self):
		def hash(val):
			return val[0] << 8 + val[1]
		return hash(self._start) << 24 + hash(self._end)

	def __str__(self):
		return Range.coordinate_to_string(self._start) + ":" + Range.coordinate_to_string(self._end)

	def __len__(self):
		if self._start[0] == self._end[0]:
			return self.width
		else:
			return self.height

	def __eq__(self, other):
		return self._start == other._start and self._end == other._end
	
	def __ne__(self, other):
		return not (self == other)
	
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
	def string_to_coordinate(s):
		# Convert a base-26 name to integer
		y = 0
		for i in range(len(s)):
			if ord(s[i]) < Range.A or ord(s[i]) > Range.Z:
				s = s[i:]
				break
			y *= 26
			y += ord(s[i]) - Range.A + 1
		return (int(s), y)

	_cts_cache = {}
	@staticmethod
	def coordinate_to_string(coord):
		# convert an integer to base-26 name
		y = coord[1] - 1
		if y not in Range._cts_cache:
			s = ""	
			while y >= 0:
				s = chr((y % 26) + Range.A) + s
				y = int(y / 26) - 1
			Range._cts_cache[y] = s
		return Range._cts_cache[y] + str(coord[0])
	
	@staticmethod
	def to_coordinate(value):
		if isinstance(value, six.string_types):
			value = Range.string_to_coordinate(value)
		if (value[0] < 1 or value[0] > 65536) and value[1] != float('inf'):
			raise Exception("Row index out of bounds")
		if (value[1] < 1 or value[1] > 256) and value[1] != float('inf'):
			raise Exception("Column index out of bounds")
		return value
