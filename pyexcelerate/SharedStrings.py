# NB: Not actually used atm

class SharedStrings(object):
	def __init__(self, workbook):
		self._parent = workbook
		self._map = {}
		self._index = 1
		
	@property
	def workbook(self):
		return self._parent
		
	def get_key(self, s):
		# get the key for s
		if s not in self._map:
			self._map[s] = self._index
			self._index += 1
		return self._map[s]
		