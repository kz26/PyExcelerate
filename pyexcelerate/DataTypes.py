from datetime import datetime

class DataTypes(object):
	BOOLEAN = 0
	DATE = 1
	ERROR = 2
	INLINE_STRING = 3
	NUMBER = 4
	SHARED_STRING = 5
	STRING = 6
	FORMULA = 7
	
	_enumerations = ["b", "d", "e", "inlineStr", "n", "s", "str"]
	@staticmethod
	def to_enumeration_value(index):
		return DataTypes._enumerations[index]
		
	@staticmethod
	def get_type(value):
		if isinstance(value, basestring):
			if len(value) > 0 and value[0] == '=':
				return DataTypes.FORMULA
			else:
				return DataTypes.INLINE_STRING
		elif isinstance(value, (int, long, float, complex)):
			return DataTypes.NUMBER
		elif isinstance(value, (datetime)):
			return DataTypes.DATE
		else:
			return DataTypes.ERROR