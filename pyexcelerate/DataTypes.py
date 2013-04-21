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
		# Using value.__class__ over isinstance for speed
		if value.__class__ == str:
			if len(value) > 0 and value[0] == '=':
				return DataTypes.FORMULA
			else:
				return DataTypes.INLINE_STRING
		# not using in (int, float, long, complex) for speed
		elif value.__class__ == int or value.__class__ == float or value.__class__ == long or value.__class__ == complex:
			return DataTypes.NUMBER
		elif value.__class__ == datetime:
			return DataTypes.DATE
		else:
			return DataTypes.ERROR