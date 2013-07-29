from datetime import datetime
from . import six

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
		if value.__class__ in six.string_types:
			if len(value) > 0 and value[0] == '=':
				return DataTypes.FORMULA
			else:
				return DataTypes.INLINE_STRING
		# not using in (int, float, long, complex) for speed
		elif value.__class__ in six.integer_types + (float, complex):
			return DataTypes.NUMBER
		# fall back to the slower isinstance
		elif isinstance(value, six.string_types):
			if len(value) > 0 and value[0] == '=':
				return DataTypes.FORMULA
			else:
				return DataTypes.INLINE_STRING
		elif isinstance(value, six.integer_types + (float, complex)):
			return DataTypes.NUMBER
		elif isinstance(value, datetime):
			return DataTypes.DATE
		else:
			return DataTypes.ERROR