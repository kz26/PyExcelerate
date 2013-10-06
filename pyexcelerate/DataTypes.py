from datetime import datetime, date, time
from . import six
try:
	import numpy as np
	HAS_NUMPY = True
except:
	HAS_NUMPY = False

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
	_numberTypes = six.integer_types + (float, complex)

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
		elif value.__class__ in DataTypes._numberTypes:
			return DataTypes.NUMBER
		# fall back to the slower isinstance
		elif isinstance(value, six.string_types):
			if len(value) > 0 and value[0] == '=':
				return DataTypes.FORMULA
			else:
				return DataTypes.INLINE_STRING
		elif isinstance(value, DataTypes._numberTypes):
			return DataTypes.NUMBER
		elif HAS_NUMPY and isinstance(value, (np.floating, np.integer, np.complexfloating, np.unsignedinteger)):
			return DataTypes.NUMBER
		elif isinstance(value, (datetime, date, time)):
			return DataTypes.DATE
		else:
			return DataTypes.ERROR

