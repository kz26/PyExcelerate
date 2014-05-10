from datetime import datetime, date, time
import decimal
import six
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
	EXCEL_BASE_DATE = datetime(1900, 1, 1, 0, 0, 0)
	
	_numberTypes = six.integer_types + (float, complex, decimal.Decimal)
		
	@staticmethod
	def get_type(value):
		# Using value.__class__ over isinstance for speed
		if value.__class__ in six.string_types:
			if len(value) > 0 and value[0] == '=':
				return DataTypes.FORMULA
			else:
				return DataTypes.INLINE_STRING
		# not using in (int, float, long, complex) for speed
		elif value.__class__ == bool:
			return DataTypes.BOOLEAN
		elif value.__class__ in DataTypes._numberTypes:
			return DataTypes.NUMBER
		# fall back to the slower isinstance
		elif isinstance(value, six.string_types):
			if len(value) > 0 and value[0] == '=':
				return DataTypes.FORMULA
			else:
				return DataTypes.INLINE_STRING
		elif isinstance(value, bool):
			return DataTypes.BOOLEAN
		elif isinstance(value, DataTypes._numberTypes):
			return DataTypes.NUMBER
		elif HAS_NUMPY and isinstance(value, (np.floating, np.integer, np.complexfloating, np.unsignedinteger)):
			return DataTypes.NUMBER
		elif isinstance(value, (datetime, date, time)):
			return DataTypes.DATE
		else:
			return DataTypes.ERROR
	
	@staticmethod
	def to_excel_date(d):
		if isinstance(d, datetime):
			delta = d - DataTypes.EXCEL_BASE_DATE
			excel_date = delta.days + (float(delta.seconds) + float(delta.microseconds) / 1E6) / (60 * 60 * 24) + 1
			return excel_date + (excel_date > 59)
		elif isinstance(d, date):
			# this is why python sucks >.<
			return DataTypes.to_excel_date(datetime(*(d.timetuple()[:6])))
		elif isinstance(d, time):
			return DataTypes.to_excel_date(datetime(*(DataTypes.EXCEL_BASE_DATE.timetuple()[:3]), hour=d.hour, minute=d.minute, second=d.second, microsecond=d.microsecond)) - 1
