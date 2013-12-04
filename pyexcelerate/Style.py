import six
from . import Font, Fill, Format
from .Utility import Utility

class Style(object):
	_DEFAULT_FORMAT = Format.Format()
	_DEFAULT_FILL = Fill.Fill()
	_DEFAULT_FONT = Font.Font()
	def __init__(self, font=None, fill=None, format=None):
		self._font = font
		self._fill = fill
		self._format = format

	@property
	def is_default(self):
		return not (self._font or self._fill or self._format)

	@property
	def format(self):
		# don't use default because default should be const
		return Utility.lazy_get(self, '_format', Format.Format())
	
	@format.setter
	def format(self, value):
		Utility.lazy_set(self, '_format', Style._DEFAULT_FORMAT, value)
	
	@property
	def font(self):
		return Utility.lazy_get(self, '_font', Font.Font())
	
	@font.setter
	def font(self, value):
		Utility.lazy_set(self, '_font', Style._DEFAULT_FONT, value)
	
	@property
	def fill(self):
		return Utility.lazy_get(self, '_fill', Fill.Fill())
	
	@fill.setter
	def fill(self, value):
		Utility.lazy_set(self, '_fill', Style._DEFAULT_FILL, value)
	
	def get_xml_string(self):
		# Precondition: Workbook._align_styles has been run.
		# Be careful when using this function as id's may be inaccurate if precondition not met.
		tag = []
		if not self.format.is_default:
			tag.append("numFmtId=\"%d\"" % self.format.id)
		if not self.font.is_default:
			tag.append("applyFont=\"1\" fontId=\"%d\"" % (self.font.id))
		if not self.fill.is_default:
			tag.append("applyFill=\"1\" fillId=\"%d\"" % (self.fill.id + 1))
		return "<xf xfId=\"0\" borderId=\"0\" %s/>" % (" ".join(tag))
		
	def __hash__(self):
		return hash((hash(self._font), hash(self._fill), hash(self._format)))
	
	def __eq__(self, other):
		if other is None:
			return self.is_default
		else:
			return self._to_tuple() == other._to_tuple()
	
	def __or__(self, other):
		return self._binary_operation(other, Utility.nonboolean_or)
			
	def __and__(self, other):
		return self._binary_operation(other, Utility.nonboolean_and)

	def __xor__(self, other):
		return self._binary_operation(other, Utility.nonboolean_xor)

	def _binary_operation(self, other, operation):
		return Style( \
			font=operation(self._font, other._font, None), \
			fill=operation(self._fill, other._fill, None), \
			format=operation(self._format, other._format, None))
	
	def _to_tuple(self):
		return (self._font, self._fill, self._format)
	
			
	def __str__(self):
		return "%s %s %s" % (self.font, self.fill, self.format)
		
	def __repr__(self):
		return "<%s>" % self.__str__()
