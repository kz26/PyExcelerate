from . import Font, Fill, Format, Alignment, Borders
from . import Utility
import six

class Style(object):
	def __init__(self, font=None, fill=None, format=None, alignment=None, borders=None, size=None):
		self._font = font
		self._fill = fill
		self._format = format
		self._alignment = alignment
		self._borders = borders
		self._size = size

	@property
	def size(self):
		return self._size		
	
	@property
	def is_default(self):
		return not (self._font or self._fill or self._format or self._alignment or self._borders or self._size is not None)

	@property
	def borders(self):
		return Utility.lazy_get(self, '_borders', Borders.Borders())
		
	@borders.setter
	def borders(self, value):
		Utility.lazy_set(self, '_borders', None, value)
		
	@property
	def alignment(self):
		return Utility.lazy_get(self, '_alignment', Alignment.Alignment())
		
	@alignment.setter
	def alignment(self, value):
		Utility.lazy_set(self, '_alignment', None, value)
		
	@property
	def format(self):
		# don't use default because default should be const
		return Utility.lazy_get(self, '_format', Format.Format())
	
	@format.setter
	def format(self, value):
		Utility.lazy_set(self, '_format', None, value)
	
	@property
	def font(self):
		return Utility.lazy_get(self, '_font', Font.Font())
	
	@font.setter
	def font(self, value):
		Utility.lazy_set(self, '_font', None, value)
	
	@property
	def fill(self):
		return Utility.lazy_get(self, '_fill', Fill.Fill())
	
	@fill.setter
	def fill(self, value):
		Utility.lazy_set(self, '_fill', None, value)
	
	def get_xml_string(self):
		# Precondition: Workbook._align_styles has been run.
		# Be careful when using this function as id's may be inaccurate if precondition not met.
		tag = []
		if not self._format is None:
			tag.append("numFmtId=\"%d\"" % self._format.id)
		if not self._font is None:
			tag.append("applyFont=\"1\" fontId=\"%d\"" % (self._font.id))
		if not self._fill is None:
			tag.append("applyFill=\"1\" fillId=\"%d\"" % (self._fill.id + 1))
		if not self._borders is None:
			tag.append("applyBorder=\"1\" borderId=\"%d\"" % (self._borders.id))
		if self._alignment is None:
			return "<xf xfId=\"0\" %s/>" % (" ".join(tag))
		else:
			return "<xf xfId=\"0\"  %s applyAlignment=\"1\">%s</xf>" % (" ".join(tag), self._alignment.get_xml_string())
		
	def __hash__(self):
		return hash((self._font, self._fill))
	
	def __eq__(self, other):
		if other is None:
			return self.is_default
		elif Utility.YOLO:
			return self._format == other._format and self._alignment == other._alignment and self._borders == other._borders
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
			font=operation(self.font, other.font), \
			fill=operation(self.fill, other.fill), \
			format=operation(self.format, other.format), \
			alignment=operation(self.alignment, other.alignment), \
			borders=operation(self.borders, other.borders) \
		)
	
	def _to_tuple(self):
		return (self._font, self._fill, self._format, self._alignment, self._borders, self._size)
	
			
	def __str__(self):
		return "%s %s %s %s" % (self.font, self.fill, self.format, self.alignment)
		
	def __repr__(self):
		return "<%s>" % self.__str__()
