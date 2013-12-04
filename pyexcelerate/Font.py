import six
from .Utility import Utility
from . import Color

class Font(object):
	def __init__(self, bold=False, italic=False, underline=False, strikethrough=False, family='Calibri', size=11, color=None):
		self.bold = bold
		self.italic = italic
		self.underline = underline
		self.strikethrough = strikethrough
		self.family = family
		self.size = size
		self._color = color
	
	def get_xml_string(self):
		tokens = ["<sz val=\"%d\"/><name val=\"%s\"/>" % (self.size, self.family)]
		# sure, we could do this with an enum, but this is faster :D
		if self.bold:
			tokens.append('<b/>')
		if self.italic:
			tokens.append('<i/>')
		if self.underline:
			tokens.append('<u/>')
		if self.strikethrough:
			tokens.append('<strike/>')
		if self._color:
			tokens.append("<color rgb=\"%s\"/>" % self._color.hex)
		return "<font>%s</font>" % "".join(tokens)

	@property
	def color(self):
		return Utility.lazy_get(self, '_color', Color.Color())
		
	@color.setter
	def color(self, value):
		Utility.lazy_set(self, '_color', Color.Color(), value)
		
	@property
	def is_default(self):
		return self == Font()

	def __or__(self, other):
		return self._binary_operation(other, Utility.nonboolean_or)

	def __and__(self, other):
		return self._binary_operation(other, Utility.nonboolean_and)
	
	def __xor__(self, other):
		return self._binary_operation(other, Utility.nonboolean_xor)
	
	def _binary_operation(self, other, operation):
		return Font( \
			bold = operation(self.bold, other.bold), \
			italic = operation(self.italic, other.italic), \
			underline = operation(self.underline, other.underline), \
			strikethrough = operation(self.strikethrough, other.strikethrough), \
			family = operation(self.family, other.family, 'Calibri'), \
			size = operation(self.size, other.size, 11), \
			color = operation(self._color, other._color, None)
		)

	def __eq__(self, other):
		if other is None:
			return self.is_default
		else:
			return self._to_tuple() == other._to_tuple()

	def __hash__(self):
		return hash(self._to_tuple())

	def _to_tuple(self):
		return (self.bold, self.italic, self.underline, self.strikethrough, self.family, self.size)

	def __str__(self):
		tokens = ["%s, %dpt" % (self.family, self.size)]
		# sure, we could do this with an enum, but this is faster :D
		if self.bold:
			tokens.append('b')
		if self.italic:
			tokens.append('i')
		if self.underline:
			tokens.append('u')
		if self.strikethrough:
			tokens.append('s')
		return "Font: %s" % ' '.join(tokens)
	
	def __repr__(self):
		return "<%s>" % self.__str__()
		
