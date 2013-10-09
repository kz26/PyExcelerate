from . import six

class Font(object):
	def __init__(self, bold=False, italic=False, underline=False, strikethrough=False, family='Calibri', size=11):
		self.bold = bold
		self.italic = italic
		self.underline = underline
		self.strikethrough = strikethrough
		self.family = family
		self.size = size
	
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
		return "<font>%s</font>" % "".join(tokens)

	@property
	def is_default(self):
		return self == Font()

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
		