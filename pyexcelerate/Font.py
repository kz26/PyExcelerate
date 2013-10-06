from . import six

class Font(object):
	def __init__(self, bold=False, italic=False, underline=False, strikethrough=False, family='Calibri', size=11):
		self.bold = bold
		self.italic = italic
		self.underline = underline
		self.strikethrough = strikethrough
		self.family = family
		self.size = size
	