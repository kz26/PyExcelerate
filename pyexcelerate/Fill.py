from . import Color

class Fill(object):
	def __init__(self):
		self.background = Color.Color()

	@property
	def is_default(self):
		return self == Fill()

	def __eq__(self, other):
		if other is None:
			return self.is_default
		else:
			return self.background == other.background

	def __hash__(self):
		return hash(self.background)
		
	def get_xml_string(self):
		if self.background == Color.Color.TRANSPARENT:
			return '<fill><patternFill patternType="none"/></fill>'
		else:
			return "<fill><patternFill patternType=\"solid\"><fgColor rgb=\"%s\"/></patternFill></fill>" % self.background.hex
	
	def __str__(self):
		return "Fill: #%s" % self.background.hex
	
	def __repr__(self):
		return "<%s>" % self.__str__()