from . import six
from . import Font
from . import Fill

class Style(object):
	def __init__(self):
		self.id = -1 # set by the stylesheet object when needed
		self.font = Font.Font()
		self.fill = Fill.Fill()
		self.format = None

	@property
	def numFmtId(self):
		if self.format == None:
			return 0
		else:
			return 1000 + self.id
