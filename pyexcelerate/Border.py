from . import Utility
from . import Color

#
# An object representing a single border
#
class Border(object):
	STYLE_MAPPING = { \
		'dashDot': ('.-', '-.', 'dash dot'), \
		'dashDotDot': ('..-', '-..', 'dash dot dot'), \
		'dashed': ('--'), \
		'dotted': ('..', ':'), \
		'double': ('='), \
		'hair': ('hairline', '.'), \
		'medium': (), \
		'mediumDashDot': ('medium dash dot', 'medium -.', 'medium .-'), \
		'mediumDashDotDot': ('medium dash dot dot', 'medium -..', 'medium ..-'), \
		'mediumDashed': ('medium dashed', 'medium --'), \
		'slantDashDot': ('/-.', 'slant dash dot'), \
		'thick': (), \
		'thin': ('_') \
	}
	
	def __init__(self, color=None, style='thin'):
		self._color = color
		self._style = Border.get_style_name(style)
	
	@property
	def color(self):
		return Utility.lazy_get(self, '_color', Color.Color(0, 0, 0))
	
	@color.setter
	def color(self, value):
		Utility.lazy_set(self, '_color', None, value)
		
	@property
	def style(self):
		return self._style
		
	@style.setter
	def style(self, value):
		self._style = Border.get_style_name(value)
	
	@staticmethod
	def get_style_name(style):
		for key, values in Border.STYLE_MAPPING.items():
			if style == key or style in values:
				return key
		# TODO: warn the user?
		return 'thin'
	
	@property
	def is_default(self):
		return self._color is None and self._style == 'thin'
	
	def __eq__(self, other):
		if other is None:
			return self.is_default
		else:
			return self._color == other._color and self._style == other._style
			
	def __hash__(self):
		return hash(self._style)