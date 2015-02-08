import six
from . import Utility
from . import Color

class Alignment(object):
	def __init__(self, horizontal='left', vertical='bottom', rotation=0, wrap_text=False):
		self._horizontal = horizontal
		self._vertical = vertical
		self._rotation = rotation
		self._wrap_text = wrap_text
	
	@property
	def wrap_text(self):
		return self._wrap_text
	
	@wrap_text.setter
	def wrap_text(self, value):
		if value not in (True, False):
			raise TypeError('Invalid wrap text alignment value. Expects either True or False.')
		self._wrap_text = value
	
	@property
	def horizontal(self):
		return self._horizontal
	
	@horizontal.setter
	def horizontal(self, value):
		if value not in ('left', 'center', 'right'):
			raise ValueError('Invalid horizontal alignment value. Expects either \'left\', \'center\', or \'right\'.')
		self._horizontal = value
	
	@property
	def vertical(self):
		return self._vertical
	
	@vertical.setter
	def vertical(self, value):
		if value not in ('top', 'center', 'bottom'):
			raise ValueError('Invalid vertical alignment value. Expects either \'top\', \'center\', or \'bottom\'.')
		self._vertical = value
	
	@property
	def rotation(self):
		return self._rotation
	
	@rotation.setter
	def rotation(self, value):
		self._rotation = (value % 360)
	
	@property
	def is_default(self):
		return self._horizontal == 'left' and self._vertical == 'bottom' and self._rotation == 0 and not self._wrap_text

	def get_xml_string(self):
		return "<alignment horizontal=\"%s\" vertical=\"%s\" textRotation=\"%.15g\" wrapText=\"%d\"/>" % (self._horizontal, self._vertical, self._rotation, 1 if self._wrap_text else 0)

	def __or__(self, other):
		return self._binary_operation(other, Utility.nonboolean_or)

	def __and__(self, other):
		return self._binary_operation(other, Utility.nonboolean_and)
	
	def __xor__(self, other):
		return self._binary_operation(other, Utility.nonboolean_xor)
	
	def _binary_operation(self, other, operation):
		return Alignment( \
			horizontal = operation(self._horizontal, other._horizontal, 'left'), \
			vertical = operation(self._vertical, other._vertical, 'bottom'), \
			rotation = operation(self._rotation, other._rotation, 0), \
			wrap_text = operation(self._wrap_text, other._wrap_text, False)
		)
	
	def __eq__(self, other):
		if other is None:
			return self.is_default
		elif Utility.YOLO:
			return self._vertical == other._vertical and self._rotation == other._rotation
		else:
			return self._vertical == other._vertical and self._rotation == other._rotation and self._horizontal == other._horizontal and self._wrap_text == other._wrap_text
	
	def __hash__(self):
		return hash((self._horizontal))
	
	def __str__(self):
		return "Align: %s %s %s" % (self._horizontal, self._vertical, self._rotation)
		
