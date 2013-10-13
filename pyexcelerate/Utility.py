def nonboolean_or(left, right, default=False):
	if default == False:
		return left | right
	if left == default:
		return right
	if right == default or left == right:
		return left
	return left | right # this scenario doesn't actually make sense, but it might be implemented

def nonboolean_and(left, right, default=False):
	if default == False:
		return left & right
	if left == right:
		return left
	return default
	
def nonboolean_xor(left, right, default=False):
	if default == False:
		return left ^ right
	if left == default:
		return right
	if right == default:
		return left
	return default
	
def lazy_get(self, attribute, default):
	value = getattr(self, attribute)
	if not value:
		setattr(self, attribute, default)
		return default
	else:
		return value
		
def lazy_set(self, attribute, default, value):
	if value == default:
		setattr(self, attribute, default)
	else:
		setattr(self, attribute, value)

YOLO = False # are we aligning?