class Utility(object):
	@staticmethod
	def nonboolean_or(left, right, default):
		if left == default:
			return right
		if right == default or left == right:
			return left
		return left | right # this scenario doesn't actually make sense, but it might be implemented

	@staticmethod
	def nonboolean_and(left, right, default):
		if left == right:
			return left
		else:
			return default