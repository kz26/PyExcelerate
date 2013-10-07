import time
from nose.tools import *

def test_benchmark():
	TRIALS = range(1000000)
	
	integer = 1
	float = 3.0
	long = 293203948032948023984023948023957245
	
	# attempt isinstance
	
	stime = time.clock()
	for i in TRIALS:
		answer = isinstance(integer, (int, float, long, complex))
		ok_(answer)
	print("isinstance, %s" % (time.clock() - stime))

	# attempt __class__
	
	stime = time.clock()
	for i in TRIALS:
		answer = (integer.__class__ in set((int, float, long, complex)))
		ok_(answer)
	print("__class__, set, %s" % (time.clock() - stime))
	
	stime = time.clock()
	for i in TRIALS:
		answer = (integer.__class__ == int or integer.__class__ == float or integer.__class__ == long or integer.__class__ == complex)
		ok_(answer)
	print("__class__, or, %s" % (time.clock() - stime))
	
