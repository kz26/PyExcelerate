import time
from datetime import datetime

number = 23234

stime = time.clock()
for i in range(1000000):
	v = number.__class__ == int or number.__class__ == long or number.__class__ == float or number.__class__ == complex
	if not v:
		print "incorrect answer"
		print number.__class__
		break
print "isinstance, %s" % (time.clock() - stime)



stime = time.clock()
for i in range(1000000):
	v = number.__class__ in (int, long, float, complex)
	if not v:
		print "incorrect answer"
		print number.__class__
		break
print "__class__, %s" % (time.clock() - stime)