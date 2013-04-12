import Worksheet

class Workbook(object):
	def __init__(self, encoding='utf-8'):
		self._worksheets = {}
		self._encoding = encoding
		
	def create_sheet(self, sheet_name):
		worksheet = Worksheet.Worksheet(sheet_name, self)
		self._worksheets.append(worksheet)
		return worksheet

	def get_sheet(self, sheet):
		return self._worksheets[sheet]