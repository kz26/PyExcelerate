import Worksheet

class Workbook(object):
	def __init__(self, encoding='utf-8'):
		self._worksheets = []
		self._encoding = encoding
		
	def create_sheet(self, sheet_name):
		worksheet = Worksheet.Worksheet(sheet_name, self)
		self._worksheets.append(worksheet)
		return worksheet

	def get_xml_data(self):
		data = []
		for i in range(len(self._worksheets)):
			data.append((i + 1, self._worksheets[i].name))
		return data