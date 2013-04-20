import Worksheet
from Writer import Writer

class Workbook(object):
    def __init__(self, encoding='utf-8'):
        self._worksheets = []
        self._encoding = encoding
        self._writer = Writer(self)

    def add_sheet(self, worksheet):
        self._worksheets.append(worksheet)
        
    def new_sheet(self, sheet_name):
        worksheet = Worksheet.Worksheet(sheet_name, self)
        self._worksheets.append(worksheet)
        return worksheet

    def get_xml_data(self):
        for index, ws in enumerate(self._worksheets, 1):
            yield (index, ws)

    def __len__(self):
        return len(self._worksheets)

    def _save(self, file_handle):
        self._writer.save(file_handle)

    def save(self, filename):
        self._save(open(filename, 'w'))
