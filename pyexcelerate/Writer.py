import os
import sys
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED
from datetime import datetime
import time
from jinja2 import Environment, FileSystemLoader
from . import Color

if getattr(sys, 'frozen', False):
	_basedir = os.path.join(sys._MEIPASS, 'pyexcelerate')
else:
	_basedir = os.path.dirname(__file__)
_TEMPLATE_PATH = os.path.join(_basedir, 'templates')

class Writer(object):
	env = Environment(loader=FileSystemLoader(_TEMPLATE_PATH), auto_reload=False)
	_docProps_app_template = env.get_template("docProps/app.xml")
	_docProps_core_template = env.get_template("docProps/core.xml")
	_content_types_template = env.get_template("[Content_Types].xml")
	_rels_template = env.get_template("_rels/.rels")
	_styles_template = env.get_template("xl/styles.xml") 
	_empty_styles_template = env.get_template("xl/styles.empty.xml") 
	_workbook_template = env.get_template("xl/workbook.xml")
	_workbook_rels_template = env.get_template("xl/_rels/workbook.xml.rels")
	_worksheet_template = env.get_template("xl/worksheets/sheet.xml")

	def __init__(self, workbook):
		self.workbook = workbook

	def _render_template_wb(self, template, extra_context=None):
		context = {'workbook': self.workbook}
		if extra_context:
			context.update(extra_context)
		return template.render(context).encode('utf-8')

	def _get_utc_now(self):
		now = datetime.utcnow()
		return now.strftime("%Y-%m-%dT%H:%M:00Z")


	def save(self, f):
		zf = ZipFile(f, 'w', ZIP_DEFLATED)
		zf.writestr("docProps/app.xml", self._render_template_wb(self._docProps_app_template))
		zf.writestr("docProps/core.xml", self._render_template_wb(self._docProps_core_template, {'date': self._get_utc_now()}))
		zf.writestr("[Content_Types].xml", self._render_template_wb(self._content_types_template))
		zf.writestr("_rels/.rels", self._rels_template.render().encode('utf-8'))
		if self.workbook.has_styles:
			zf.writestr("xl/styles.xml", self._render_template_wb(self._styles_template))
		else:
			zf.writestr("xl/styles.xml", self._render_template_wb(self._empty_styles_template))
		zf.writestr("xl/workbook.xml", self._render_template_wb(self._workbook_template))
		zf.writestr("xl/_rels/workbook.xml.rels", self._render_template_wb(self._workbook_rels_template))
		for index, sheet in self.workbook.get_xml_data():
			tfd, tfn = tempfile.mkstemp()
			tf = os.fdopen(tfd, 'wb')
			sheetStream = self._worksheet_template.generate({'worksheet': sheet})
			for s in sheetStream:
				tf.write(s.encode('utf-8'))
			tf.close()
			zf.write(tfn, "xl/worksheets/sheet%s.xml" % (index))
			os.remove(tfn)
		zf.close()
