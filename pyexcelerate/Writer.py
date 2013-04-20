import os
import tempfile
from zipfile import ZipFile, ZIP_DEFLATED

from jinja2 import Environment, FileSystemLoader

class Writer(object):
    TEMPLATE_PATH = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'templates')
    env = Environment(loader=FileSystemLoader(TEMPLATE_PATH), auto_reload=False)

    _docProps_app_template = env.get_template("docProps/app.xml")
    _docProps_core_template = env.get_template("docProps/core.xml")
    _content_types_template = env.get_template("[Content_Types].xml")
    _rels_template = env.get_template("_rels/.rels")
    _workbook_template = env.get_template("xl/workbook.xml")
    _workbook_rels_template = env.get_template("xl/_rels/workbook.xml.rels")
    _worksheet_template = env.get_template("xl/worksheets/sheet.xml")

    def __init__(self, workbook):
        self.workbook = workbook

    def _render_template_wb(self, template):
        return template.render({'workbook': self.workbook})

    def save(self, f):
        zf = ZipFile(f, 'w', ZIP_DEFLATED)
        zf.writestr("docProps/app.xml", self._render_template_wb(self._docProps_app_template))
        zf.writestr("docProps/core.xml", self._render_template_wb(self._docProps_core_template))
        zf.writestr("[Content_Types].xml", self._render_template_wb(self._content_types_template))
        zf.writestr("_rels/.rels", self._rels_template.render())
        zf.writestr("xl/workbook.xml", self._render_template_wb(self._workbook_template))
        zf.writestr("xl/_rels/workbook.xml.rels", self._render_template_wb(self._workbook_rels_template))
        for index, sheet in self.workbook.get_xml_data():
            tf = tempfile.NamedTemporaryFile()
            sheetStream = self._worksheet_template.generate({'worksheet': sheet})
            for s in sheetStream:
                tf.write(s)
            tf.seek(0)
            zf.write(tf.name, "xl/worksheets/sheet%s.xml" % (index))
            tf.close()
        zf.close()
