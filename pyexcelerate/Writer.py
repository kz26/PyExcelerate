import os
import sys
import tempfile
import zipfile
from datetime import datetime
import time
from jinja2 import Environment, FileSystemLoader
from . import Color

try:
		import zlib # We may need its compression method
		crc32 = zlib.crc32
except ImportError:
		zlib = None
		crc32 = binascii.crc32
		
if getattr(sys, 'frozen', False):
	_basedir = os.path.join(sys._MEIPASS, 'pyexcelerate')
else:
	_basedir = os.path.dirname(__file__)
_TEMPLATE_PATH = os.path.join(_basedir, 'templates')

class ZipFile(zipfile.ZipFile):
	def writeiter(self, zinfo_or_arcname, iter, compress_type=None):
		if not isinstance(zinfo_or_arcname, zipfile.ZipInfo):
			zinfo = zipfile.ZipInfo(filename=zinfo_or_arcname,
															date_time=time.localtime(time.time())[:6])
			zinfo.compress_type = self.compression
			if zinfo.filename[-1] == '/':
				zinfo.external_attr = 0o40775 << 16	 # drwxrwxr-x
				zinfo.external_attr |= 0x10					 # MS-DOS directory flag
			else:
				zinfo.external_attr = 0o600 << 16		 # ?rw-------
		else:
			zinfo = zinfo_or_arcname

		if not self.fp:
			raise RuntimeError(
						"Attempt to write to ZIP archive that was already closed")

		if compress_type is not None:
			zinfo.compress_type = compress_type

		zinfo.header_offset = self.fp.tell()		# Start of header bytes
		
		# Must overwrite CRC and sizes with correct data later
		zinfo.file_size = file_size = 0
		zinfo.CRC = CRC = 0
		zinfo.compress_size = compress_size = 0
		
		self._writecheck(zinfo)
		self._didModify = True
		
		# Compressed size can be larger than uncompressed size
		zip64 = self._allowZip64 and \
						zinfo.file_size * 1.05 > zipfile.ZIP64_LIMIT
		self.fp.write(zinfo.FileHeader(zip64))
		
		if zinfo.compress_type == zipfile.ZIP_DEFLATED:
			cmpr = zlib.compressobj(zlib.Z_DEFAULT_COMPRESSION,
					 zlib.DEFLATED, -15)
		else:
			cmpr = None
		
		for buf in iter:
			buf = buf.encode('utf-8')
			file_size = file_size + len(buf)
			CRC = crc32(buf, CRC) & 0xffffffff
			if cmpr:
				buf = cmpr.compress(buf)
				compress_size = compress_size + len(buf)
			self.fp.write(buf)
				
		if cmpr:
			buf = cmpr.flush()
			compress_size = compress_size + len(buf)
			self.fp.write(buf)
			zinfo.compress_size = compress_size
		else:
			zinfo.compress_size = file_size
				
		zinfo.CRC = CRC
		zinfo.file_size = file_size
		
		if not zip64 and self._allowZip64:
			if file_size > zipfile.ZIP64_LIMIT:
				raise RuntimeError('File size has increased during compressing')
			if compress_size > zipfile.ZIP64_LIMIT:
				raise RuntimeError('Compressed size larger than uncompressed size')
		# Seek backwards and write file header (which will now include
		# correct CRC and file sizes)
		position = self.fp.tell() # Preserve current position in file
		self.fp.seek(zinfo.header_offset, 0)
		self.fp.write(zinfo.FileHeader(zip64))
		self.fp.seek(position, 0)
		self.filelist.append(zinfo)
		self.NameToInfo[zinfo.filename] = zinfo

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
		zf = ZipFile(f, 'w', zipfile.ZIP_DEFLATED)
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
			template_stream = self._worksheet_template.stream({'worksheet': sheet})
			template_stream.enable_buffering(1024 * 8)
			zf.writeiter("xl/worksheets/sheet%s.xml" % (index), template_stream)
		zf.close()
