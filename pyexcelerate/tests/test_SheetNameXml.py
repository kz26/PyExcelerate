# -*- coding: utf-8 -*-

from __future__ import unicode_literals

from io import BytesIO

import openpyxl
from nose.tools import eq_

from ..Workbook import Workbook


def test_sheet_name_with_xml_characters():
    workbook = Workbook()
    workbook.new_sheet("R&D", data=[[1]])
    workbook.new_sheet('<Q&A "draft">', data=[[2]])
    output = BytesIO()
    workbook.save(output)
    output.seek(0)

    eq_(openpyxl.load_workbook(output).sheetnames, ["R&D", '<Q&A "draft">'])
