#! python
# -*- coding:utf-8 -*-

import os
import openpyxl

from optparse import OptionParser

try:
    basestring
except NameError:
    basestring = str

TYPE_ROW = 1
SRV_ROW  = 2
CLT_ROW  = 3

class Sheet:

    def __init__(self,base_name):
        self.base_name = base_name
        self.errors = 0
        self.warns  = 0

    def decode_sheet(self,wb_sheet):
        print( "    decoding %s" % wb_sheet.title )
        print( wb_sheet.cell(row=1, column=2).value )
        # 空的时候，wb_sheet.cell(row=1, column=2).value == None

class ExcelDoc:

    def __init__(self, file):
        self.file = file
        self.errors = 0
        self.warns  = 0

    def decode(self):
        print( "start decode %s ..." % self.file )
        base_name = os.path.splitext( self.file )[0]  #去除后缀

        wb = openpyxl.load_workbook( self.file )

        for wb_sheet in wb.worksheets:
            sheet = Sheet( base_name )
            sheet.decode_sheet( wb_sheet )