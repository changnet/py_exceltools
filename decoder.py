#! python
# -*- coding:utf-8 -*-

import os
import sys
import openpyxl

from optparse import OptionParser

try:
    basestring
except NameError:
    basestring = str

TYPE_ROW = 1
SRV_ROW  = 2
CLT_ROW  = 3
KEY_COL  = 1

TYPES = { "int":1,"number":2,"int64":3,"string":4,"json":5 }

class Sheet:

    def __init__(self):
        self.rows   = []
        self.types  = []
        self.srv_fields = []
        self.clt_fields = []
    
    def decode_type(self,wb_sheet):
        # key的类型可以不填写，默认为None
        self.types.append( wb_sheet.cell(row=TYPE_ROW, column=KEY_COL).value )

        for col_index in range( KEY_COL + 1,wb_sheet.max_column + 1 ):
            value = wb_sheet.cell( row = TYPE_ROW, column = col_index ).value

            # 单元格为空的时候，wb_sheet.cell(row=1, column=2).value == None
            if value == None: break
            if not TYPES[value]:
                raise Exception( "invalid type",value )
            
            self.types.append( value )

    def decode_field(self,wb_sheet,fields,row_index):
        for col_index in range( KEY_COL,len( self.types ) + 1 ):
            value = wb_sheet.cell( row = row_index, column = col_index ).value

            # 对于不需要导出的field，可以为空。即value为None
            fields.append( value )
    
    def decode_cell(self,wb_sheet):
        for row_index in range( CLT_ROW + 1,wb_sheet.max_row + 1 ):
            column_values = []
            # 从第一列开始解析，即包括key
            for col_index in range( KEY_COL,len( self.types ) + 1 ):
                value = wb_sheet.cell( 
                    row = row_index, column = col_index ).value
                column_values.append( value )
            
            self.rows.append( column_values )

    def write_files(self,srv_path,clt_path,writer):
        wt = writer( self.types,self.srv_fields,self.rows )
        print( wt.comment() )

    def decode_sheet(self,wb_sheet):

        if wb_sheet.max_row <= CLT_ROW or wb_sheet.max_column <= KEY_COL:
            print( "    decode sheet %s nothing to decode,abort" \
            % wb_sheet.title.ljust(24,".") )
            return False
        
        self.decode_type( wb_sheet )
        if len( self.types ) <= TYPE_ROW:
            print( "    decode sheet %s nothing to decode,abort" \
            % wb_sheet.title.ljust(24,".") )
            return False

        self.decode_field( wb_sheet,self.srv_fields,SRV_ROW )
        self.decode_field( wb_sheet,self.clt_fields,CLT_ROW )

        self.decode_cell( wb_sheet )

        print( "    decode sheet %s done" % wb_sheet.title.ljust(24,".") )
        return True

class ExcelDoc:

    def __init__(self, file):
        self.file = file

    def decode(self,srv_path,clt_path,writer):
        print( "start decode %s ..." % self.file )

        wb = openpyxl.load_workbook( self.file )

        for wb_sheet in wb.worksheets:
            sheet = Sheet()
            if sheet.decode_sheet( wb_sheet ):
                sheet.write_files( srv_path,clt_path,writer )