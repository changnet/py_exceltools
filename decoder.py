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

    def __init__(self,base_name,wb_sheet):
        self.rows   = []
        self.types  = []

        self.srv_fields = []
        self.clt_fields = []

        self.wb_sheet  = wb_sheet
        self.base_name = base_name

    def decode_type(self):
        # key的类型可以不填写，默认为None
        self.types.append( self.wb_sheet.cell(row=TYPE_ROW, column=KEY_COL).value )

        for col_index in range( KEY_COL + 1,self.wb_sheet.max_column + 1 ):
            value = self.wb_sheet.cell( row = TYPE_ROW, column = col_index ).value

            # 单元格为空的时候，wb_sheet.cell(row=1, column=2).value == None
            if value == None: break
            if not TYPES[value]:
                raise Exception( "invalid type",value )
            
            self.types.append( value )

    def decode_field(self,fields,row_index):
        for col_index in range( KEY_COL,len( self.types ) + 1 ):
            value = self.wb_sheet.cell( 
                row = row_index, column = col_index ).value

            # 对于不需要导出的field，可以为空。即value为None
            fields.append( value )

    def decode_cell(self):
        for row_index in range( CLT_ROW + 1,self.wb_sheet.max_row + 1 ):
            column_values = []
            # 从第一列开始解析，即包括key
            for col_index in range( KEY_COL,len( self.types ) + 1 ):
                value = self.wb_sheet.cell( 
                    row = row_index, column = col_index ).value
                column_values.append( value )
            
            self.rows.append( column_values )

    def write_one_file(self,fields,base_path,writer):
        wt = writer( self.types,fields,self.rows )
        ctx = wt.content()
        suffix = wt.suffix()

        #必须为wb，不然无法写入utf-8
        if not os.path.exists( base_path ) : os.makedirs( base_path )
        path = base_path + self.base_name + "_" + self.wb_sheet.title + suffix
        file = open( path, 'wb' )
        file.write( ctx.encode( "utf-8" ) )
        file.close()

    def write_files(self,srv_path,clt_path,writer):
        self.write_one_file( self.srv_fields,srv_path,writer )
        self.write_one_file( self.clt_fields,clt_path,writer )

    def decode_sheet(self):
        wb_sheet = self.wb_sheet
        if wb_sheet.max_row <= CLT_ROW or wb_sheet.max_column <= KEY_COL:
            print( "    decode sheet %s nothing to decode,abort" \
            % wb_sheet.title.ljust(24,".") )
            return False
        
        self.decode_type()
        if len( self.types ) <= TYPE_ROW:
            print( "    decode sheet %s nothing to decode,abort" \
            % wb_sheet.title.ljust(24,".") )
            return False

        self.decode_field( self.srv_fields,SRV_ROW )
        self.decode_field( self.clt_fields,CLT_ROW )

        self.decode_cell()

        print( "    decode sheet %s done" % wb_sheet.title.ljust(24,".") )
        return True

class ExcelDoc:

    def __init__(self, file):
        self.file = file

    def decode(self,srv_path,clt_path,writer):
        print( "start decode %s ..." % self.file )

        base_name = os.path.splitext( self.file )[0]  #去除后缀
        wb = openpyxl.load_workbook( self.file )

        for wb_sheet in wb.worksheets:
            sheet = Sheet( base_name,wb_sheet )
            if sheet.decode_sheet():
                sheet.write_files( srv_path,clt_path,writer )