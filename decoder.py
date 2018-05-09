#! python
# -*- coding:utf-8 -*-

import os
import error
import openpyxl
from slpp.slpp import slpp as lua

# 数组模式下，各个栏的分布
ACMT_ROW = 1  # comment row 注释
ATPE_ROW = 2  # type    row 类型
ASRV_ROW = 3  # server  row 服务器
ACLT_ROW = 4  # client  row 客户端
AKEY_COL = 1  # key  column key所在列

# kv模式下，各个栏的分布
OCMT_COL = 1  # comment column 注释
OTPE_COL = 2  # type    column 类型
OSRV_COL = 3  # server  column 服务器
OCLT_COL = 4  # client  column 客户端
OCTX_COL = 5  # content column 内容所在列
OFLG_ROW = 1  # flag    row    server client所在行

SRV_FLAG = "server"
CLT_FLAG = "client"

SHEET_FLAG_ROW = 1
SHEET_FLAG_COL = 1

ARRAY_FLAG  = "array"
OBJECT_FLAG = "object"

TYPES = { "int":1,"number":2,"int64":3,"string":4,"json":5 }

# 继承object类，以解决在python2中的错误：TypeError: must be type, not classobj
class Sheet(object):

    def __init__(self,base_name,wb_sheet,srv_writer,clt_writer):
        self.rows   = []
        self.types  = []

        self.srv_writer = srv_writer
        self.clt_writer = clt_writer

        self.srv_fields = []
        self.clt_fields = []

        self.wb_sheet  = wb_sheet
        self.base_name = base_name

    def write_one_file(self,fields,base_path,writer):
        if len( fields ) <= 0 : return

        wt = writer.Writer( self.base_name,
            self.wb_sheet.title,self.row_offset,self.col_offset )
        ctx = self.writer_content( wt,fields )
        suffix = wt.suffix()

        #必须为wb，不然无法写入utf-8
        path = base_path + self.base_name + "_" + self.wb_sheet.title + suffix
        file = open( path, 'wb' )
        file.write( ctx.encode( "utf-8" ) )
        file.close()

    def write_files(self,srv_path,clt_path):
        if None != srv_path and None != self.srv_writer :
            self.write_one_file( self.srv_fields,srv_path,self.srv_writer )
        if None != clt_path and None != self.clt_writer :
            self.write_one_file( self.clt_fields,clt_path,self.clt_writer )

class ArraySheet(Sheet):

    def __init__(self,base_name,wb_sheet,srv_writer,clt_writer):
        self.row_offset = ACLT_ROW
        self.col_offset = AKEY_COL
        super( ArraySheet, self ).__init__(
            base_name,wb_sheet,srv_writer,clt_writer )

    def decode_type(self):
        # key的类型可以不填写，默认为None
        self.types.append( self.wb_sheet.cell(row=ATPE_ROW, column=AKEY_COL).value )

        for col_index in range( AKEY_COL + 1,self.wb_sheet.max_column + 1 ):
            value = self.wb_sheet.cell( row = ATPE_ROW, column = col_index ).value

            # 单元格为空的时候，wb_sheet.cell(row=1, column=2).value == None
            if value == None: break
            if not TYPES[value]:
                raise Exception( "invalid type",value )

            self.types.append( value )

    def decode_field(self,fields,row_index):
        for col_index in range( AKEY_COL,len( self.types ) + 1 ):
            value = self.wb_sheet.cell(
                row = row_index, column = col_index ).value

            # 对于不需要导出的field，可以为空。即value为None
            fields.append( value )

    def decode_cell(self):
        for row_index in range( ACLT_ROW + 1,self.wb_sheet.max_row + 1 ):
            column_values = []
            # 从第一列开始解析，即包括key
            for col_index in range( AKEY_COL,len( self.types ) + 1 ):
                value = self.wb_sheet.cell(
                    row = row_index, column = col_index ).value
                column_values.append( value )

            self.rows.append( column_values )

    def writer_content(self,writer,fields):
        return writer.array_content( self.types,fields,self.rows )

    def decode_sheet(self):
        wb_sheet = self.wb_sheet

        self.decode_type()
        if len( self.types ) <= ATPE_ROW:
            print( "    decode sheet %s nothing to decode,abort" \
            % wb_sheet.title.ljust(24,".") )
            return False

        self.decode_field( self.srv_fields,ASRV_ROW )
        self.decode_field( self.clt_fields,ACLT_ROW )

        self.decode_cell()

        print( "    decode sheet %s done" % wb_sheet.title.ljust(24,".") )
        return True

class ObjectSheet(Sheet):

    def __init__(self,base_name,wb_sheet,srv_writer,clt_writer):
        self.row_offset = OCLT_COL
        self.col_offset = OFLG_ROW
        super( ObjectSheet, self ).__init__(
            base_name,wb_sheet,srv_writer,clt_writer )

    def decode_type(self):
        for row_index in range( OFLG_ROW + 1,self.wb_sheet.max_row + 1 ):
            value = self.wb_sheet.cell(
                row = row_index, column = OTPE_COL ).value

            # 单元格为空的时候，wb_sheet.cell(row=1, column=2).value == None
            if value == None: break
            if not TYPES[value]:
                raise Exception( "invalid type",value )

            self.types.append( value )

    def decode_field(self,fields,col_index):
        for row_index in range( OFLG_ROW + 1,len( self.types ) + 2 ):
            value = self.wb_sheet.cell(
                row = row_index, column = col_index ).value

            # 对于不需要导出的field，可以为空。即value为None
            fields.append( value )

    def decode_cell(self):
        # 第一行为flag行，包括最后一行，所以要types + 2
        for row_index in range( OFLG_ROW + 1,len( self.types ) + 2 ):
            value = self.wb_sheet.cell(
                row = row_index, column = OCTX_COL ).value

            self.rows.append( value )

    def writer_content(self,writer,fields):
        return writer.object_content( self.types,fields,self.rows )

    def decode_sheet(self):
        wb_sheet = self.wb_sheet

        self.decode_type()
        if len( self.types ) <= 0:
            print( "    decode sheet %s nothing to decode,abort" \
            % wb_sheet.title.ljust(24,".") )
            return False

        self.decode_field( self.srv_fields,OSRV_COL )
        self.decode_field( self.clt_fields,OCLT_COL )

        self.decode_cell()

        print( "    decode sheet %s done" % wb_sheet.title.ljust(24,".") )
        return True

class ExcelDoc:

    def __init__(self, file,abspath):
        self.file = file
        self.abspath = abspath

    # 是否需要解析
    # 返回解析的对象类型
    def need_decode(self,wb_sheet):
        sheet_val = wb_sheet.cell(
            row = SHEET_FLAG_ROW, column = SHEET_FLAG_ROW ).value

        sheeter = None
        srv_value = None
        clt_value = None
        if ARRAY_FLAG == sheet_val :
            if wb_sheet.max_row <= ACLT_ROW or wb_sheet.max_column <= AKEY_COL:
                return None

            sheeter = ArraySheet
            srv_value = wb_sheet.cell( row = ASRV_ROW, column = AKEY_COL ).value
            clt_value = wb_sheet.cell( row = ACLT_ROW, column = AKEY_COL ).value
        elif OBJECT_FLAG == sheet_val :
            sheeter = ObjectSheet
            srv_value = wb_sheet.cell( row = OFLG_ROW, column = OSRV_COL ).value
            clt_value = wb_sheet.cell( row = OFLG_ROW, column = OCLT_COL ).value
        else :
            return None

        # 没有这两个标识，说明不是配置表。可能是策划的一些备注说明
        if SRV_FLAG != srv_value or CLT_FLAG != clt_value: return None
        return sheeter

    def decode(self,srv_path,clt_path,srv_writer,clt_writer):
        print( "start decode %s ..." % self.file )

        base_name = os.path.splitext( self.file )[0]  #去除后缀
        wb = openpyxl.load_workbook( self.abspath )

        for wb_sheet in wb.worksheets:
            Sheeter = self.need_decode( wb_sheet )

            if None == Sheeter :
                print( "    decode sheet %s no need to decode,abort" % wb_sheet.title.ljust(24,".") )
                continue

            sheet = Sheeter( base_name,wb_sheet,srv_writer,clt_writer )
            if sheet.decode_sheet(): sheet.write_files( srv_path,clt_path )