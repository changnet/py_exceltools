#! python
# -*- coding:utf-8 -*-

import os
import re
import sys
import json
import openpyxl

from slpp.slpp import slpp as lua

# array数组模式下，各个栏的分布
# 1、2行是文档注释，不导出
ACMT_ROW = 3  # comment row 注释
ATPE_ROW = 4  # type    row 类型
ASRV_ROW = 5  # server  row 服务器
ACLT_ROW = 6  # client  row 客户端
AKEY_COL = 1  # key  column key所在列

# object(kv)模式下，各个栏的分布
OCMT_COL = 1  # comment column 注释
OTPE_COL = 2  # type    column 类型
OSRV_COL = 3  # server  column 服务器
OCLT_COL = 4  # client  column 客户端
OCTX_COL = 5  # content column 内容所在列

# 1、2行是文档注释，不导出
OFLG_ROW = 3  # flag    row    server client所在行

SRV_FLAG = "server"
CLT_FLAG = "client"

SHEET_FLAG_ROW = 3
SHEET_FLAG_COL = 1

ARRAY_FLAG = "array"
OBJECT_FLAG = "object"

# 索引的位置
INDEX_ROW = 1
INDEX_COL = 1

TYPES = {"num": True, "str": True, "json": True, "lua": True}

# 用于判断是否整数的正则
INT_RE = re.compile(r"^[-]?\d+$")


def is_int_str(s):
    return INT_RE.match(s) is not None

# 类型转换器


class ValueConverter(object):
    def __init__(self):
        pass

    # 在python中，字符串和unicode是不一样的。默认从excel读取的数据都是unicode。
    # str可以通过decode转换为unicode
    # ascii' codec can't encode characters
    # 这个函数在python3中没用
    def to_unicode_str(self, val):
        if isinstance(val, str):
            return val
        else:
            return str(val).decode("utf8")

    def to_value(self, val_type, val):
        if "num" == val_type:
            if is_int_str(val):
                return int(val)

            return float(val)
        elif "str" == val_type:
            return self.to_unicode_str(val)
        elif "json" == val_type:
            return json.loads(val)
        elif "lua" == val_type:
            return lua.decode(val)
        else:
            self.raise_error("invalid type", val_type)

# 继承object类，以解决在python2中的错误：TypeError: must be type, not classobj


class Sheet(object):

    def __init__(self, base_name, wb_sheet, index, srv_writer, clt_writer):
        self.types = []  # 记录各列字段的类型

        self.srv_writer = srv_writer
        self.clt_writer = clt_writer

        self.fields = []  # 各列字段名
        self.index = index  # 索引数量，0表示kv模式

        self.converter = ValueConverter()

        self.wb_sheet = wb_sheet
        self.base_name = base_name

        self.dir_path = None  # 导出目录
        self.file_path = None  # 卖出文件名

        # 记录出错时的行列，方便定位问题
        self.error_row = 0
        self.error_col = 0

    # 记录出错位置
    def mark_error_pos(self, row, col):
        if row > 0:
            self.error_row = row
        if col > 0:
            self.error_col = col

    # 发起一个解析错误
    def raise_error(self, what, val):
        excel_info = format("DOC:%s,SHEET:%s,ROW:%d,COLUMN:%d" %
                            (self.base_name, self.wb_sheet.title, self.error_row, self.error_col))
        raise Exception(what, val, excel_info)

    def to_value(self, val_type, val):
        try:
            return self.converter.to_value(val_type, val)
        except Exception:
            t, e = sys.exc_info()[:2]
            self.raise_error("ConverError", e)

    # 解析基础信息，如目录、导出文件名等
    def decode_info(self):
        self.dir_path = self.wb_sheet.cell(
            row=INDEX_ROW, column=INDEX_COL + 1).value
        self.file_path = self.wb_sheet.cell(
            row=INDEX_ROW, column=INDEX_COL + 2).value

    # 解析一个表格
    def decode_sheet(self):
        wb_sheet = self.wb_sheet

        self.decode_info()  # 解析基础信息，如目录、导出文件名等
        self.decode_type()  # 解析字段类型，num、str等
        self.decode_opt()  # 解析导出参数，是否导出服务端、客户端
        self.decode_ctx()  # 解析内容

        print("    decode sheet %s done" % wb_sheet.title.ljust(24, "."))
        return True

    # 写入配置到文件
    def write_one_file(self, ctx, base_path, writer):
        # 有些配置可能只导出客户端或只导出服务器
        if not any(ctx):
            return

        wt = writer(self.base_name, self.wb_sheet.title)

        ctx = wt.context(ctx)
        suffix = wt.suffix()

        # 必须为wb，不然无法写入utf-8
        path = base_path + self.base_file_name\
            + "_" + self.wb_sheet.title + suffix
        file = open(path, 'wb')
        file.write(ctx.encode("utf-8"))
        file.close()

    # 分别写入到服务端、客户端的配置文件
    def write_files(self, srv_path, clt_path):
        if None != srv_path and None != self.srv_writer:
            self.write_one_file(self.srv_ctx, srv_path, self.srv_writer)
        if None != clt_path and None != self.clt_writer:
            self.write_one_file(self.clt_ctx, clt_path, self.clt_writer)

# 导出数组类型配置


class ArraySheet(Sheet):

    def __init__(self, base_name, wb_sheet, srv_writer, clt_writer):
        # 记录导出各行的内容
        self.srv_ctx = []
        self.clt_ctx = []

        super(ArraySheet, self).__init__(
            base_name, wb_sheet, srv_writer, clt_writer)

    # 解析各列的类型(string、number...)
    def decode_type(self):
        # 第一列没数据，类型可以不填，默认为None，但是这里要占个位
        self.types.append(None)

        for col_idx in range(AKEY_COL + 1, self.wb_sheet.max_column + 1):
            self.mark_error_pos(ATPE_ROW, col_idx)
            value = self.wb_sheet.cell(row=ATPE_ROW, column=col_idx).value

            # 单元格为空的时候，wb_sheet.cell(row=1, column=2).value == None
            # 类型那一行必须连续，空白表示后面的数据都不导出了
            if value == None:
                break
            if value not in TYPES:
                self.raise_error("invalid type", value)

            self.types.append(value)

    # 解析客户端、服务器的字段名(server、client)那两行
    def decode_one_field(self, fields, row_index):
        for col_index in range(AKEY_COL, len(self.types) + 1):
            value = self.wb_sheet.cell(
                row=row_index, column=col_index).value

            # 对于不需要导出的field，可以为空。即value为None
            fields.append(value)

    # 导出客户端、服务端字段名(server、client)那一列
    def decode_opt(self):
        self.decode_one_field(self.srv_fields, ASRV_ROW)
        self.decode_one_field(self.clt_fields, ACLT_ROW)

    # 解析出一个格子的内容
    def decode_cell(self, row_idx, col_idx):
        value = self.wb_sheet.cell(row=row_idx, column=col_idx).value
        if None == value:
            return None

        # 类型是从0下标开始，但是excel的第一列从1开始
        self.mark_error_pos(row_idx, col_idx)
        return self.to_value(self.types[col_idx - 1], value)

    # 解析出一行的内容
    def decode_row(self, row_idx):
        srv_row = {}
        clt_row = {}

        # 第一列没数据，从第二列开始解析
        for col_idx in range(AKEY_COL + 1, len(self.types) + 1):
            value = self.decode_cell(row_idx, col_idx)
            if None == value:
                continue

            srv_key = self.srv_fields[col_idx - 1]
            clt_key = self.clt_fields[col_idx - 1]

            if srv_key:
                srv_row[srv_key] = value
            if clt_key:
                clt_row[clt_key] = value

        return srv_row, clt_row  # 返回一个tuple

    # 解析导出的内容
    def decode_ctx(self):
        for row_idx in range(ACLT_ROW + 1, self.wb_sheet.max_row + 1):
            srv_row, clt_row = self.decode_row(row_idx)

            # 不为空才追加
            if any(srv_row):
                self.srv_ctx.append(srv_row)
            if any(clt_row):
                self.clt_ctx.append(clt_row)

# 导出object类型的结构


class ObjectSheet(Sheet):

    def __init__(self, base_name, wb_sheet, index, srv_writer, clt_writer):
        # 记录导出各行的内容
        self.srv_ctx = {}
        self.clt_ctx = {}

        super(ObjectSheet, self).__init__(
            base_name, wb_sheet, index, srv_writer, clt_writer)

    # 解析各字段的类型
    def decode_type(self):
        col = 2  # 类型所在列
        empty_cnt = 0
        for row_idx in range(2, self.wb_sheet.max_row + 1):
            self.mark_error_pos(row_idx, col)
            value = self.wb_sheet.cell(row=row_idx, column=col).value

            # 中间可能有些注释字段，但如果太多的话，可能是出错
            if value == None:
                empty_cnt = empty_cnt + 1
                if empty_cnt > 32:
                    # print("warning sheet %s has too many empty type" %
                    #       self.wb_sheet.title)
                    break
            elif value not in TYPES:
                self.raise_error("invalid type", value)

            self.types.append(value)

    # 导出客户端、服务端字段名(server、client)那一列
    def decode_one_field(self, fields, col_idx):
        for row_idx in range(OFLG_ROW + 1, len(self.types) + OFLG_ROW + 1):
            value = self.wb_sheet.cell(row=row_idx, column=col_idx).value

            # 对于不需要导出的field，可以为空。即value为None
            fields.append(value)

    # 导出客户端、服务端字段名(server、client)那一列
    def decode_opt(self):
        self.decode_one_field(self.srv_fields, OSRV_COL)
        self.decode_one_field(self.clt_fields, OCLT_COL)

    # 解析一个单元格内容
    def decode_cell(self, row_idx):
        value = self.wb_sheet.cell(row=row_idx, column=OCTX_COL).value
        if None == value:
            return None

        # 在object的结构中，数据是从第二行开始的，所以types的下标偏移2
        self.mark_error_pos(row_idx, OCTX_COL)
        return self.to_value(self.types[row_idx - OFLG_ROW - 1], value)

    # 解析表格的所有内容
    def decode_ctx(self):
        for row_idx in range(OFLG_ROW + 1, len(self.types) + OFLG_ROW + 1):
            value = self.decode_cell(row_idx)
            if None == value:
                continue

            srv_key = self.srv_fields[row_idx - OFLG_ROW - 1]
            clt_key = self.clt_fields[row_idx - OFLG_ROW - 1]

            if srv_key:
                self.srv_ctx[srv_key] = value
            if clt_key:
                self.clt_ctx[clt_key] = value


class ExcelDoc:

    def __init__(self, file, abspath):
        self.file = file
        self.abspath = abspath

    # 是否需要解析
    def need_decode(self, wb_sheet):
        sheet_val = wb_sheet.cell(
            row=INDEX_ROW, column=INDEX_COL).value

        # 判断索引是否正确，-1表示kv模式，其他表示索引数量
        if "-1" == sheet_val:
            return -1
        elif sheet_val.isdigit():
            index = int(sheet_val)
            if index > 8:
                print("sheet %s too many index, abort" %
                      wb_sheet.title.ljust(24, "."))
                return -2
            return index
        else:
            return -2

    def decode(self, srv_path, clt_path, srv_writer, clt_writer):
        print("start decode %s ..." % self.file)

        base_name = os.path.splitext(self.file)[0]  # 去除后缀
        wb = openpyxl.load_workbook(self.abspath)

        for wb_sheet in wb.worksheets:
            index = self.need_decode(wb_sheet)
            if index < -1:
                print("    sheet %s no need to decode,abort" %
                      wb_sheet.title.ljust(24, "."))
                continue
            elif -1 == index:
                sheet = ObjectSheet(base_name, wb_sheet,
                                    index, srv_writer, clt_writer)
            else:
                sheet = ArraySheet(base_name, wb_sheet, index,
                                   srv_writer, clt_writer)

            if sheet.decode_sheet():
                sheet.write_files(srv_path, clt_path)
