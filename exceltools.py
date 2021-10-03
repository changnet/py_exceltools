#! python
# -*- coding:utf-8 -*-

import os
import time
from writer import *
from optparse import OptionParser

from decoder import ExcelDoc


class Reader:
    # @input_path:excel文件所在目录
    # @srv_path  :server输出目录
    # @clt_path  :客户端输出目录
    # @timeout   :只处理文档最后更改时间在N秒内的文档
    # @suffix    :excel文件后缀
    def __init__(self, input_path,
                 srv_path, clt_path, timeout, suffix, srv_writer, clt_writer):
        self.input_path = input_path
        self.srv_path = srv_path
        self.clt_path = clt_path
        self.timeout = timeout
        self.suffix = suffix

        self.srv_writer = None
        self.clt_writer = None

        # json则对应的类为JsonWriter
        if None != srv_writer:
            self.srv_writer = eval(srv_writer.capitalize() + "Writer")
        if None != clt_writer:
            self.clt_writer = eval(clt_writer.capitalize() + "Writer")

    def attention(self):
        # print("********excel转换********")
        pass

    def can_read(self, file, abspath):
        if not os.path.isfile(abspath):
            return False
        # ~开头的excel文件是临时文件
        # linux下wps临时文件以.~开头
        # 永中office是$开头
        # 这些文件都忽略掉
        if file.startswith("~") \
                or file.startswith(".") or file.startswith("$"):
            return False
        if "" != self.suffix and not file.endswith(self.suffix):
            return False

        # 按修改时间导出
        if self.timeout > 0:
            now = time.time()
            mtime = os.path.getmtime(abspath)

            if now - mtime > self.timeout:
                return False

        return True

    def read(self):
        if self.timeout > 0:
            print("read %s files from %s modified \
                within %d seconds" % (self.suffix, self.input_path, self.timeout))
        else:
            print("read %s files from %s" % (self.suffix, self.input_path))

        if None != self.srv_path and not os.path.exists(self.srv_path):
            os.makedirs(self.srv_path)
        if None != self.clt_path and not os.path.exists(self.clt_path):
            os.makedirs(self.clt_path)

        now = time.time()
        file_list = os.listdir(options.input_path)
        for file in file_list:
            abspath = os.path.join(self.input_path, file)
            if self.can_read(file, abspath):
                self.read_one(file, abspath)

        print("done,%d seconds elapsed" % (time.time() - now))

    def read_one(self, file, abspath):
        doc = ExcelDoc(file, abspath)
        doc.decode(self.srv_path,
                   self.clt_path, self.srv_writer, self.clt_writer)


# 该文件是否需要处理
# @file 文件名
# @path 完整的文件路径
# @timeout 仅导出修改时间在N秒内的文件
def need(file, path, timeout):
    if not os.path.isfile(path):
        return False
    # ~开头的excel文件是临时文件
    # linux下wps临时文件以.~开头
    # 永中office是$开头
    # 这些文件都忽略掉
    if file.startswith("~") \
            or file.startswith(".") or file.startswith("$"):
        return False

    # 目前仅检测两种常用的excel文件，以后有需要再加
    if not file.endswith(".xlsx") and not file.endswith(".xlsm"):
        return False

    # 按修改时间导出
    if timeout > 0:
        mtime = os.path.getmtime(path)

        if time.time() - mtime > timeout:
            return False

    return True

# 开始处理单个excel文件
def do_one_excel(file, path, srv_path, clt_path, timeout, SrvWriter, CltWriter):
    if not need(file, path, timeout): return False

    doc = ExcelDoc(file, path)
    doc.decode(srv_path, clt_path, SrvWriter, CltWriter)
    return True

# 开始执行excel转换
# @input_path:excel文件所在目录
# @srv_path  :server输出目录
# @clt_path  :客户端输出目录
# @timeout   :只处理文档最后更改时间在N秒内的文档
def do_excel_tools(input_path,
                 srv_path, clt_path, timeout, srv_writer, clt_writer):

    if timeout > 0:
        print("read excel files from %s modified \
            within %d seconds" % (input_path, timeout))
    else:
        print("read excel files from %s" % (input_path))

    # 如果导出的前后端目录不存在，则创建
    if None != srv_path and not os.path.exists(srv_path):
        os.makedirs(srv_path)
    if None != clt_path and not os.path.exists(clt_path):
        os.makedirs(clt_path)

    SrvWriter = None
    CltWriter = None
    # json则对应的类为JsonWriter
    if None != srv_writer:
        SrvWriter = eval(srv_writer.capitalize() + "Writer")
    if None != clt_writer:
        CltWriter = eval(clt_writer.capitalize() + "Writer")

    count = 0
    now = time.time()
    file_list = os.listdir(input_path)
    for file in file_list:
        path = os.path.join(input_path, file)
        if do_one_excel(file, path, srv_path, clt_path, timeout, SrvWriter, CltWriter):
            count = count + 1

    print("done,%d files, %d seconds elapsed" % (count, time.time() - now))

if __name__ == '__main__':

    parser = OptionParser()

    parser.add_option("-i", "--input", dest="input_path",
                      default="xls/",
                      help="read all files from this path")
    parser.add_option("-s", "--srv", dest="srv_path",
                      help="write all server file to this path")
    parser.add_option("-c", "--clt", dest="clt_path",
                      help="write all client file to this path")
    parser.add_option("-t", "--timeout", dest="timeout", type="int",
                      default="-1",
                      help="only converte files modified within seconds")
    parser.add_option("-w", "--swriter", dest="srv_writer",
                      help="which server writer you wish to use:lua xml json")
    parser.add_option("-l", "--cwriter", dest="clt_writer",
                      help="which client writer you wish to use:lua xml json")

    options, args = parser.parse_args()

    do_excel_tools(options.input_path, options.srv_path, options.clt_path,
                    options.timeout, options.srv_writer, options.clt_writer)

