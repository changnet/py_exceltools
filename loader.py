#! python
# -*- coding:utf-8 -*-

import os
import time
import importlib
from optparse import OptionParser

from decoder import ExcelDoc

class Loader:
    # @input_path:excel文件所在目录
    # @srv_path  :server输出目录
    # @clt_path  :客户端输出目录
    # @timeout   :只处理文档最后更改时间在N秒内的文档
    # @suffix    :excel文件后缀
    def __init__(self,input_path,
        srv_path,clt_path,timeout,suffix,srv_writer,clt_writer):
        self.input_path = input_path
        self.srv_path   = srv_path
        self.clt_path   = clt_path
        self.timeout    = timeout
        self.suffix     = suffix

        self.srv_writer = importlib.import_module( "writer_" + srv_writer )
        self.clt_writer = importlib.import_module( "writer_" + clt_writer )

    def attention(self):
        print("********excel转换********")
        print("**第一行为数据类型，支持int、number、int64、string、json")
        print("**第二行为server字段名，为空则不导出该字段")
        print("**第三行为client字段名，为空则不导出该字段")
        print("**第一列为主键，不可重复。可以为字符串，为空则转换为数组")
        print("***************************************************\n")

    def can_load(self,file):
        if not os.path.isfile( file ): return False
        # ~开头的excel文件是临时文件
        if file.startswith( "~" ): return False
        if "" != self.suffix and not file.endswith( self.suffix ): return False

        if self.timeout > 0:
            now = time.time()
            mtime = os.path.getmtime( file )

            if now - mtime > self.timeout: return False

        return True

    def load(self):
        print("load %s files from %s modified in the last %d seconds" 
            % (self.suffix,self.input_path,self.timeout))

        if not os.path.exists( self.srv_path ) : os.makedirs( self.srv_path )
        if not os.path.exists( self.clt_path ) : os.makedirs( self.clt_path )
        now = time.time()
        file_list = os.listdir( options.input_path )
        for file in file_list:
            if self.can_load( file ):self.load_one( file )
        print( "load done,%d second elapsed" % ( time.time() - now ) )
    
    def load_one(self,file):
        doc = ExcelDoc( file )
        doc.decode( self.srv_path,
            self.clt_path,self.srv_writer,self.clt_writer )

if __name__ == '__main__':

    parser = OptionParser()

    parser.add_option( "-i", "--input", dest="input_path",
                     default="xls/",
                     help="read all files from this path" )
    parser.add_option( "-s", "--srv", dest="srv_path",
                     default="server/",
                     help="write all server file to this path" )
    parser.add_option( "-c", "--clt", dest="clt_path",
                     default="client/",
                     help="write all client file to this path" )
    parser.add_option( "-t", "--timeout", dest="timeout",type="int",
                     default="-1",
                     help="only converte files modified within seconds" )
    parser.add_option( "-f", "--suffix", dest="suffix",
                     default="",
                     help="what type of file will be loaded.empty mean all files" )
    parser.add_option( "-m","--swriter", dest="srv_writer",
                     default="lua",
                     help="which server writer you wish to use:lua xml json" )
    parser.add_option( "-n","--cwriter", dest="clt_writer",
                     default="lua",
                     help="which client writer you wish to use:lua xml json" )

    options, args = parser.parse_args()

    loader = Loader( options.input_path,options.srv_path,options.clt_path,
        options.timeout,options.suffix,options.srv_writer,options.clt_writer )
    loader.attention()
    loader.load()