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

        self.srv_writer = None
        self.clt_writer = None
        if None != srv_writer :
            self.srv_writer = importlib.import_module( "writer_" + srv_writer )
        if None != clt_writer :
            self.clt_writer = importlib.import_module( "writer_" + clt_writer )

    def attention(self):
        print("********excel转换********")
        print("**第一行为数据类型，支持int、number、int64、string、json")
        print("**第二行为server字段名，为空则不导出该字段")
        print("**第三行为client字段名，为空则不导出该字段")
        print("**第一列为主键，不可重复。可以为字符串，为空则转换为数组")
        print("***************************************************\n")

    def can_load(self,file,abspath):
        if not os.path.isfile( abspath ): return False
        # ~开头的excel文件是临时文件，linux下wps临时文件以.~开头
        if file.startswith( "~" ) or file.startswith( "." ): return False
        if "" != self.suffix and not file.endswith( self.suffix ): return False

        if self.timeout > 0:
            now = time.time()
            mtime = os.path.getmtime( abspath )

            if now - mtime > self.timeout: return False

        return True

    def load(self):
        print("load %s files from %s modified in the last %d seconds" 
            % (self.suffix,self.input_path,self.timeout))

        if None != self.srv_path and not os.path.exists( self.srv_path ) :
            os.makedirs( self.srv_path )
        if None != self.clt_path and not os.path.exists( self.clt_path ) :
            os.makedirs( self.clt_path )

        now = time.time()
        file_list = os.listdir( options.input_path )
        for file in file_list:
            abspath = os.path.join( self.input_path,file )
            if self.can_load( file,abspath ):self.load_one( file,abspath )

        print( "done,%d second elapsed" % ( time.time() - now ) )
    
    def load_one(self,file,abspath):
        doc = ExcelDoc( file,abspath )
        doc.decode( self.srv_path,
            self.clt_path,self.srv_writer,self.clt_writer )

if __name__ == '__main__':

    parser = OptionParser()

    parser.add_option( "-i", "--input", dest="input_path",
                     default="xls/",
                     help="read all files from this path" )
    parser.add_option( "-s", "--srv", dest="srv_path",
                     help="write all server file to this path" )
    parser.add_option( "-c", "--clt", dest="clt_path",
                     help="write all client file to this path" )
    parser.add_option( "-t", "--timeout", dest="timeout",type="int",
                     default="-1",
                     help="only converte files modified within seconds" )
    parser.add_option( "-f", "--suffix", dest="suffix",
                     default="",
                     help="what type of file will be loaded.empty mean all files" )
    parser.add_option( "-w","--swriter", dest="srv_writer",
                     help="which server writer you wish to use:lua xml json" )
    parser.add_option( "-l","--cwriter", dest="clt_writer",
                     help="which client writer you wish to use:lua xml json" )

    options, args = parser.parse_args()

    loader = Loader( options.input_path,options.srv_path,options.clt_path,
        options.timeout,options.suffix,options.srv_writer,options.clt_writer )
    loader.attention()
    loader.load()