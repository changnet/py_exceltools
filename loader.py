#! python
# -*- coding:utf-8 -*-

import os
import importlib
from optparse import OptionParser

try:
    basestring
except NameError:
    basestring = str

class Loader:
    # @input_path:excel文件所在目录
    # @srv_path  :server输出目录
    # @clt_path  :客户端输出目录
    # @timeout   :只处理文档最后更改时间在N秒内的文档
    # @suffix    :excel文件后缀
    def __init__(self,input_path,srv_path,clt_path,timeout,suffix):
        self.input_path = input_path
        self.srv_path   = srv_path
        self.clt_path   = clt_path
        self.timeout    = timeout
        self.suffix     = suffix
        #importlib.import_module("matplotlib.text")

    def done(self):
        print("all done... %d error %d warning" % (self.errors,self.warns))

    def attention(self):
        print("********excel转换********")
        print("**第一行为数据类型，支持number、string、json")
        print("**第二行为server字段名，可为空")
        print("**第三行为client字段名，可为空")
        print("**第一列为主键，不可重复。可以为字符串，为空则转换为数组")
        print("***************************************************\n")

    def load(self):
        print("load %s files from %s within %d seconds" 
            % (self.suffix,self.input_path,self.timeout))

        suffix = "." + self.suffix
        file_list = os.listdir( options.input_path )
        for file in file_list:
            if ( os.path.isfile( file ) and 
            ("" == self.suffix or file.endswith( suffix )) ):
                self.load_one( file )
    
    def load_one(self,file):
        print( file )

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

    options, args = parser.parse_args()

    loader = Loader( options.input_path,
        options.srv_path,options.clt_path,options.timeout,options.suffix )
    loader.load()