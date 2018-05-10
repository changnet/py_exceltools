#! python
# -*- coding:utf-8 -*-

import os
import sys
import json

class Writer:

    def __init__(self,doc_name,sheet_name):
        self.doc_name   = doc_name
        self.sheet_name = sheet_name

    # 文件后缀
    def suffix(self):
        return ".json"

    # 文件内容(字符串)
    def context(self,ctx):
        return json.dumps(ctx,ensure_ascii=False,\
            indent=4,sort_keys=True,separators=(',', ':') )
