#! python
# -*- coding:utf-8 -*-

import os
import sys
import json

class Writer:

    def __init__(self):
        pass

    # 文件后缀
    def suffix(self):
        return ".json"

    # 文件内容(字符串)
    def context(self,ctx):
        return json.dumps(ctx,ensure_ascii=False,\
            indent=4,sort_keys=True,separators=(',', ':') )
