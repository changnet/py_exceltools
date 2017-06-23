# py_exceltools
convert excel to lua、xml、json with python

# openpyxl
A Python library to read/write Excel 2010 xlsx/xlsm files.https://pypi.python.org/pypi/openpyxl
## install
* apt-get install python-pip
* pip install openpyxl

# TEST
* 测试10行以后超长字符串是否会被截断
* 测试数据类型与单元格不匹配是否报错
* 测试key为默认数组、int、number、string是否正确

# TODO
* 允许前后端writer不同，分开配置
* basestring版本兼容不需要了
* 考虑用[=[ string ]=]来表示lua中的字符串
* 目录存在检测放到loader
