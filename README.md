# py_exceltools
convert excel to lua、xml、json with python

# openpyxl
A Python library to read/write Excel 2010 xlsx/xlsm files.https://pypi.python.org/pypi/openpyxl
## install
* apt-get install python-pip
* pip install openpyxl

# usage
* lancher.bat(win) or lancher.sh(linux) to run program
* check server、client directory to check output file.

# notice
* 在string中无法直接使用换行。请用\n替代。
* 数据偏平化。例如一个玩家身上有8种装备，不应该配8个表，而是在装备表中加一个字段pos

# TEST
* 测试10行以后超长字符串是否会被截断
* 测试数据类型与单元格不匹配是否报错
* 测试key为默认数组、int、number、string是否正确

# TODO
* 如何导出零碎的配置。例如：
```lua
level = 90,
system_id = 1,
```
* 考虑用[=[ ... ]=]来表示lua中的字符串，避免特殊字符。
* TODO 检测server和client行。如果不存在，则不导出些表。方便策划做备注。
* 测试在string中出现int类型，在int类型中出现string的异常处理

