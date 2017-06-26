# py_exceltools
基于openpyxl的excel转换工具。支持xlsx文件转换为lua、xml、json等配置文件。
由于解析库从xlrd更换为openpyxl，不再支持xls文件的转换。

关于openpyxl库：https://pypi.python.org/pypi/openpyxl。


## 安装

linux安装  

    apt-get install python-pip
    pip install openpyxl


win安装

    安装python(同时安装pip并添加到Path):https://www.python.org/downloads/windows/
    安装openpyxl,在cmd中运行:pip install openpyxl

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

