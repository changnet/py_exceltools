# py_exceltools
基于openpyxl的excel转换工具。支持xlsx文件转换为lua、xml、json等配置文件。
由于解析库从xlrd更换为openpyxl，不再支持xls文件的转换。

关于openpyxl库：https://pypi.python.org/pypi/openpyxl。

此工具兼容python2和python3。


## 安装

linux安装  

    apt-get install python-pip
    pip install openpyxl


win安装

    安装python(同时安装pip并添加到Path):https://www.python.org/downloads/windows/
    安装openpyxl,在cmd中运行:pip install openpyxl

# 使用
    lancher.bat(win)和lancher.sh(linux)为对应脚本，并且配置了对应的参数来转换
    example.xlsx，生成的配置文件分别在server、client文件夹中。

    参数：
    --input   ：需要转换的excel文件所在目录
    --srv     : 服务端配置文件输出目录
    --clt     : 客户端配置文件输出目录
    --timeout : 需要转换的excel文件最后更新时间距当前时间秒数。-1转换所有
    --suffix  ：excel文件后缀，通常为.xlsx 
    --swriter : 服务端配置文件转换器，可以指定为lua、json、xml
    --cwriter : 客户端配置文件转换器，可以指定为lua、json、xml

# 建议
* 在string中无法直接使用换行等特殊称号。请用\n等转义字符替代。
* 设置表结构时，数据尽量偏平化。例如一个玩家身上有8种装备，不应该配8个表，而是在装备表中加一个字段pos
* 由于xml并不存在数组等结构，不建议使用。
* 工具会检测server和client标识。如果不存在，则不导出些表。方便策划做备注

# TODO
* 如何导出零碎的配置。例如：
```lua
level = 90,
system_id = 1,


