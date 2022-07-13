# csv2xlsx
A file converter for CSV to XLSX, written in Python. Support large CSV file more than million rows (Will be split into multiple XLSX file, each hold one million rows.)

一个用Python实现的CSV格式转XLSX格式文件转换器。支持超过100万条记录的大型CSV文件（将输出多个XLSX文件，每个100万条）。

We are a webscraping company since 2010 (you can learn more about us on our website http://www.site-digger.com). In our everyday work, we ofen convert the output CSV files into XlSX format for our clients, this process is cumbersome, so we create this script. Now we share this tool here to everyone who need it.

我们是一家超过10年历史的Web数据采集公司，公司网站是http://www.site-digger.com 。 在我们的日常工作中，经常需要将产出的CSV格式文件转换为XLSX格式再交付给客户，这个过程很繁琐，所以我们编写了这个程序。现在公开分享它，任何需要的人可以使用。


## Usage/用法

### 运行Python源码
若以源码形式运行，请先安装依赖库xlsxwriter：

`pip install xlsxwriter`

用法及命令行参数含义如下：
```
Usage: csv2xlsx.py path-of-input-csv-file [-e <file_encoding>] [-p <max_rows_allowed_per_outfile>] [-n <max_rows_to_convert>]

Options:
  -h, --help            show this help message and exit
  -e FILE_ENCODING, --file_encoding=FILE_ENCODING
                        The encoding of the input file. Default is utf-8.
  -p ROWS_PER_FILE, --rows_per_file=ROWS_PER_FILE
                        The max rows allowed for per output XLSX file. If it
                        exceeds, it will be split into multiple files.
  -n MAX_ROWS, --max_rows=MAX_ROWS
                        The max rows to convert.
```

### Windows下运行打包好的exe
Windows下可用直接下载[release](https://github.com/kunzhipeng/csv2xlsx/releases)里使用pyinstaller打包好的exe文件。

#### 方法1
将要转换的CSV文件拖拽到csv2xlsx.exe上，即可自动完成转换。

#### 方法2
使用[ContextMenuManager](https://github.com/BluePointLilac/ContextMenuManager)给CSV类型文件添加一个右键菜单项，比如叫“转XLSX格式”。后面在CSV文件上点击这个右键菜单项即可完成转换，长久来说这种方法一劳永逸。

#### 方法3
在命令行下使用。

建议在命令行下使用，支持更多选项，例如指定CSV字符编码，设置最大转换行数。另外，还支持使用通配符一次转换多个CSV文件，例如：
`csv2xlsx.exe *.csv`
