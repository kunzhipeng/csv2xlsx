# csv2xlsx
A file converter for CSV to XLSX, written in Python.

一个用Python实现的CSV格式转XLSX格式文件转换器。

We are a webscraping company since 2010 (you can learn more about us on our website http://www.site-digger.com). In our everyday work, we ofen convert the output CSV files into XlSX format for our clients, this process is cumbersome, so we create this script. Now we share this tool here to everyone who need it.

我们是一家超过10年历史的Web数据采集公司，公司网站是http://www.site-digger.com 。 在我们的日常工作中，经常需要将产出的CSV格式文件转换为XLSX格式再交付给客户，这个过程很繁琐，所以我们编写了这个程序。现在公开分享它，任何需要的人可以使用。


## Usage/用法

```
Usage: csv2xlsx.py path-of-input-csv-file [-e <file_encoding>] [-n <max_rows_to_convert>]

Options:
  -h, --help            show this help message and exit
  -e FILE_ENCODING, --file_encoding=FILE_ENCODING
                        The encoding of the input file.
  -n MAX_ROWS, --max_rows=MAX_ROWS
                        The max rows to convert.
```

若以源码形式运行，请先安装依赖库xlsxwriter：

`pip install xlsxwriter`

Windows下可用直接下载release里使用pyinstaller打包好的exe文件。将要转换的CSV文件拖拽到csv2xlsx.exe上即可自动完成转换。
