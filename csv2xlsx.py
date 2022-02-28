# coding: utf-8
# csv2xlsx.py
# 使用xlsxwriter将CSV文件转换为XLSX格式，支持超过100万行的CSV

from __future__ import print_function
import sys
import os
import csv
import re
import xlsxwriter
from optparse import OptionParser


# 每个文件最多允许100万行
# https://support.microsoft.com/en-us/office/excel-specifications-and-limits-1672b34d-7043-467e-8e27-269d656771c3
MAX_ROWS_PER_FILE = 1000000

def convert(csv_file, encoding='utf-8', max_rows=0):
    """将CSV文件转换为XLSX格式
       csv_file - 要转换的CSV文件路径;
       encoding - 文件的字符编码;
       max_rows - 最大转换行数;
    """
    if re.compile(r'\.csv$', re.IGNORECASE).search(csv_file):
        if os.path.exists(csv_file):
            output_files = []
            # 输出文件路径
            xlsx_file = csv_file[:-4] + '.xlsx'
            print('Convertting "{}" into "{}"...\n'.format(csv_file, xlsx_file))
            print('The input file encoding is "{}", Will export {} rows.\n'.format(encoding, 'first {}'.format(max_rows) if max_rows else 'ALL'))
            
            header = None
            part = 1
            output_files.append(xlsx_file)
            workbook = xlsxwriter.Workbook(xlsx_file)
            worksheet = workbook.add_worksheet()
            
            # 先获取CSV总行数
            print('Calculating the number of input CSV file rows...\n')
            total_rows = 0 # 总条数
            with open(csv_file, 'rb') as f:
                for r in csv.reader(f):
                    total_rows += 1
            print('Total found {} rows in the input CSV file.\n'.format(total_rows))
            if max_rows and max_rows > total_rows:
                max_rows = total_rows
            num_to_exported = max_rows if max_rows else total_rows

            # 遍历CSV文件转XLSX
            rows_num_processed = 0 # 总条数
            row_num = 0         # 索引行号
            with open(csv_file, 'rb') as f:
                for r in csv.reader(f):
                    rows_num_processed += 1
                    if not header:
                        # 表头
                        header = r
                    if rows_num_processed <= num_to_exported:
                        col_num = 0
                        for c in r:
                            worksheet.write_string(row_num, col_num, c.decode(encoding,'ignore'))
                            col_num += 1
                        row_num += 1
                        percent = round(rows_num_processed * 100 / float(num_to_exported), 1)
                        print('Conversion progress: {}% ({}/{})'.format(percent, rows_num_processed, num_to_exported), end='\r')
                        sys.stdout.flush()
                        # 更新进度条
                        
                        if row_num >= MAX_ROWS_PER_FILE:
                            # 行数超过限制，拆分
                            part += 1
                            workbook.close()
                            # 重建文件
                            new_xlsx_file = xlsx_file[:-5] + '-{}.xlsx'.format(part)
                            output_files.append(new_xlsx_file)
                            workbook = xlsxwriter.Workbook(new_xlsx_file)
                            worksheet = workbook.add_worksheet()
                            # 索引行号重置
                            row_num = 0
                            # 写表头
                            col_num = 0
                            for c in header:
                                worksheet.write_string(row_num, col_num, c.decode(encoding, 'ignore'))
                                col_num += 1
                            row_num += 1
                    else:
                        break
            print('\n\nSaving data into "{}"...'.format(';'.join(output_files)))
            print('\nOutfile is not ready, do not close this window.')
            workbook.close()
            print('\nOutfile "{}" is ready. You can close this window now.'.format(';'.join(output_files)))
        else:
            print('\nCSV file "{}" does not exist.'.format(csv_file))
    else:
        print('\nThe input file must be a .csv file.')
    raw_input('\nPress ENTER to exit.')


if __name__ == '__main__':
    print('*' * 90)
    print('** csv2xlsx V1.0 - File converter for CSV to XLSX.')
    print('** Created by qi@site-digger.com')
    print('** We are specializing in web scraping since 2010, learn more on http://www.site-digger.com.')
    print('*' * 90)
    print('')
    
    parser = OptionParser(usage='%prog path-of-input-csv-file [-e <file_encoding>] [-n <max_rows_to_convert>]')
    parser.add_option('-e', '--file_encoding', dest='file_encoding', default='utf-8', type='string', help='The encoding of the input file.')
    parser.add_option('-n', '--max_rows', dest='max_rows', default=0, type='int', help='The max rows to convert.')
    options, args = parser.parse_args()
    
    if not args:
        parser.print_help()
    else:
        csv_file = args[0]
        convert(csv_file=csv_file, encoding=options.file_encoding, max_rows=options.max_rows)
