import pandas as pd
import re
import time
import xlrd
import pyexcelerate

def pandas_read_range():
    spefici_range = "B5:B238246"
    # start, end = spefici_range.split(":")
    # start_row, end_row = re.search('\d+', start).group(0), re.search('\d+', end).group(0)
    # start_col, end_col = re.search('[a-zA-Z]+', start).group(0), re.search('[a-zA-Z]+', end).group(0)
    # data = pd.read_excel('dataset.xls', sheet_name='1', usecols=(start_col + ":" + end_col),
    #                      nrows=int(end_row) - int(start_row), skiprows=int(start_row) - 1)
    # print(data)


start_time = time.time()
pandas_read_range()
print("--- %s seconds ---" % (time.time() - start_time))

import openpyxl
from openpyxl import load_workbook


#
def open_py_excel():
    wb = .load_workbook('asd.xlsx')
    spefici_range = "A5:BB238246"
    start, end = spefici_range.split(":")
    sheet = wb.worksheets[3]
    cell_range = sheet[start:end]
    df = []
    # print(cell_range)
    for cell in cell_range:
        df.append(cell[0].value)
    df = pd.DataFrame(df)
    print(df)
    # df.head()

#
# print('-----------------------')
# start_time = time.time()
# open_py_excel()
print("--- %s seconds ---" % (time.time() - start_time))
# openpyexcel
# write direct to workbook
# choosing some range
# and save as excel file
# and check time
