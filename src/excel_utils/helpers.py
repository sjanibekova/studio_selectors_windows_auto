import pandas as pd
import numpy as np
import timeit

import pyexcelerate


def pyexecelerate_to_excel(workbook_or_filename, df, sheet_name='Sheet1', origin=(1, 1), columns=True, index=False):
    """
    Write DataFrame to excel file using pyexelerate library
    """
    if not isinstance(workbook_or_filename, pyexcelerate.Workbook):
        location = workbook_or_filename
        workbook_or_filename = pyexcelerate.Workbook()
    else:
        location = None
    worksheet = workbook_or_filename.new_sheet(sheet_name)

    # Account for space needed for index and column headers
    column_offset = 0
    row_offset = 0

    if index:
        index = df.index.tolist()
        ro = origin[0] + row_offset
        co = origin[1] + column_offset
        worksheet.range((ro, co), (ro + 1, co)).value = [['Index']]
        worksheet.range((ro + 1, co), (ro + 1 + len(index), co)).value = list(map(lambda x: [x], index))
        column_offset += 1
    if columns:
        columns = df.columns.tolist()
        ro = origin[0] + row_offset
        co = origin[1] + column_offset
        worksheet.range((ro, co), (ro, co + len(columns))).value = [[*columns]]
        row_offset += 1

    # Write the data
    row_num = df.shape[0]
    col_num = df.shape[1]
    ro = origin[0] + row_offset
    co = origin[1] + column_offset
    worksheet.range((ro, co), (ro + row_num, co + col_num)).value = df.values.tolist()

    if location:
        workbook_or_filename.save(location)


