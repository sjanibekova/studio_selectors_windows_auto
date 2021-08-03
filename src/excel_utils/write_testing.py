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


def pandas_to_excel(location, df):
    writer = pd.ExcelWriter(location, engine='xlsxwriter')
    df.to_excel(writer)
    writer.save()


# Create test data
nRows = 100000
nColumns = 30
df = pd.DataFrame(np.random.randn(nRows, nColumns), columns=['C%02d' % d for d in range(nColumns)])

# Example: Write to excel file
print('Write excel using pyexelerate ...')
pyexecelerate_to_excel('test_file.xlsx', df)

# Example: write to multiple sheets
print('Write to multiple sheets using pyexelerate ...')
wb = pyexcelerate.Workbook()
pyexecelerate_to_excel(wb, df, sheet_name='Sheet 1')
pyexecelerate_to_excel(wb, df, sheet_name='Sheet 2')
wb.save('test_file.xlsx')

# Compare performance
print('Compare the runtime between xlsxwriter and pyexecelerate ...')
pandas_time = timeit.timeit("pandas_to_excel('test.xlsx', df)", number=1, globals=globals())

pyexecelerate_time = timeit.timeit("pyexecelerate_to_excel('test.xlsx', df)", number=1, globals=globals())

print(f'Pandas took {pandas_time} seconds\nPyexecelerate took {pyexecelerate_time} seconds')