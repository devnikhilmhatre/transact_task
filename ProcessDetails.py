import pandas as pd
import numpy as np
import xlsxwriter


def readProcessDetailsFile(path):
    excluded = {}
    df = pd.read_excel(path)
    duplicate_records = df[df.duplicated(subset=None, keep='first')]
    no_process = df[pd.isnull(df['Process 1'])]

    _list = [*duplicate_records.index.tolist(), *no_process.index.tolist()]

    df.drop(_list, inplace=True)
    return df, no_process, duplicate_records


if __name__ == "__main__":
    df, no_process, duplicate_records = readProcessDetailsFile(
        'ProcessDetails.xlsx')

    workbook = xlsxwriter.Workbook('ProcessDetails_output.xlsx')
    worksheet = workbook.add_worksheet('Processes')

    # -------------------------------
    duplicate_records_sheet = workbook.add_worksheet('Duplicate Records')

    duplicate_records_sheet.write(
        0,
        0,
        'Following are the Ids of duplicate records',
    )
    rowNo1 = 1
    colNo1 = 0
    for i, row in duplicate_records.iterrows():
        duplicate_records_sheet.write(rowNo1, colNo1, row[0])
        rowNo1 += 1
    # ---------------------------------

    no_process_sheet = workbook.add_worksheet('No Process')

    no_process_sheet.write(
        0,
        0,
        'Following are the Ids which does not have any proccess',
    )
    rowNo1 = 1
    colNo1 = 0
    for i, row in no_process.iterrows():
        no_process_sheet.write(rowNo1, colNo1, row[0])
        rowNo1 += 1
    # ----------------------------------

    df.fillna(value='No Value', inplace=True)

    rowNo = 0
    colNo = 0

    for i, row in df.iterrows():

        if (row[1] != 'No Value'):
            worksheet.write(rowNo, colNo, row[0])
            worksheet.write(rowNo, colNo + 1, 1)
            worksheet.write(rowNo, colNo + 2, row[1])
            rowNo += 1
        if (row[2] != 'No Value'):
            worksheet.write(rowNo, colNo, row[0])
            worksheet.write(rowNo, colNo + 1, 2)
            worksheet.write(rowNo, colNo + 2, row[2])
            rowNo += 1
        if (row[3] != 'No Value'):
            worksheet.write(rowNo, colNo, row[0])
            worksheet.write(rowNo, colNo + 1, 3)
            worksheet.write(rowNo, colNo + 2, row[3])
            rowNo += 1
        if (row[4] != 'No Value'):
            worksheet.write(rowNo, colNo, row[0])
            worksheet.write(rowNo, colNo + 1, 4)
            worksheet.write(rowNo, colNo + 2, row[4])
            rowNo += 1

    workbook.close()
