import pandas as pd
import xlsxwriter

ITEMNAME = 'Item Name'
RAWMATERAIL = 'Raw material'
LEVEL = 'Level'
QUANTITY = 'Quantity'
UNIT = 'Unit'


def readBOMFile(path):
    output = {}
    df = pd.read_excel(path)
    parent_row = None
    item = None

    for i, row in df.iterrows():

        if item is None or row[ITEMNAME] != item:
            item = row[ITEMNAME]
            parent_row = None

        if parent_row is None or int(
                row[LEVEL][-1]) != int(parent_row[LEVEL][-1]) + 1:
            parent_row = row

        material = (row[RAWMATERAIL], row[QUANTITY], row[UNIT])

        if int(row[LEVEL][-1]) > 1:
            value = output.get(parent_row[RAWMATERAIL], None)
            if value:
                value.append(material)
            else:
                output[parent_row[RAWMATERAIL]] = [material]
        else:
            value = output.get(row[ITEMNAME], None)
            if value:
                value.append(material)
            else:
                output[row[ITEMNAME]] = [material]

    return output


if __name__ == "__main__":
    output = readBOMFile('BOM.xlsx')
    workbook = xlsxwriter.Workbook('BOM_output.xlsx')

    for key in output:

        worksheet = workbook.add_worksheet(key)

        worksheet.write('A1', 'Finished Good List')
        worksheet.write('A2', '#')
        worksheet.write('B2', 'Item Description')
        worksheet.write('C2', 'Quantity')
        worksheet.write('D2', 'Unit')

        worksheet.write('A3', 1)
        worksheet.write('B3', key)
        worksheet.write('C3', 1)
        worksheet.write('D3', 'Pc')

        worksheet.write('A4', 'END of FG')
        worksheet.write('A5', 'Raw Material List')

        worksheet.write('A6', '#')
        worksheet.write('B6', 'Item Description')
        worksheet.write('C6', 'Quantity')
        worksheet.write('D6', 'Unit')

        row = 6
        col = 0
        for i, value in enumerate(output[key]):
            worksheet.write(row, col, i + 1)
            worksheet.write(row, col + 1, value[0])
            worksheet.write(row, col + 2, value[1])
            worksheet.write(row, col + 3, value[2])
            row += 1
        worksheet.write(row, col, 'End of RM')

    workbook.close()
