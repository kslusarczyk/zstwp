from xlrd import *
import csv


def csv_from_excel():
    b = open_workbook('HuaweiU2000Export.xlsx')
    s = b.sheet_by_name('AlarmExport')
    bc = open('HuaweiU2000Export.csv', 'w')
    bcw = csv.writer(bc, csv.excel)
    for row in range(s.nrows):
        this_row = []
        for col in range(s.ncols):
            val = s.cell_value(row, col)
            if isinstance(val, unicode):
                val = val.encode('utf8')
            this_row.append(val)
        bcw.writerow(this_row)


# csv_from_excel()
