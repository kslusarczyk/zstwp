from xlrd import *
import csv

class Alarms_Correlation:

    def __init__(self):
        self.xls_file = 'HuaweiU2000Export_sorted_by_event_time.xlsx'
        self.csv_name = 'HuaweiU2000Export_sorted_by_event_time.csv'
        self.sheet_name = 'AlarmExport'


    def csv_from_excel(self):
        b = open_workbook(self.xls_file)
        s = b.sheet_by_name(self.sheet_name)
        bc = open(self.csv_name, 'w')
        bcw = csv.writer(bc, csv.excel)
        for row in range(s.nrows):
            this_row = []
            for col in range(s.ncols):
                val = s.cell_value(row, col)
                if isinstance(val, unicode):
                    val = val.encode('utf8')
                this_row.append(val)
            bcw.writerow(this_row)


alarms_correlation = Alarms_Correlation()
alarms_correlation.csv_from_excel()