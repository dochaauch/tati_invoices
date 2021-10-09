import xlrd
import xlwt
import xlutils
import pprint
import os
import openpyxl
import re

fs_filename = r'C:\Users\docha\OneDrive\Leka\PSB Finstratum\September\FINSTRATUM FS 09.xls'
rmtp_file = r'C:\Users\docha\OneDrive\Leka\PSB Finstratum\September\rmtp  September 2021.xlsx'

fs_book = xlrd.open_workbook(filename=fs_filename)
fs_hours = dict()
sh = fs_book.sheet_by_index(0)
head_row = 4
first_col_week = 43
last_col_week = 47
#print(os.path.basename(fs_filename))
pprint.pprint(sh.row_values(head_row)[first_col_week - 1: last_col_week + 2])
for rx in range(head_row + 1, sh.nrows):
    fs_name = sh.row(rx)[0]
    if len(fs_name.value.split()) >= 2:
        range_hours = sh.row(rx)[first_col_week: last_col_week + 1]
        if fs_name.value:
            list_hours = []
            for l_ in range_hours:
                list_hours.append(l_.value)
            fs_hours[fs_name.value.strip()] = list_hours

pprint.pprint(fs_hours)
#print('Всего работников: ', len(fs_hours))


rmtp_book = openpyxl.load_workbook(filename=rmtp_file)
rmtp_sh = rmtp_book['Evikit']
cnt_wrk = 0
rmtp_list_wrk = []
column_first_week = 5
for row in rmtp_sh.rows:
    rmtp_name = row[0]
    if row[0].value is None:
        pass
    else:
        for i in range(5):
            r = re.compile('^[a-zA-Z ]*$') # строка содержит только буквы и пробелы
            #print(rmtp_name.value, fs_hours.get(rmtp_name.value), len(rmtp_name.value.split()),
            #      r.match(rmtp_name.value))

            if len(rmtp_name.value.split()) == 2 and r.match(rmtp_name.value):
                if rmtp_name.value.strip() in fs_hours.keys():
                    fs_time = fs_hours.get(rmtp_name.value.strip())
                    cell = rmtp_sh.cell(row=rmtp_name.row, column=column_first_week + i * 2)
                    cell.value = fs_time[i]
                    cell.number_format = '0.0'
                    rmtp_list_wrk.append(rmtp_name.value.strip())
                #print(str(rmtp_name.value), rmtp_name, cell, fs_hours.get(rmtp_name.value))
print('Всего занесено в таблицу сотрудников:', len(set(rmtp_list_wrk)))
#pprint.pprint(set(rmtp_list_wrk))
for k, v in fs_hours.items():
    if k not in set(rmtp_list_wrk):
        print('Сотрудник есть в часах FS, но нет в зарплате:', k, v)
rmtp_book.save(rmtp_file)
