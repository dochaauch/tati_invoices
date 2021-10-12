import openpyxl
import xlrd
import os
import pprint
import re
from openpyxl.comments import Comment

target_folder = r'C:\Users\docha\OneDrive\Leka\PSB Finstratum\September'
stema_file = r'C:\Users\docha\OneDrive\Leka\PSB Finstratum\September\FINSTRATUM FS 09.xlsx'


def list_of_files(your_target_folder):
    week_files = []
    week_list = []
    for dirpath, _, filenames in os.walk(your_target_folder):
        r = re.compile(r'.*Vko(\d{2}).*')
        for items in filenames:
            if items.lower().startswith('finstratum vko'):
                file_full_path = os.path.abspath(os.path.join(dirpath, items))
                week_nr = r.match(items)[1]
                week_files.append(file_full_path)
                week_list.append(week_nr)
    return week_files, week_list


def change_time_to_norm(str_):
    if str_.value:
        return float(str(str_.value).replace('.3', '.5'))


def read_every_page_week_file(week_file):
    week_book = xlrd.open_workbook(filename=week_file)
    fs_hours = dict()
    for i in range(1, 8):  # перебираем по дню листы каждой книги недели
        sh = week_book.sheet_by_index(i)
        head_row = 1
        #print(os.path.basename(week_file), sh)
        wh_list = []
        for rx in range(head_row + 1, sh.nrows):
            rate = sh.row(rx)[5]
            remarks = sh.row(rx)[10]
            r = re.compile(r'.*((\d)h\sFS).*')  # ищем в примечании часы FS
            if rate.value == 'FS' or r.match(remarks.value):
                fs_name = sh.row(rx)[0]
                start_hour = change_time_to_norm(sh.row(rx)[3])
                end_hour = change_time_to_norm(sh.row(rx)[4])
                if rate.value == 'FS':
                    if end_hour > start_hour:
                        work_hour = end_hour - start_hour - 0.5
                    else:
                        work_hour = 24 - start_hour + end_hour - 0.5
                        print('***работает после 00:00', os.path.basename(week_file), sh, fs_name.value,
                              start_hour, end_hour, work_hour)
                if r.match(remarks.value):
                    work_hour = float(r.match(remarks.value).group(2))
                print(fs_name, start_hour, end_hour, work_hour, rate, remarks)
                wh_list.append(work_hour)
                fs_hours[fs_name] = [wh_list]
                {'Aleksei Ukhanov': [{1: [8, remarks], 2: [8]}, {1: [8], 2: [8]}]}
    return fs_hours


def main():
    week_files, week_list = list_of_files(target_folder)
    print(week_list)
    for cnt_f, w_file in enumerate(week_files):
        fs_hours = read_every_page_week_file(w_file)

    print(fs_hours)


if __name__ == "__main__":
    main()

