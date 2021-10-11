import xlrd
import os
import pprint


def read_hours(file):
    ttl_hour_book = xlrd.open_workbook(filename=file)
    ttl_hours = dict()
    sh = ttl_hour_book.sheet_by_index(0)
    head_row = 4
    first_col_week = 43
    last_col_week = 47
    print(os.path.basename(file))
    pprint.pprint(sh.row_values(head_row)[first_col_week - 1: last_col_week + 2])
    for rx in range(head_row + 1, sh.nrows):
        fs_name = sh.row(rx)[0]
        if len(fs_name.value.split()) >= 2:
            range_hours = sh.row(rx)[first_col_week: last_col_week + 1]
            if fs_name.value:
                list_hours = []
                for l_ in range_hours:
                    list_hours.append(l_.value)
                ttl_hours[fs_name.value.strip()] = list_hours
    return ttl_hours