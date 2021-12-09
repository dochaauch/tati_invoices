from fillpdf import fillpdfs
import pprint
import csv
from collections import defaultdict, OrderedDict
import openpyxl
import os


def create_new_database_csv(b): #функция для создания новой таблицы, используется только один раз
    #ключи - это первая строка (заголовки), значения - вторая строка
    with open(new_csv, 'w', encoding="utf-8-sig", newline='') as output:
        #encoding sig чтобы при открытии ТОЖЕ отражались все эст.буквы экселем
        w = csv.DictWriter(output, b.keys())
        w.writeheader()
        w.writerow(b)

#*** часть для создания новой таблицы с данными. используется только один раз
file_ = r'C:\Users\docha\OneDrive\Leka\PPA Elamisluba\ESTandmedperekonnaliikmetekohta_EST.pdf'
b = fillpdfs.get_form_fields(file_)

#new_csv = r'C:\Users\docha\OneDrive\Leka\PPA LTR\database_ltr.csv'
pprint.pprint(b)
print(len(b))
print(c)
#create_new_database_csv(b)
#***конец части для создания новой таблицы