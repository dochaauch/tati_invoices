from fillpdf import fillpdfs
import pprint
import csv
from collections import defaultdict, OrderedDict
import openpyxl
import os
import pdfrw

ANNOT_KEY = '/Annots'  # key for all annotations within a page
ANNOT_FIELD_KEY = '/T'  # Name of field. i.e. given ID of field
ANNOT_FORM_type = '/FT'  # Form type (e.g. text/button)
ANNOT_FORM_button = '/Btn'  # ID for buttons, i.e. a checkbox
ANNOT_FORM_text = '/Tx'  # ID for textbox
ANNOT_FORM_options = '/Opt'
ANNOT_FORM_combo = '/Ch'
SUBTYPE_KEY = '/Subtype'
WIDGET_SUBTYPE_KEY = '/Widget'
ANNOT_FIELD_PARENT_KEY = '/Parent'  # Parent key for older pdf versions
ANNOT_FIELD_KIDS_KEY = '/Kids'  # Kids key for older pdf versions
ANNOT_VAL_KEY = '/V'
ANNOT_RECT_KEY = '/Rect'


def find_options_pdf(input_pdf_path):
    """
вытащено из fillpdfs
    """
    template_pdf = pdfrw.PdfReader(input_pdf_path)
    for Page in template_pdf.pages:
        if Page[ANNOT_KEY]:
            for annotation in Page[ANNOT_KEY]:
                target = annotation if annotation[ANNOT_FIELD_KEY] else annotation[ANNOT_FIELD_PARENT_KEY]
                if annotation[ANNOT_FORM_type] == None:
                    pass
                if target and annotation[SUBTYPE_KEY] == WIDGET_SUBTYPE_KEY:
                    key = target[ANNOT_FIELD_KEY][1:-1]  # Remove parentheses

                    if target[ANNOT_FORM_type] == ANNOT_FORM_button:
                        # button field i.e. a radiobuttons
                        if not annotation['/T']:
                            if annotation['/AP']:
                                keys = annotation['/AP']['/N'].keys()
                                if keys[0]:
                                    if keys[0][0] == '/':
                                        keys[0] = str(keys[0][1:])
                                list_delim, tuple_delim = '-', '^'

                                annotation = annotation['/Parent']
                                options = []
                                for each in annotation['/Kids']:
                                    keys2 = each['/AP']['/N'].keys()
                                    if ['/Off'] in keys:
                                        keys2.remove('/Off')
                                    export = keys2[0]
                                    if '/' in export:
                                        options.append(export[1:])
                                    else:
                                        options.append(export)
                            print(key, options)

def create_new_database_csv(b): #функция для создания новой таблицы, используется только один раз
    #ключи - это первая строка (заголовки), значения - вторая строка
    with open(new_csv, 'w', encoding="utf-8-sig", newline='') as output:
        #encoding sig чтобы при открытии ТОЖЕ отражались все эст.буквы экселем
        w = csv.DictWriter(output, b.keys())
        w.writeheader()
        w.writerow(b)

#*** часть для создания новой таблицы с данными. используется только один раз
file_ = r'C:\Users\docha\OneDrive\Leka\PPA Elamisluba\TEtaotlus_EST_20181108.pdf'
b = fillpdfs.get_form_fields(file_)
#new_csv = r'C:\Users\docha\OneDrive\Leka\PPA LTR\database_ltr.csv'
pprint.pprint(b)
print(len(b))
print(find_options_pdf(file_))
#create_new_database_csv(b)
#***конец части для создания новой таблицы




