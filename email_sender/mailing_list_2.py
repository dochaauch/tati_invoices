import pandas as pd
import yagmail
import csv
import os
import re

your_target_folder = r'C:\Users\docha\OneDrive\Leka\Mailing\Tickets'

excel_file_mail = r"C:\Users\docha\OneDrive\Leka\Mailing\Mail list.xlsx"
csv_file_mail = r'C:\Users\docha\OneDrive\Leka\Mailing\mail_list.csv'


read_file = pd.read_excel(excel_file_mail)
read_file.to_csv(csv_file_mail, index=None, header=True)

nl = '\n'
body = f"Здравствуйте.{nl}Внимание!{nl}" \
       f"В приложении новое место встречи в Аэропорту Аликанте для прибывающих 19.01.22{nl}{nl}" \
       f"кафе Deli&Cia Coffe" \
       f"{nl}{nl}Best Regargs{nl}Tatjana"

subj = 'Документы для Испании, январь 2022. Новое место встречи 19.01.22'

list_attachments = [r'C:\Users\docha\OneDrive\Leka\Mailing\Новое место встречи.pdf',]
list_copy_bcc = ['oleksandr.ivanov89@mail.ru', 'uniondocument@mail.ru', 'tanja.svcentre@gmail.com',]

# перебор списка эмейлов
with open(csv_file_mail) as file:
    reader = csv.DictReader(file, delimiter=',')
    for row in reader:
        bcc = row['e-mail']
        list_copy_bcc.append(bcc)

#print(subj, body, list_attachments)
#print(list_copy_bcc)
#print(len(list_copy_bcc))

# yag = yagmail.SMTP(user='tanja.svcentre@gmail.com', password='Danila11',
#                     host='smtp.gmail.com')
#                     #использует пароль приложения, двухфакторная аутентификация
# yag.send(
#     #to=receiver,
#     #cc=list_cc,
#     bcc=list_copy_bcc,
#     subject=subj,
#     contents=body,
#     attachments=list_attachments,
#         )
