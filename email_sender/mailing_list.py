import pandas as pd
import yagmail
import csv
import os
import re

#https://mailtrap.io/blog/yagmail-tutorial/

your_target_folder = r'C:\Users\docha\OneDrive\Leka\Mailing\Tickets\Shore pass'

excel_file_mail = r"C:\Users\docha\OneDrive\Leka\Mailing\Mail list.xlsx"
csv_file_mail = r'C:\Users\docha\OneDrive\Leka\Mailing\mail_list.csv'

read_file = pd.read_excel(excel_file_mail)
read_file.to_csv(csv_file_mail, index=None, header=True)

#img__ = yagmail.inline(r'C:\Users\docha\OneDrive\Leka\Mailing\QR Code example.png')

nl = '\n'
body = f"Здравствуйте.{nl}В приложении ваш пропуск и инструкция для тех, кто еще не сделал QR код.{nl}{nl}" \
       f"Важное требование от заказчика - у каждого работника с собой должен быть паспорт.{nl}{nl}" \
       f"Не ИД карта, не копия документа, а обязательно именно ПАСПОРТ.{nl}{nl}" \
       f"Убедитесь, что вы получили QR код для въезда в Испанию'.{nl}" \
       f"В приложении образец, как должен выглядеть ваш QR код.{nl}{nl}" \
       f"{nl}{nl}Best Regargs{nl}Tatjana"
namesakes_list = ['Gerov',]

# перебор списка эмейлов
with open(csv_file_mail) as file:
    reader = csv.DictReader(file, delimiter=',')
    for row in reader:
        list_attachments = [r'C:\Users\docha\OneDrive\Leka\Mailing\QR Code.pdf',
                            r'C:\Users\docha\OneDrive\Leka\Mailing\QR Code example.png']

        receiver = row['e-mail']
        perekonnanimi = row['Surname'].strip()
        nimi = row['Name'].strip()
        no_seaman_docs = row['no_seaman']
        no_tickets = row['no_ticket']
        copy_cc = row['copy CC']
        list_cc = copy_cc.split('; ')

        subj = f'{perekonnanimi} {nimi}, документы для Испании, январь 2022. Пропуск, ПАСПОРТ и QR код.'
        print()
        print(receiver, perekonnanimi, nimi)

        for dirpath, _, filenames in os.walk(your_target_folder):
            if perekonnanimi in namesakes_list:
                perek = perekonnanimi.lower()
                nim = nimi[0].lower()
                find_file_ = re.compile(rf'.*({perek},*\s*{nim}).*')
            else:
                find_file_ = re.compile(rf'.*({perekonnanimi.lower()}).*')
            for file_ in filenames:
                our_file_ = find_file_.match(file_.lower())
                if our_file_:
                    list_attachments.append(os.path.join(dirpath, file_))
                    #new_file_name = os.path.join(dirpath, '_' + file_)  #добавить в имя файла _
                    #new_file_name = os.path.join(dirpath, file_[1:]) #имя файла со второго знака без _
                    #old_file_name = os.path.join(dirpath, file_)
                    #print(new_file_name)
                    #os.rename(old_file_name, new_file_name)  # переименовать файлы

        print(list_attachments)

        #if len(list_attachments) < 2:
        #    print(receiver, perekonnanimi, nimi, subj, body)
        #    print(list_attachments)

        yag = yagmail.SMTP(user='tanja.svcentre@gmail.com', password='Danila11',
                           host='smtp.gmail.com')
                            #использует пароль приложения, двухфакторная аутентификация
        yag.send(
            to=receiver,
            cc=list_cc,
            bcc=['archive.jakobson@gmail.com','tanja.svcentre@gmail.com'],
            subject=subj,
            #contents=[body, img__], #очень долго
            contents=body,
            attachments=list_attachments,
                )

        print('done')
