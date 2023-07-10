import yagmail

subj = 'документы для Испании, январь 2022. test'
nl = '\n'
img_ = r'C:\Users\docha\OneDrive\Leka\Mailing\QR Code example.png'
body = f"Здравствуйте.{nl}В приложении ваш пропуск.{nl}{nl}" \
       f"Важное требование от заказчика - у каждого работника с собой должен быть паспорт.{nl}{nl}" \
       f"Не ИД карта, не копия документа, а обязательно именно{nl}{nl}" \
       f"ПАСПОРТ.{nl}{nl}" \
       f"Убедитьесь, что вы получили QR код для въезда в Испанию.{nl}" \
       f"{nl}{nl}Best Regargs{nl}Tatjana"
img__ = yagmail.inline(r'C:\Users\docha\OneDrive\Leka\Mailing\QR Code example.png')
#f"{[yagmail.inline(img_)]}" \
#f"{[yagmail.inline(img_)]}" \

yag = yagmail.SMTP(user='jelena.jakobson@gmail.com', password='rinhon-tewqy6-vaHved',
                           host='smtp.gmail.com')
#yag = yagmail.SMTP(user='tanja.svcentre@gmail.com', password='Danila11',
#                           host='smtp.gmail.com')
list_of_attach = [r'C:\Users\docha\OneDrive\Leka\Mailing\QR Code example.png']
yag.send(
   #to='jelena.jakobson@gmail.com',
   #cc='',
   bcc=['archive.jakobson@gmail.com'],
   subject=subj,
   contents=[body, img__],
   #attachments=list_of_attach,
       )