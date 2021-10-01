import os
import PyPDF2
import pdfplumber
import re
import conf


your_target_folder = conf.your_target_folder


def list_of_files(your_target_folder):
    pdf_files = []
    for dirpath, _, filenames in os.walk(your_target_folder):
        for items in filenames:
            file_full_path = os.path.abspath(os.path.join(dirpath, items))
            if file_full_path.lower().endswith('.pdf'):
                pdf_files.append(file_full_path)
    return pdf_files


pdf_files = list_of_files(your_target_folder)

pdf_files.sort(key=str.lower)
pdfWriter = PyPDF2.PdfFileWriter()
folder_text = ''

for files_address in pdf_files:
    nr_of_doc = 'notext'
    inv_ship_name = ''
    pdfFileObj = open(files_address, 'rb')
    pdfReader = PyPDF2.PdfFileReader(pdfFileObj, strict=False)

    for page in range(pdfReader.getNumPages()):
        pdf_writer = PyPDF2.PdfFileWriter()
        pdf_writer.addPage(pdfReader.getPage(page))

        with pdfplumber.open(files_address) as pdf:
            page_pdf = pdf.pages[page]
            text = page_pdf.extract_text()
            #print(text)
            nr_of_doc_akt = re.compile(r'№ акта\.:\s+\b(.*)\b')
            nr_of_doc_inv = re.compile(r'Invoice nr:\s+\b(.*)\b')
            inv_ship = re.compile(r'Ship:\s+(.*)')
            akt_ship = re.compile(r'Объект:\s+(.*)')
            try:
                nr_of_doc = nr_of_doc_akt.search(text)
                if nr_of_doc:
                    nr_of_doc = nr_of_doc.group(1)
                    nr_of_doc = nr_of_doc.replace('/', '_')
                    if nr_of_doc.endswith('A'):
                        ship_name = akt_ship.search(text)
                        ship_name = ship_name.group(1).strip()
                        fold_name = fr'{nr_of_doc[:-1]}'

                else:
                    nr_of_doc = nr_of_doc_inv.search(text)
                    nr_of_doc = nr_of_doc.group(1)
                    nr_of_doc = nr_of_doc.replace('/', '_')

                    ship_name = inv_ship.search(text)
                    ship_name = ship_name.group(1).strip()
                    fold_name = fr'{nr_of_doc}'
            except (Exception, ):
                pass

        try:
            os.mkdir(f'{your_target_folder}/{fold_name} {ship_name}')
            print(fold_name, ship_name)
            folder_text += f'{fold_name} {ship_name}' + '\n'
        except FileExistsError:
            pass
        if nr_of_doc == 'notext':
            output_filename = your_target_folder + '/{}.txt'.format(nr_of_doc)
        else:
            output_filename = your_target_folder + '/{} {}/{}.pdf'.format(fold_name, ship_name, nr_of_doc)

        with open(output_filename, 'wb') as out:
            try:
                pdf_writer.write(out)
            except (Exception, ):
                pass

        # print('Created: {}'.format(output_filename))
with open(fr'{your_target_folder}/folder.txt', 'w') as f:
    f.write(folder_text)
# print(folder_text)
