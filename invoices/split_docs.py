import os
import PyPDF2
import pdfplumber
import re
import conf


def choose_user():
    print('docha (d), tati (нажать любую клавишу):')
    who = input()
    if who == 'd':
        your_target_folder = conf.your_target_folder_docha
    else:
        your_target_folder = conf.your_target_folder
    return your_target_folder


def list_of_files(your_target_folder):
    pdf_files = []
    for dirpath, _, filenames in os.walk(your_target_folder):
        for items in filenames:
            file_full_path = os.path.abspath(os.path.join(dirpath, items))
            if file_full_path.lower().endswith('.pdf'):
                pdf_files.append(file_full_path)
    return pdf_files


def extract_nr_of_doc(text):
    #nr_of_doc = 'notext'
    nr_of_doc_akt = re.compile(r'№ акта\.:\s+\b(.*)\b')
    nr_of_doc_inv = re.compile(r'Invoice nr:\s+\b(.*)\b')
    nr_of_doc = nr_of_doc_akt.search(text)
    if nr_of_doc:
        nr_of_doc = nr_of_doc.group(1)
        nr_of_doc = nr_of_doc.replace('/', '_')
        return nr_of_doc
    else:
        nr_of_doc = nr_of_doc_inv.search(text)
        nr_of_doc = nr_of_doc.group(1)
        nr_of_doc = nr_of_doc.replace('/', '_')
        return nr_of_doc


def create_fold_name(nr_of_doc):
    if nr_of_doc.endswith("A"):
        return nr_of_doc[:-1]
    return fr'{nr_of_doc}'


def extract_ship_name(text):
    #ship_name = 'object'
    inv_ship = re.compile(r'Ship:\s+(.*)')
    akt_ship = re.compile(r'Объект:\s+(.*)')

    ship_name = akt_ship.search(text)
    if ship_name:
        ship_name = ship_name.group(1).strip()
    else:
        ship_name = inv_ship.search(text)
        if ship_name:
            ship_name = ship_name.group(1).strip()

    if ship_name is None or 'внутренние' in ship_name:
        ship_name = 'object'

    return ship_name


def processing(target_folder):
    pdf_files = list_of_files(target_folder)

    pdf_files.sort(key=str.lower)
    #pdfWriter = PyPDF2.PdfFileWriter()
    folder_text = ''

    for files_address in pdf_files:
        print(files_address)
        pdfFileObj = open(files_address, 'rb')
        pdfReader = PyPDF2.PdfFileReader(pdfFileObj, strict=False)

        for page in range(pdfReader.getNumPages()):
            pdf_writer = PyPDF2.PdfFileWriter()
            pdf_writer.addPage(pdfReader.getPage(page))

            with pdfplumber.open(files_address) as pdf:
                page_pdf = pdf.pages[page]
                text = page_pdf.extract_text()
                nr_of_doc = extract_nr_of_doc(text)
                fold_name = create_fold_name(nr_of_doc)
                ship_name = extract_ship_name(text)

            output_filename, folder_text = create_subfolder(target_folder,
                                               fold_name, ship_name, folder_text, nr_of_doc)

            write_to_pdf(output_filename, pdf_writer)
    return folder_text


def create_subfolder(your_target_folder, fold_name, ship_name, folder_text, nr_of_doc):
    try:
        os.mkdir(f'{your_target_folder}/{fold_name} {ship_name}')
        print(f'{your_target_folder}/{fold_name} {ship_name}')
        folder_text += f'{fold_name} {ship_name}' + '\n'
    except FileExistsError:
        pass
    if nr_of_doc == 'notext':
        output_filename = your_target_folder + '/{}.txt'.format(nr_of_doc)
    else:
        output_filename = your_target_folder + '/{} {}/{}.pdf'.format(fold_name, ship_name, nr_of_doc)
    return output_filename, folder_text


def write_to_pdf(output_filename, pdf_writer):
    with open(output_filename, 'wb') as out:
        try:
            pdf_writer.write(out)
        except (Exception, ):
            pass
    # print('Created: {}'.format(output_filename))


def note_folder_to_be_deleted_to_txt(folder_text, your_target_folder):
    print(folder_text)
    with open(fr'{your_target_folder}/folder.txt', 'w') as f:
        f.write(folder_text)



def main():
    your_target_folder = choose_user()
    folder_text = processing(your_target_folder)
    note_folder_to_be_deleted_to_txt(folder_text, your_target_folder)


if __name__ == "__main__":
    main()


