import os, stat
import PyPDF2
import conf
import shutil


your_target_folder = conf.your_target_folder

with open(f'{your_target_folder}/folder.txt', 'r') as f:
    folder_names = f.readlines()
    for fn in folder_names:
        folder_name = fn.strip()
        fold_name = folder_name.split(' ')[0]

        list_of_pdfs = [f_ for f_ in os.listdir(f'{your_target_folder}/{folder_name}')
                        if f_.endswith('.pdf')]

        mergeFile = PyPDF2.PdfFileMerger(strict=False)
        inv = fr'{your_target_folder}/{folder_name}/{fold_name}.pdf'
        akt = fr'{your_target_folder}/{folder_name}/{fold_name}A.pdf'

        if not os.path.exists(inv):
            print(folder_name, 'нет счета')
        else:
            list_of_pdfs.remove(os.path.basename(inv))
            mergeFile.append(PyPDF2.PdfFileReader(inv, 'rb'))

        if not os.path.exists(akt):
            print(folder_name, 'нет акта')
        else:
            list_of_pdfs.remove(os.path.basename(akt))
            mergeFile.append(PyPDF2.PdfFileReader(akt, 'rb'))

        if list_of_pdfs:
            for _pdf in list_of_pdfs:
                vedom = fr'{your_target_folder}/{folder_name}/{_pdf}'
                mergeFile.append(PyPDF2.PdfFileReader(vedom, 'rb'))
                os.chmod(vedom, stat.S_IWRITE)
            print(folder_name, 'количество ведомостей: ', len(list_of_pdfs))
        else:
            print(folder_name, 'нет ведомости')

        file_merge = fr'{your_target_folder}/output/{fold_name}.pdf'
        out_fold = fr'{your_target_folder}/output'
        if not os.path.isdir(out_fold):
            os.mkdir(out_fold)
        mergeFile.write(file_merge)
        mergeFile.close()

        mydir_ = fr'{your_target_folder}/{folder_name}'
        shutil.rmtree(mydir_)

notext_file = f'{your_target_folder}/notext.txt'
if os.path.exists(notext_file):
    os.remove(notext_file)
os.remove(f'{your_target_folder}/folder.txt')
