import pandas as pd

excel_file_mail = r"C:\Users\docha\OneDrive\Leka\Mailing\Mail list.xlsx"
csv_file_mail = r'C:\Users\docha\OneDrive\Leka\Mailing\mail_list.csv'

read_file = pd.read_excel(excel_file_mail)
read_file.to_csv(csv_file_mail, index=None, header=True)