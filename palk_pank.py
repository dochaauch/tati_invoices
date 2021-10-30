import pprint
import datetime
import pandas as pd
from textwrap import wrap
import re
from sys import argv
import conf
from datetime import date, timedelta
from pyautogui import typewrite


def check_if_not_nan(str_):
    if not pd.isna(str_):
        return str_
    else:
        return ''


def not_use_if_nan(first_, str_, last_):
    if len(check_if_not_nan(str_)) > 0:
        return f"{first_}{str_}{last_}"
    else:
        return ''


def del_empty_line(str_):
    emp = re.compile(r'^\n\r$')
    for s in str_:
        if emp.match(s):
            continue
        else:
            return s


your_target_folder = conf.your_target_folder
my_file = rf'{your_target_folder}\Union Salary\salary_template.xlsx'
my_check = rf'{your_target_folder}\Union Salary\konto_check.xlsx'
company_name = 'UNIONCOMPANY OÜ'
our_iban = 'EE102200221076946964'
our_bic = 'HABAEE2X'
our_reg_nr = '16239200'

#данные из bat файла
#date_of_payment = str(argv[1])
#month_of_calc = str(argv[2])

date_of_payment = date.today().strftime("%Y-%m-%d")
month_of_calc = (date.today().replace(day=1) - datetime.timedelta(days=1)).strftime("%b%y").upper()

#print("enter folder name: ")
#typewrite("Default Value")
#folder = input()

print('дата платежа:')
typewrite(date_of_payment)
date_of_payment = input()

print('месяц зарплаты и командировок (предыдущий месяц от даты платежа):')
typewrite(month_of_calc)
month_of_calc = input()

print('дата платежа:', date_of_payment)
print('месяц зарплаты и командировок (предыдущий месяц от даты платежа):', month_of_calc)

#читаем эксель с данными
df = pd.read_excel(my_file)

# заменить пустые значения в зарплате и командировке
df = df.fillna({'Scht': 0, 'Kom': 0})

#читаем эксель с проверенными счетами
df_check = pd.read_excel(my_check)

#соединяем данные из файла с зарплатой с контрольным файлом
df_merge = pd.merge(df, df_check, on=['Name'], how='inner')
df_merge = df_merge.fillna('_')

#добавляем столбцы со сравнением
df_merge['Acc_com'] = df_merge['Bank account'] == df_merge['Account']
df_merge['Bank_com'] = df_merge['Address of the bank'] == df_merge['Bank']
df_merge['SWIFT_com'] = df_merge['Swift cod bank'] == df_merge['SWIFT']
df_merge['home_com'] = df_merge['Worker&apos;s address'] == df_merge['Home_address']

#создаем датафрейм только со столбцами сравнения
diff_df = df_merge[['Name', 'Acc_com', 'Bank_com', 'SWIFT_com', 'home_com']]
list_of_col = ['Acc_com', 'Bank_com', 'SWIFT_com', 'home_com']
diff_df = diff_df.copy()
diff_df['Sum_of_true'] = diff_df.loc[:, list_of_col].sum(axis=1)
#создаем датафрейм с несовпадениями
false_df = diff_df[diff_df['Sum_of_true'] < 4]

if len(false_df) > 0:
    print('данные не сходятся с контрольной базой:')
    print(false_df)

# ищем, кого нет в контрольной базе
df_not_in_base = df[~df.Name.isin(df_check.Name)]
if len(df_not_in_base) > 0:
    print()
    print('Работника нет в контрольной базе')
    print(df_not_in_base.Name)

d_now = datetime.datetime.now()
d_00 = d_now.strftime('%d.%m.%y')
d_01 = d_now.strftime('%Y-%m-%d')

xml_first_part = f'''<?xml version='1.0' encoding='UTF-8'?>
<Document xmlns='urn:iso:std:iso:20022:tech:xsd:pain.001.001.03' xmlns:xsi='http://www.w3.org/2001/XMLSchema-instance'>
<CstmrCdtTrfInitn>
<GrpHdr>
<MsgId>{d_00}</MsgId>
<CreDtTm>{d_01}T00:00:00.00</CreDtTm>
<NbOfTxs>{df['Total'].count()}</NbOfTxs>
<CtrlSum>{sum(df['Total']):.2f}</CtrlSum>
<InitgPty>
<Nm>{company_name}</Nm>
</InitgPty>
</GrpHdr>'''

xml_row = ''
xml_rows = ''
country_dict = {'UA': ['Ukraine', 'Ukraina'],
                'RU': ['Russia'],
                }
for index, row in df.iterrows():
    # owner
    another_name = re.compile(r'.*\((.*)\)')
    if another_name.search(row['Name']):
        owner = another_name.search(row['Name']).group(1).strip()
        person = row['Name'].split('(')[0] + ' '
    else:
        owner = row['Name'].strip()
        person = ''

    # address of worker, county of residence
    country_adr = ''
    for k, v_list in country_dict.items():
        if pd.notna(row['Worker&apos;s address']):
            for v in v_list:
                if v.lower() in row['Worker&apos;s address'].lower():
                    country_adr = k
    if pd.notna(row['Worker&apos;s address']) and country_adr == '':
        print('Страна не найдена.', owner, row['Worker&apos;s address'])
    worker_adr_l = (row['Worker&apos;s address'].split(',')[1:]
                    if not pd.isna(row['Worker&apos;s address']) else '')
    worker_adr_l = [wa.strip() for wa in worker_adr_l]
    worker_adr = ', '.join(worker_adr_l)
    worker_adr0 = ''
    if len(worker_adr) > 70:
        worker_adr, worker_adr0 = wrap(worker_adr, 70)

    # IBAN or not IBAN
    iban = ''
    non_iban = ''
    if row['Bank account'].isnumeric():
        account = f"<Othr><Id>{row['Bank account']}</Id></Othr>"
    else:
        account = f"<IBAN>{row['Bank account']}</IBAN>"

    # если сумма зарплаты и командировки не равна переводу
    if row['Scht'] + row['Kom'] != row['Total']:
        print(f"{row['Name']}: cумма зп {row['Scht']} + сумма командировки {row['Kom']} = "
              f"{row['Scht'] + row['Kom']} не сходится с Total {row['Total']}")

    xml_row = f'''<PmtInf>
<PmtInfId>PMTID001</PmtInfId>
<PmtMtd>TRF</PmtMtd>
<NbOfTxs>1</NbOfTxs>
<CtrlSum>{row['Total']:.2f}</CtrlSum>
<PmtTpInf>
<SvcLvl>
<Cd>SEPA</Cd>
</SvcLvl>
</PmtTpInf>
<ReqdExctnDt>{date_of_payment}</ReqdExctnDt>
<Dbtr>
<Nm>{company_name}</Nm>
<Id>
<OrgId>
<Othr>
<Id>{our_reg_nr}</Id>
</Othr>
</OrgId>
</Id>
</Dbtr>
<DbtrAcct>
<Id>
<IBAN>{our_iban}</IBAN>
</Id>
<Ccy>EUR</Ccy>
</DbtrAcct>
<DbtrAgt>
<FinInstnId>
<BIC>{our_bic}</BIC>
</FinInstnId>
</DbtrAgt>
<CdtTrfTxInf>
<PmtId>
<InstrId>{index + 1}</InstrId>
<EndToEndId>{index + 1}</EndToEndId>
</PmtId>
<Amt>
<InstdAmt Ccy='EUR'>{row['Total']:.2f}</InstdAmt>
</Amt>
<CdtrAgt>
<FinInstnId>
{not_use_if_nan(r"<BIC>", row['Swift cod bank'], r"</BIC>")}
</FinInstnId>
</CdtrAgt>
<Cdtr>
<Nm>{owner}</Nm>
<PstlAdr>
{not_use_if_nan(r"<Ctry>", country_adr, r"</Ctry>")}
{not_use_if_nan(r"<AdrLine>", f"{worker_adr}", r"</AdrLine>")}
{not_use_if_nan(r"<AdrLine>", f"{worker_adr0}", r"</AdrLine>")}
</PstlAdr>
{not_use_if_nan(r"<CtryOfRes>", country_adr, r"</CtryOfRes>")}
</Cdtr>
<CdtrAcct>
<Id>
{account}
</Id>
</CdtrAcct>
<RmtInf>
<Ustrd>{month_of_calc}: {person}palk {row['Scht']:.2f} + komandeeringuraha {row['Kom']:.2f}</Ustrd>
</RmtInf>
</CdtTrfTxInf>
</PmtInf>'''
    xml_rows += xml_row + '\n'

xml_rows = xml_rows.strip()
xml_end = '''</CstmrCdtTrfInitn>
</Document>'''

result = f'''{xml_first_part}
{xml_rows}
{xml_end}'''

#result = del_empty_line(result)

#print(result)

with open(fr'{your_target_folder}/Union Salary/pank.xml', 'w', encoding='utf-8') as f:
    f.write(result)
