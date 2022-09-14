import os
from datetime import datetime

import comtypes.client
import docx
import pandas as pd

src = 'U:/Документы ИТ/МТО/Заявки'
dst = 'U:/Документы ИТ/МТО/Заявки/Реестр'
arc = 'U:/Документы ИТ/МТО/Заявки/Архив'
file_list = []


def convert_to_docx(file_doc, c_time, m_time):
    abs_src = r'U:\Документы ИТ\МТО\Заявки\\'
    word = comtypes.client.CreateObject('Word.Application')
    doc = word.Documents.Open(f'{abs_src}{file_doc}')
    doc.SaveAs(f'{abs_src}{os.path.splitext(file_doc)[0]}', FileFormat=16)
    doc.Close()
    os.remove(f'{abs_src}{file_doc}')
    file_docx = file_doc.replace('.doc', '.docx')
    os.utime(f'{abs_src}{file_docx}', (c_time, m_time))
    return file_docx


def make_file_list(f_docx, c_date):
    doc = docx.Document(f'{src}/{f_docx}')
    cost = 'Стоимость не указана'
    for paragraph in doc.paragraphs:
        if 'Предположительная стоимость заявки:' in paragraph.text:
            cost = paragraph.text.split(':')[-1].strip()

    file_docx = os.path.splitext(f_docx)[0]
    if '-' in file_docx:
        status = 'заявка запланирована в работу'
    elif file_docx.count('+') == 1:
        status = 'заявка подписана и передана в мто'
    elif file_docx.count('+') == 2:
        status = 'заявка отработана, есть договор и счет'
    elif file_docx.count('+') == 3:
        status = 'счет подписан и передан в оплату'
    elif file_docx.count('+') == 4:
        status = 'товар поставлен'
    elif '=' in file_docx:
        status = 'заявка готовится к торгам'
    else:
        status = 'статус не присвоен'

    file_list.append([file_docx.lower(), c_date.strftime('%Y.%m.%d'), cost, status])


def transfer_to_archive(f_docx):
    if not os.path.isdir(f'{arc}/{date.year}'):
        os.mkdir(f'{arc}/{date.year}')
    os.replace(f'{src}/{f_docx}', f'{arc}/{date.year}/{date.strftime("%Y.%m.%d")} - {f_docx.split("+", 5)[-1].strip()}')


for file in os.listdir(src):
    created_time = os.path.getctime(f'{src}/{file}')
    modified_time = os.path.getmtime(f'{src}/{file}')
    date = datetime.fromtimestamp(modified_time)

    if file.lower().endswith('.doc') and '~' not in file:
        file = convert_to_docx(file, created_time, modified_time)
        make_file_list(file, date)

    elif file.lower().endswith('.docx') and '~' not in file:
        if file.count('+') < 5:
            make_file_list(file, date)
        elif file.count('+') == 5:
            transfer_to_archive(file)

    else:
        print(f'rejected: {file}')

df = pd.DataFrame(file_list, columns=['Заявка', 'Дата', 'Стоимость', 'Статус']) \
    .sort_values(by='Статус', ascending=True)

with pd.ExcelWriter(f'{dst}/{datetime.now().year}.xlsx', engine='xlsxwriter') as wb:
    df.to_excel(wb, sheet_name='Реестр', index=False, startrow=1, startcol=1)
    sheet = wb.sheets['Реестр']

    sheet.set_column('B:B', 50)
    sheet.set_column('C:C', 10)
    sheet.set_column('D:D', 20)
    sheet.set_column('E:E', 33)

print('\nRegistry created')
