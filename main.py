import datetime
import os

import pandas as pd

src = 'U:/Документы ИТ/МТО/Заявки/2022'
dst = 'U:/Документы ИТ/МТО/Заявки/tst'
ext = ['doc', 'docx']

file_list = []
for file in os.listdir(src):
    if file.lower().endswith(tuple(ext)) and '~$' not in file:
        time = os.path.getctime(f'{src}/{file}')
        date = datetime.datetime.fromtimestamp(time)
        file = file.split('-')[1]
        file = file.split('.')[0]
        if file.count('+') == 1:
            status = 'заявка подписана и передана в мто'
        elif file.count('+') == 2:
            status = 'заявка отработана, есть договор и счет'
        elif file.count('+') == 3:
            status = 'счет подписан и передан в оплату'
        elif file.count('+') == 4:
            status = 'товар поставлен'
        elif file.count('+') == 5:
            status = 'договор закрыт'
        elif '=' in file:
            status = 'заявка готовится к торгам'
        else:
            status = ''
        file_list.append([file.lower(), date.strftime('%Y-%m-%d'), status])
    else:
        print(f'rejected: {file}')

df = pd.DataFrame(file_list, columns=['Заявка', 'Дата', 'Статус'])\
    .sort_values(by='Статус', ascending=False)

with pd.ExcelWriter(f'{dst}/registry.xlsx', engine='xlsxwriter') as wb:
    df.to_excel(wb, sheet_name='Registry', index=False, startrow=1, startcol=1)
    sheet = wb.sheets['Registry']

    sheet.set_column('B:B', 50)
    sheet.set_column('C:C', 10)
    sheet.set_column('D:D', 33)

print('Registry created')

