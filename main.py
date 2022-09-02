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
        file_list.append([file.lower(), date.strftime('%Y-%m-%d'), 'проверка...'])
    else:
        print(f'rejected: {file}')

df = pd.DataFrame(file_list, columns=['Наименование', 'Дата', 'Статус'])\
    .sort_values(by='Дата', ascending=False)

with pd.ExcelWriter(f'{dst}/registry.xlsx', engine='xlsxwriter') as wb:
    df.to_excel(wb, sheet_name='Registry', index=False, startrow=1, startcol=1)
    sheet = wb.sheets['Registry']

    sheet.set_column('B:B', 50)
    sheet.set_column('C:C', 10)
    sheet.set_column('D:D', 11)

print('Registry created')

