import re

import pandas as pd

dst = 'U:/Документы ИТ/МТО/Заявки/Реестр'


def reformat(name):
    name = name.lower()
    name = name.replace('"', '')
    name = name.replace('«', '')
    name = name.replace('»', '')
    name = name.replace('-', ' ')
    name = name.replace('ип', '')
    name = name.replace('ооо', '')
    name = name.replace('000', '')
    name = name.strip()
    return name.title()


def list_to_string(src_list):
    el_list = []
    for i in src_list:
        el_list.append(str(i).lower().strip())
    string = ','.join(str(el).lower() for el in list(set(el_list)))
    string = string.replace(',', ', ')
    string = string.replace('<', ' ')
    string = string.replace('>', ' ')
    string = string.replace('  ', ' ')
    string = string.strip()
    return string


def check_email(data):
    email_list = []
    split_data = str(data).split()
    for i in split_data:
        i = i.lower()
        if re.match('[^@]+@[^@]+\.[^@]+', i):
            i = i.replace(',', '')
            i = i.replace('/', '.')
            email_list.append(i)
    email = ', '.join(str(el) for el in list(set(email_list)))
    return email


def check_number(data):
    number_list = []
    split_data = str(data).split()
    for i in split_data:
        i = i.lower()
        if not re.match('[^@]+@[^@]+\.[^@]+', i):
            i = i.replace(',', '')
            i = i.replace('|', '')
            i = i.replace('e-mail:', '')
            i = i.replace('email:', '')
            i = i.replace('mail:', '')
            i = i.replace('скрин', '')
            i = i.replace('шот', '')
            if i not in number_list:
                number_list.append(i)
    number_list = ' '.join(str(el).title() for el in number_list)
    number = number_list.replace('  ', ' ')
    return number


def convert_to_excel(file):
    df_xlsx = pd.read_excel(file)
    df_xlsx = df_xlsx[['Наименование контрагента', 'Контактная информация', 'Перечень товаров, работ, услуг']] \
        .where(df_xlsx['Исполнитель'] == 'it').dropna()
    df_xlsx['Наименование контрагента'] = df_xlsx['Наименование контрагента'].apply(reformat)
    return df_xlsx


kp_2018 = convert_to_excel('kp_2018.xlsx')
kp_2019 = convert_to_excel('kp_2019.xlsx')
kp_2020 = convert_to_excel('kp_2020.xlsx')
kp_2021 = convert_to_excel('kp_2021.xlsx')
kp_2022 = convert_to_excel('kp_2022.xlsx')
request_2020 = convert_to_excel('request_2020.xlsx')
request_2021 = convert_to_excel('request_2021.xlsx')
request_2022 = convert_to_excel('request_2022.xlsx')

kp = kp_2018\
    .merge(kp_2019, how='outer')\
    .merge(kp_2020, how='outer')\
    .merge(kp_2021, how='outer')\
    .merge(kp_2022, how='outer')

request = request_2020\
        .merge(request_2021, how='outer')\
        .merge(request_2022, how='outer')

df = kp.merge(request, how='outer')
df = df.rename(columns={"Наименование контрагента": "agend",
                        "Контактная информация": "contact",
                        "Перечень товаров, работ, услуг": "tasks"})

df_info = df.groupby('agend')['contact'].unique().reset_index()
df_info['contact'] = df_info['contact'].apply(list_to_string)

df_task = df.groupby('agend')['tasks'].unique().reset_index()
df_task['tasks'] = df_task['tasks'].apply(list_to_string)

df_merge = df_info.merge(df_task, how='outer').reset_index(drop=True) \
    .sort_values(by='agend') \
    .reset_index(drop=True)

# df_merge['contact'] = df_merge['contact'].apply(data_filter)
df_merge['e-mail'] = df_merge['contact'].apply(check_email)
df_merge['number'] = df_merge['contact'].apply(check_number)
df_all = df_merge[['agend', 'e-mail', 'number', 'tasks']]

print('File created')

with pd.ExcelWriter(f'{dst}/Контрагенты.xlsx', engine='xlsxwriter') as wb:
    df_all.to_excel(wb, sheet_name='Контрагенты', index=False, startrow=1, startcol=1)
    sheet = wb.sheets['Контрагенты']

    sheet.set_column('B:B', 35)
    sheet.set_column('C:C', 35)
    sheet.set_column('D:D', 70)
    sheet.set_column('E:E', 100)
