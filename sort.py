import os
from datetime import datetime

src = 'U:/Документы ИТ/МТО/Заявки'
arc = 'U:/Документы ИТ/МТО/Заявки/Архив'

for file in os.listdir(src):
    time = os.path.getctime(f'{src}/{file}')
    date = datetime.fromtimestamp(time)

    if not os.path.isdir(f'{arc}/{date.year}'):
        os.mkdir(f'{arc}/{date.year}')

    if '-' in file and '~$' not in file:
        os.replace(f'{src}/{file}', f'{arc}/{date.year}/{date.strftime("%Y.%m.%d")} - {file.split("-")[-1].strip()}')
