from ftplib import FTP

ftp = FTP('ftp.dagomys.ru')
ftp.login('ovchinnikovse', 'Zxc123zxc')
ftp.encoding = 'utf-8'
ftp.sendcmd('OPTS UTF8 ON')

# data = ftp.nlst()
# for i in data:
#     if '~' not in i:
#         print(i)

filename = 'Телефонный справочник.xlsx'
ftp.retrbinary(f'RETR {filename}', open('phonebook.xlsx', 'wb').write)

ftp.quit()
