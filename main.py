# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import sqlite3
import json
import csv
import codecs

BASE_DIR = "d:\\tmp\\rubcov\\большие\\"

PPR = BASE_DIR + "пэн-ппр\\ППР.xlsx"

FROM_PPR = ['AL', 'AM', 'AN', 'AO', 'AX', 'AY', 'AZ', 'BA', 'BB', 'DC', 'CM', 'CN', 'CO', 'CP', 'CQ', 'CR', 'CS', 'CT']

db = sqlite3.connect('xml.db')

cursor = db.cursor()

cursor.execute('''
    select count(name) from sqlite_master where type = 'table' and name = 'ppr'
''')

if cursor.fetchone()[0] == 1:
    # table exists
    pass
else:
    # table not exists
    cursor.execute('''
        create table ppr (
            id varchar (100) primary key,
            data text)
    ''')



def dictForSave(i):
    temp = {}
    for index in FROM_PPR:
        temp[index] = ppr_sheet['{0}{1}'.format(index, i)].value

    return temp

ppr_wb = load_workbook(PPR, read_only = True)

sh = ppr_wb.active
d = {} 
for row in sh.iter_rows():
    print(row[0].value)
    d[row[0].value] = [cell.value for cell in row]

exit()

with codecs.open('test.csv', 'w', 'utf-16') as file_handle:
    csv_writer = csv.writer(file_handle, dialect='excel', delimiter=';')
    for row in sh.iter_rows(): # generator; was sh.rows
        csv_writer.writerow([cell.value for cell in row])

exit()

ppr_sheetnames = ppr_wb.sheetnames

ppr_sheet = ppr_wb[ppr_sheetnames[0]]

ppr_dict = {}
i = 3

while (True):
    try:
        i += 1
        
        address = "A{0}".format(i)
        
        print('{0} - {1}'.format(i, ppr_sheet[address].value))

        #cursor.execute('''
        #    insert into ppr (id, data) values (?, ?)
        #    ''',
        #    [ppr_sheet[address].value, json.dumps(dictForSave(i))]
        #)

    except ValueError:
        print("exception")
        break;
print(i)
# ппр 38-41 50-55 91-98 AL-AM-AN-AO AX-AY-AZ-BA-BB-DC CM-CN-CO-CP-CQ-CR-CS-CT
# пэн 35-37 42-49 74-90

db.commit()

db.close()
