# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import os

BASE_DIR = "d:\\tmp\\rubcov\\большие\\"

PPR = BASE_DIR + "пэн-ппр\\ППР.xlsx"
PEN = BASE_DIR + "пэн-ппр\\ПЭН.xlsx"


def dictForSave(i):
    temp = {}
    for index in FROM_PPR:
        temp[index] = ppr_sheet['{0}{1}'.format(index, i)].value

    return temp

def makeDict(filename):
    wb = load_workbook(filename, read_only = True)
    sh = wb.active
    data = {} 

    for row in sh.iter_rows():
        data[row[0].value] = [cell.value for cell in row]

    return data
def filesKurators(base_dir):
    # формируем список файлов с заданной маской в директории
    spisok_file = []
    for i in os.listdir(base_dir):
        if os.path.isfile(os.path.join(base_dir, i)):
           if i.endswith('.xlsx'):
                spisok_file.append(i)

    return spisok_file

print('load ppr')
#ppr_dict = makeDict(PPR)

print('load pen')
#pen_dict = makeDict(PEN)

# получение списка файлов кураторов
files_kurators = filesKurators(BASE_DIR)

for fl in files_kurators:
    print(fl)

# ппр 38-41 50-55 91-98 AL-AM-AN-AO AX-AY-AZ-BA-BB-DC CM-CN-CO-CP-CQ-CR-CS-CT
# пэн 35-37 42-49 74-90

