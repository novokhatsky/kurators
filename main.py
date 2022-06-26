# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import os

BASE_DIR = "d:\\tmp\\rubcov\\большие\\"

PPR = BASE_DIR + "пэн-ппр\\ППР.xlsx"
PEN = BASE_DIR + "пэн-ппр\\ПЭН.xlsx"

PPR_INDEX = [38, 39, 40, 41, 50, 51, 52, 53, 54, 55, 91, 92, 93, 94, 95, 96, 97, 98]
PEN_INDEX = [35, 36, 37, 42, 43, 44, 45, 46, 47, 48, 49, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90]


def dictForSave(i):
    temp = {}
    for index in FROM_PPR:
        temp[index] = ppr_sheet['{0}{1}'.format(index, i)].value

    return temp


def makeDict(filename):
    wb = load_workbook(filename, read_only = True)
    sh = wb.active
    data = {} 

    # нужно пропустить три строки
    enable_add = False
    for row in sh.iter_rows():

        if row[0].value == 'Идентификатор':
            enable_add = True
            continue
    
        if enable_add:
            data[row[0].value] = [cell.value for cell in row]

    return data


def filesKurators(base_dir):
    # формируем список файлов с заданной маской в директории
    spisok_file = []
    for i in os.listdir(base_dir):
        if os.path.isfile(os.path.join(base_dir, i)):
           if i.endswith('.xlsx'):
                spisok_file.append(os.path.join(base_dir, i))

    return spisok_file


print('load ppr')
ppr_dict = makeDict(PPR)

print('load pen')
pen_dict = makeDict(PEN)

# получение списка файлов кураторов
files_kurators = filesKurators(BASE_DIR)

for fl in files_kurators:

    #fl = files_kurators[1]

    print('processing {0}'.format(fl))

    wb = load_workbook(fl)
    sh = wb.active
    data = {} 

    # нужно пропустить три строки
    enable_add = False
    for row in sh.iter_rows():

        if row[0].value == 'Идентификатор':
            enable_add = True
            continue

        if enable_add:
            key = row[0].value

            if key in ppr_dict:
                # есть идентификатор в ппр
                for index in PPR_INDEX:
                    row[index-1].value = ppr_dict[key][index - 1]

            if key in pen_dict:
                # есть идентификатор в ппр
                for index in PEN_INDEX:
                    row[index-1].value = pen_dict[key][index - 1]

    wb.save(fl)
# ппр 38-41 50-55 91-98 AL-AM-AN-AO AX-AY-AZ-BA-BB-DC CM-CN-CO-CP-CQ-CR-CS-CT
# пэн 35-37 42-49 74-90

