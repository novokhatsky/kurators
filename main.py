# -*- coding: utf-8 -*-

from openpyxl import load_workbook, Workbook
import os
from copy import copy
from datetime import date
from datetime import datetime
import shutil
import traceback

#BASE_DIR = "d:\\work\\in\\"
#BASE_OUT = "d:\\work\\out\\"
BASE_DIR = "d:\\tmp\\rubcov\\pro\\in\\"
BASE_OUT = "d:\\tmp\\rubcov\\pro\\out\\"
BACKUP_PATH = "d:\\tmp\\rubcov\\pro\\backup\\"

PPR = BASE_DIR + "пэн-ппр\\ППР.xlsx"
PEN = BASE_DIR + "пэн-ппр\\ПЭН.xlsx"

PPR_OUT = BASE_OUT + "пэн-ппр\\ППР.xlsx"
PEN_OUT = BASE_OUT + "пэн-ппр\\ПЭН.xlsx"

# поля которые переносятся из ПЭН и ППР в файлы кураторов
PPR_INDEX_I = [21, 22, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 74]
PEN_INDEX_I = [21, 22, 35, 36, 37, 38, 39, 40, 41, 42, 43, 44, 45, 46, 47, 48, 49, 74]

# поля которые переносятся из кураторов в ПЭН и ППР
PPR_INDEX_II = [19, 20, 50, 51, 52, 53, 54, 55, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98]
PEN_INDEX_II = [19, 20, 50, 51, 52, 53, 54, 55, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90, 91, 92, 93, 94, 95, 96, 97, 98]


def makeOut(filename):
    el = filename.split("\\")

    return BASE_OUT + el[-1]


def makeDict(filename):
    wb = load_workbook(filename, read_only = True)
    sh = wb.active
    data = {} 

    # нужно пропустить три строки
    enable_add = False
    for row in sh.iter_rows():
    
        if enable_add:
            data[row[0].value] = [cell.value for cell in row]
            continue

        if row[0].value == 'Идентификатор':
            enable_add = True

    wb.close()

    return data


def listFiles(base_dir):
    # формируем список файлов с заданной маской в директории
    spisok_file = []
    for i in os.listdir(base_dir):
        if os.path.isfile(os.path.join(base_dir, i)):
           if i.endswith('.xlsx') or i.endswith('.xlsm'):
                spisok_file.append(os.path.join(base_dir, i))

    return spisok_file


def currDateTime():
    return 'Обновлено: ' + datetime.today().strftime('%d-%m-%Y %H:%M')

print('load ppr')
ppr_dict = makeDict(PPR)

print('load pen')
pen_dict = makeDict(PEN)

# получение списка файлов кураторов
files_kurators = listFiles(BASE_DIR)

for fl in files_kurators:

    print('processing {0}'.format(fl))

    if fl.endswith('.xlsm'):
        wb = load_workbook(fl, read_only = False, keep_vba = True)
    else:
        wb = load_workbook(fl)

    sh = wb.active
    data = {} 

    # нужно пропустить три строки
    enable_work = False
    for row in sh.iter_rows():

        if enable_work:
            key = row[0].value

            if key in ppr_dict:
                # есть идентификатор в ппр

                for index in PPR_INDEX_I:

                    if row[index - 1].value != ppr_dict[key][index - 1]:
                        row[index - 1].value = ppr_dict[key][index - 1]
            else:
                pass
                #out_log("идентификатор {0} не найден в ППР".format(key))

            if key in pen_dict:
                # есть идентификатор в ппр

                for index in PEN_INDEX_I:

                    if row[index - 1].value != pen_dict[key][index - 1]:
                        row[index - 1].value = pen_dict[key][index - 1]
            else:
                pass
                #out_log("идентификатор {0} не найден в ПЭН".format(key))

            continue

        if row[0].value == 'Идентификатор':
            enable_work = True

    out_filename = makeOut(fl)
    print("save {0}".format(out_filename))

    sh['A1'] = currDateTime()

    wb.save(out_filename)
    wb.close()
    wb = None

pen_dict.clear()
ppr_dict.clear()

print('load kurators')
dicts = []
for fl in listFiles(BASE_OUT):
    dicts.append(makeDict(fl))

print('load ppr')
wb = load_workbook(PPR)
sh = wb.active

print('seek ppr')

for row in sh.iter_rows():
    key = row[0].value

    for dic in dicts:

        if key in dic.keys():
            
            for index in PPR_INDEX_II:
                row[index - 1].value = dic[key][index - 1]

print('save ppr')
wb.save(PPR_OUT)
wb.close()
wb = None

print('load pen')
wb = load_workbook(PEN, read_only = True)
sh = wb.active

theOne = Workbook(write_only = True)
o = theOne.create_sheet('ПЭН_ППР сокращ')
newSheet = theOne[o.title]

print('seek pen')

for row in sh.iter_rows():
    key = row[0].value
    new_row = [] 

    for dic in dicts:

        if key in dic.keys():

            new_row = [cell.value for cell in row]

            for index in PEN_INDEX_II:
                new_row[index - 1] = dic[key][index - 1]

            break

    if new_row == []:
        newSheet.append([cell.value for cell in row])
    else:
        newSheet.append([cell for cell in new_row])


print('clear')
dicts = None

wb.close()
wb = None

print('save pen')
theOne.save(PEN_OUT)
theOne.close()

# создание резервной копии с датой
fullFileBackup = BACKUP_PATH + date.today().isoformat()

if os.path.isdir(fullFileBackup):
    print("exists")
    exit()

# создаем каталог и переносим все вайлы из in
shutil.copytree(BASE_DIR, fullFileBackup)

# удаляем каталог ИН
shutil.rmtree(BASE_DIR)

# создаем каталог ИН и копируем туда данные из АУТ
shutil.copytree(BASE_OUT, BASE_DIR)

# удаляем каталог АУТ 
shutil.rmtree(BASE_OUT)

# создаем АУТ и пэн-ппр
os.mkdir(BASE_OUT)
os.mkdir(BASE_OUT + "пэн-ппр")

