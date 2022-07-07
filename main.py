# -*- coding: utf-8 -*-

from openpyxl import load_workbook, Workbook
import os
import logging
from copy import copy

import traceback

#BASE_DIR = "d:\\work\\in\\"
#BASE_OUT = "d:\\work\\out\\"
BASE_DIR = "d:\\tmp\\rubcov\\pro\\in\\"
BASE_OUT = "d:\\tmp\\rubcov\\pro\\out\\"

ALL = BASE_OUT + 'out_all.xlsx'

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

LOG_FILE = BASE_DIR + "process.log"

logging.basicConfig(format = u'[%(asctime)s] %(message)s', level = logging.INFO, filename = LOG_FILE)


def out_log(mess):
    logging.info(mess)


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


def saveAll():
    filesInput = listFiles(BASE_OUT)

    theOne = Workbook(write_only = True)
    o = theOne.create_sheet('ПЭН_ППР сокращ')
    newSheet = theOne[o.title]

    # для первого файла считаем все строки
    enable_add = True
    for oneFile in filesInput:
        print("load {0}".format(oneFile))
        wb = load_workbook(oneFile, read_only = True)
        sourceSheet = wb.active

        for row in sourceSheet.iter_rows():
        
            if enable_add:
                newSheet.append([cell.value for cell in row])
                continue

            if row[0].value == 'Идентификатор':
                enable_add = True

        # у второго и последующих файлов, пропускаем строки включительно до Иденификатора
        enable_add = False
        wb.close()

    print('save all')
    theOne.save(ALL)
    theOne.close()


print('load ppr')
out_log("загрузка ППР")
ppr_dict = makeDict(PPR)

print('load pen')
out_log("загрузка ПЭН")
pen_dict = makeDict(PEN)

# получение списка файлов кураторов
files_kurators = listFiles(BASE_DIR)

for fl in files_kurators:

    print('processing {0}'.format(fl))
    out_log("обработка {0}".format(fl))

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
                out_log("найден идентификатор {0} в ППР".format(key))

                for index in PPR_INDEX_I:

                    if row[index - 1].value != ppr_dict[key][index - 1]:
                        out_log("в поле {0} замена {1} на {2}".format(index, row[index - 1].value, ppr_dict[key][index - 1]))
                        row[index - 1].value = ppr_dict[key][index - 1]
            else:
                out_log("идентификатор {0} не найден в ППР".format(key))

            if key in pen_dict:
                # есть идентификатор в ппр
                out_log("найден идентификатор {0} в ПЭН".format(key))

                for index in PEN_INDEX_I:

                    if row[index - 1].value != pen_dict[key][index - 1]:
                        out_log("в поле {0} замена {1} на {2}".format(index, row[index - 1].value, pen_dict[key][index - 1]))
                        row[index - 1].value = pen_dict[key][index - 1]
            else:
                out_log("идентификатор {0} не найден в ПЭН".format(key))

            continue

        if row[0].value == 'Идентификатор':
            enable_work = True

    out_filename = makeOut(fl)
    print("save {0}".format(out_filename))
    out_log("запись {0}".format(out_filename))
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

