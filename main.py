# -*- coding: utf-8 -*-

from openpyxl import load_workbook, Workbook
import os
import logging
from copy import copy

#BASE_DIR = "d:\\work\\in\\"
#BASE_OUT = "d:\\work\\out\\"
BASE_DIR = "d:\\tmp\\rubcov\\pro\\in\\"
BASE_OUT = "d:\\tmp\\rubcov\\pro\\out\\"

PPR = BASE_DIR + "пэн-ппр\\ППР.xlsx"
PEN = BASE_DIR + "пэн-ппр\\ПЭН.xlsx"

PPR_INDEX = [38, 39, 40, 41, 50, 51, 52, 53, 54, 55, 91, 92, 93, 94, 95, 96, 97, 98]
PEN_INDEX = [35, 36, 37, 42, 43, 44, 45, 46, 47, 48, 49, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90]

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

        if row[0].value == 'Идентификатор':
            enable_add = True
            continue
    
        if enable_add:
            data[row[0].value] = [cell.value for cell in row]

    return data


def listFiles(base_dir):
    # формируем список файлов с заданной маской в директории
    spisok_file = []
    for i in os.listdir(base_dir):
        if os.path.isfile(os.path.join(base_dir, i)):
           if i.endswith('.xlsx') or i.endswith('.xlsm'):
                spisok_file.append(os.path.join(base_dir, i))

    return spisok_file


print('load ppr')
out_log("загрузка ППР")
ppr_dict = makeDict(PPR)

print('load pen')
out_log("загрузка ПЭН")
pen_dict = makeDict(PEN)

# получение списка файлов кураторов
files_kurators = listFiles(BASE_DIR)

for fl in files_kurators:

    #fl = files_kurators[1]

    print('processing {0}'.format(fl))
    out_log("обработка {0}".format(fl))

    if fl.endswith('.xlsm'):
        wb = load_workbook(fl, read_only = False, keep_vba = True)
    else:
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
                out_log("найден идентификатор {0} в ППР".format(key))

                for index in PPR_INDEX:

                    if row[index - 1].value != ppr_dict[key][index - 1]:
                        out_log("в поле {0} замена {1} на {2}".format(index, row[index - 1].value, ppr_dict[key][index - 1]))
                        row[index - 1].value = ppr_dict[key][index - 1]
            else:
                out_log("идентификатор {0} не найден в ППР".format(key))

            if key in pen_dict:
                # есть идентификатор в ппр
                out_log("найден идентификатор {0} в ПЭН".format(key))

                for index in PEN_INDEX:

                    if row[index - 1].value != pen_dict[key][index - 1]:
                        out_log("в поле {0} замена {1} на {2}".format(index, row[index - 1].value, pen_dict[key][index - 1]))
                        row[index - 1].value = pen_dict[key][index - 1]
            else:
                out_log("идентификатор {0} не найден в ПЭН".format(key))

    out_filename = makeOut(fl)
    print("save {0}".format(out_filename))
    out_log("запись {0}".format(out_filename))
    wb.save(out_filename)

filesInput = listFiles(BASE_OUT)
theOneFile = BASE_OUT + 'out_all.xlsx'

theOne = Workbook(write_only = True)
o = theOne.create_sheet('ПЭН_ППР сокращ')
safeTitle = o.title

newSheet = theOne[safeTitle]
row_index = 1
keep3row = True

for oneFile in filesInput:
    print("load {0}".format(oneFile))
    wb = load_workbook(oneFile)
    sourceSheet = wb['ПЭН_ППР сокращ']

    i = 4 if keep3row else 0

    for row in sourceSheet.rows:
        i += 1
        
        if i < 4:
            continue
        newSheet.append(row)
        row_index += 1

    keep3row = False

theOne.save(theOneFile)
# ппр 38-41 50-55 91-98 AL-AM-AN-AO AX-AY-AZ-BA-BB-DC CM-CN-CO-CP-CQ-CR-CS-CT
# пэн 35-37 42-49 74-90

