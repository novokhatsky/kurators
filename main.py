# -*- coding: utf-8 -*-

from openpyxl import load_workbook
import os
import logging

BASE_DIR = "d:\\tmp\\rubcov\\большие\\"
PREFIX_OUT = "out_"

PPR = BASE_DIR + "пэн-ппр\\ППР.xlsx"
PEN = BASE_DIR + "пэн-ппр\\ПЭН.xlsx"

PPR_INDEX = [38, 39, 40, 41, 50, 51, 52, 53, 54, 55, 91, 92, 93, 94, 95, 96, 97, 98]
PEN_INDEX = [35, 36, 37, 42, 43, 44, 45, 46, 47, 48, 49, 74, 75, 76, 77, 78, 79, 80, 81, 82, 83, 84, 85, 86, 87, 88, 89, 90]

LOG_FILE = "process.log"

logging.basicConfig(format = u'[%(asctime)s] %(message)s', level = logging.INFO, filename = LOG_FILE)


def out_log(mess):
    logging.info(mess)


def makeOut(filename):
    el = filename.split("\\")
    el[-1] = PREFIX_OUT + el[-1]

    return "\\".join(el)

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
out_log("загрузка ППР")

ppr_dict = makeDict(PPR)

print('load pen')
out_log("загрузка ПЭН")
pen_dict = makeDict(PEN)

# получение списка файлов кураторов
files_kurators = filesKurators(BASE_DIR)

for fl in files_kurators:

    #fl = files_kurators[1]

    print('processing {0}'.format(fl))
    out_log("обработка {0}".format(fl))

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
                    out_log("в поле {0} вставлено значение {1}".format(index, ppr_dict[key][index - 1]))

                    row[index - 1].value = ppr_dict[key][index - 1]

            if key in pen_dict:
                # есть идентификатор в ппр
                out_log("найден идентификатор {0} в ПЭН".format(key))

                for index in PEN_INDEX:
                    out_log("в поле {0} вставлено значение {1}".format(index, pen_dict[key][index - 1]))

                    row[index - 1].value = pen_dict[key][index - 1]

    print("save {0}".format(makeOut(fl)))
    out_log("запись {0}".format(makeOut(fl)))
    wb.save(makeOut(fl))

# ппр 38-41 50-55 91-98 AL-AM-AN-AO AX-AY-AZ-BA-BB-DC CM-CN-CO-CP-CQ-CR-CS-CT
# пэн 35-37 42-49 74-90

