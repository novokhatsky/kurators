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
           if i.endswith('.xlsx'):
                spisok_file.append(os.path.join(base_dir, i))

    return spisok_file


def makeFileKurator(fio):
    fam, nam, sec = fio.split(' ')

    return BASE_OUT + fam + '.xlsx'

    
def makeDictKurator(filename):
    wb = load_workbook(filename, read_only = True)
    sh = wb.active
    data = {} 

    # нужно пропустить три строки
    enable_add = False
    for row in sh.iter_rows():

        if row[0].value == 'Идентификатор':
            # после строки Идентификатор можно формировать словарь
            enable_add = True
            continue
    
        if enable_add:
            kurator = row[17].value

            if kurator not in data:
                data[kurator] = []

            data[kurator].append([cell.value for cell in row])

    return data


print('load ppr')
out_log("загрузка ППР")

# получили массив из ППР, где ключом является куратор
ppr_dict = makeDictKurator(PPR)

for key in ppr_dict.keys():
    print(key)
