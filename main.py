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


class fileExcel(object):
    def __init__(self, filename):
        self.filename = filename
        self.workbook = None
        self.sheet = None


    def create(self, sheetname):
        self.workbook = Workbook(write_only = True)
        self.sheet = self.workbook[self.workbook.create_sheet(sheetname).title]


    def append(self, row):
        self.sheet.append(row)


    def save(self):
        self.workbook.save(self.filename)
        self.workbook.close()


class kuratorsDict(object):
    def __init__(self):
        self.dicts = []


    def load(self):
        for fl in listFiles(BASE_OUT):
            self.dicts.append(makeDict(fl))


    def seekkey(self, key):
        for dic in self.dicts:

            if key in dic.keys():

                return dic[key]

        return False


print('load ppr')
ppr_dict = makeDict(PPR)

print('load pen')
pen_dict = makeDict(PEN)

# получение списка файлов кураторов
files_kurators = listFiles(BASE_DIR)

# создаем книгу для записи ненайденных ИД
notFoundId = fileExcel(BASE_OUT + 'no_id_kur_in_ppr_pen.xlsx')
notFoundId.create('not found')

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
                notFoundId.append([key, 'ППР', fl.replace(BASE_DIR, '')])

            if key in pen_dict:
                # есть идентификатор в ппр

                for index in PEN_INDEX_I:

                    if row[index - 1].value != pen_dict[key][index - 1]:
                        row[index - 1].value = pen_dict[key][index - 1]
            else:
                notFoundId.append([key, 'ПЭН', fl.replace(BASE_DIR, '')])

            continue

        if row[0].value == 'Идентификатор':
            enable_work = True

    out_filename = makeOut(fl)
    print("save {0}".format(out_filename))

    sh['Y1'] = currDateTime()

    wb.save(out_filename)

notFoundId.save()
del notFoundId
del pen_dict
del ppr_dict


print('load kurators')
kurators_dict = kuratorsDict()
kurators_dict.load()


print('load ppr')
wb = load_workbook(PPR)
sh = wb.active

# создаем книгу для записи ненайденных ИД
notFoundId = fileExcel(BASE_OUT + 'no_id_ppr_in_kur.xlsx')
notFoundId.create('not found')

print('seek ppr')

for row in sh.iter_rows():
    key = row[0].value
    
    found = kurators_dict.seekkey(key)

    if found:
        for index in PPR_INDEX_II:
            row[index - 1].value = found[index - 1]

    else:
        notFoundId.append([key])

notFoundId.save()
del notFoundId

print('save ppr')
wb.save(PPR_OUT)
del sh
del wb

print('load pen')
wb = load_workbook(PEN, read_only = True)
sh = wb.active

# создаем книгу для записи ненайденных ИД
notFoundId = fileExcel(BASE_OUT + 'no_id_pen_in_kur.xlsx')
notFoundId.create('not found')

# создаем книгу для создания нового файла ПЭН
newPen = fileExcel(PEN_OUT)
newPen.create('ПЭН_ППР сокращ')

print('seek pen')

for row in sh.iter_rows():
    key = row[0].value

    found = kurators_dict.seekkey(key)

    if found:
        new_row = [cell.value for cell in row]

        for index in PEN_INDEX_II:
            new_row[index - 1] = found[index - 1]

        newPen.append(new_row)

    else:
        newPen.append([cell.value for cell in row])
        notFoundId.append([key])

print('clear')
notFoundId.save()
del notFoundId
del sh
del wb

print('save pen')
newPen.save()
del newPen

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

