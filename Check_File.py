import os
import glob
from tqdm import tqdm

# ПУТЬ К КОРНЕВОЙ ПАПКЕ
INpath = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты РЖД тест'

# Названия папок
folder_names = os.listdir(INpath)

worry = [] # непонятки

FileQNT = [2, 1, 1, 4, 1, 'Any', 1, 1, 1, 1, 2, 2, 'Any']
FileQNT_IK = [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1]

def worrymessage(pt, nm1, nm2, contents):
    worry.append(pt)
    worry.append(nm1)
    worry.append(nm2)
    for i in range(len(contents)):
        worry.append(contents[i])

def process_folder_with_one_file(old_folder, contents, Pos_dest, Neg_dest):
    """Обрабатывает папки, ожидающие 1 файл."""
    if len(contents) == 1 and old_folder != Pos_dest:
        os.rename(old_folder, Pos_dest)
    elif len(contents) == 0 and old_folder != Neg_dest:
        os.rename(old_folder, Neg_dest)
    elif len(contents) > 1:
        os.rename(old_folder, Pos_dest)
        worrymessage(old_folder, len(contents), '1', contents)

def process_folder_with_four_files(old_folder, contents, Pos_dest, Neg_dest):
    """Обрабатывает папки, ожидающие 4 файла."""
    if len(contents) == 4 and old_folder != Pos_dest:
        os.rename(old_folder, Pos_dest)
    elif len(contents) < 4 and old_folder != Neg_dest:
        os.rename(old_folder, Neg_dest)
    elif len(contents) > 4:
        os.rename(old_folder, Pos_dest)
        worrymessage(old_folder, len(contents), '4', contents)

def process_folder_with_any_files(old_folder, contents, Pos_dest, Neg_dest):
    """Обрабатывает папки, где ожидается любое количество файлов (больше 1)."""
    if len(contents) > 0 and old_folder != Pos_dest:
        os.rename(old_folder, Pos_dest)
    elif len(contents) == 0 and old_folder != Neg_dest:
        os.rename(old_folder, Neg_dest)
    elif len(contents) > 20:
        os.rename(old_folder, Pos_dest)
        worrymessage(old_folder, len(contents), 'возможно не так много', contents)

def process_folder_with_two_files(old_folder, contents, Pos_dest, Norm_dest, Neg_dest):
    """Обрабатывает папки, ожидающие 2 файла."""
    if len(contents) == 2 and old_folder != Pos_dest:
        os.rename(old_folder, Pos_dest)
    elif len(contents) == 1 and old_folder != Norm_dest:
        os.rename(old_folder, Norm_dest)
    elif len(contents) == 0 and old_folder != Neg_dest:
        os.rename(old_folder, Neg_dest)
    elif len(contents) > 2:
        os.rename(old_folder, Pos_dest)
        worrymessage(old_folder, len(contents), '2', contents)

def check(FileQNT, old_folder, contents, Pos_dest, Norm_dest, Neg_dest):
    if FileQNT == 1:
        process_folder_with_one_file(old_folder, contents, Pos_dest, Neg_dest)
    if FileQNT == 2:
        process_folder_with_two_files(old_folder, contents, Pos_dest, Norm_dest, Neg_dest)
    elif FileQNT == 4:
        process_folder_with_four_files(old_folder, contents, Pos_dest, Neg_dest)
    elif FileQNT == 'Any':
        process_folder_with_any_files(old_folder, contents, Pos_dest, Neg_dest)

# Названия подпапок СИ
nameSI = ['0 Заявка и приложение', '1 Распоряжение по заявке', '2 Решение по заявке', '3 Заключения по ОМД и ТД',
          '4 Акт выбора ПК', '5 Протоколы СИ', '6 Заключение СИ', '7 Программа проверки произ', '8 Акт ПП',
          '9 Распоряжение на анализ', '10 Решение о выдаче', '11 Сертификат', '12 Доп.материалы']

# Названия подпапок ИК
nameIK = ['0 Распоряжение', '1 Письмо-уведомление', '2 Программа ИК', '3 Программа проверки произ', '4 Акт выбора ПК',
          '5 Акт проверки производства', '6 Протоколы ИК', '7 Акт по результатам ИК', '8 Распоряжение на анализ', 
          '9 Решение по ИК', '10 Доп. материалы']

for i in tqdm(range(0, len(folder_names)), desc='Идет проверка папок'):

    # Путь к подпапкам первого уровня
    name = folder_names[i]

    # Путь к подпапкам второго уровня
    pathSI = rf'{INpath}\{name}\0. СИ'
    pathIK1 = rf'{INpath}\{name}\1. ИК-1'
    pathIK2 = rf'{INpath}\{name}\2. ИК-2'

    # Проверка для папки СИ
    for j in range(0, len(FileQNT)):
        
        nameSIzv = nameSI[j] + '*'
        old_folder = glob.glob(os.path.join(pathSI, nameSIzv))
        old_folder = old_folder[0]
        contents = os.listdir(old_folder)

        folder = os.path.join(pathSI, nameSI[j])
        Pos_dest = str(folder) + ' (+)'
        Norm_dest = str(folder) + ' (+—)'
        Neg_dest = str(folder) + ' (—)'

        check(FileQNT[j], old_folder, contents, Pos_dest, Norm_dest, Neg_dest)

    # Проверка для папки ИК1
    for j in range(0, len(FileQNT_IK)):

        nameIK1zv = nameIK[j] + '*'
        old_folder = glob.glob(os.path.join(pathIK1, nameIK1zv))
        old_folder = old_folder[0]
        contents = os.listdir(old_folder)


        folder = os.path.join(pathIK1, nameIK[j])
        Pos_dest = str(folder) + ' (+)'
        Norm_dest = str(folder) + ' (+—)'
        Neg_dest = str(folder) + ' (—)'

        check(FileQNT[j], old_folder, contents, Pos_dest, Norm_dest, Neg_dest)

    # Проверка для папки ИК2
    for j in range(0, len(FileQNT_IK)):
        nameIK1zv = nameIK[j] + '*'
        old_folder = glob.glob(os.path.join(pathIK2, nameIK1zv))
        old_folder = old_folder[0]
        contents = os.listdir(old_folder)

        folder = os.path.join(pathIK2, nameIK[j])
        Pos_dest = str(folder) + ' (+)'
        Norm_dest = str(folder) + ' (+—)'
        Neg_dest = str(folder) + ' (—)'

        check(FileQNT[j], old_folder, contents, Pos_dest, Norm_dest, Neg_dest)
    
print('Проверка завершена!')
for i in range(0, len(worry), 3):
    print('Пожалуйста проверьте папку:', worry[i], 'там находится ', worry[i + 1], 'файла, вместо', worry[i + 2], ' (', worry[i + 3], ')')