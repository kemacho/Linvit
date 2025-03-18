import os
import glob
from tqdm import tqdm

a = 302
b = 397

for i in tqdm(range(a, b + 1), desc = 'Идет проверка папок'):
    name = str(i) + ' - РЖД'

    list0 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\0*')
    list1 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\1*')
    list2 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\2*')
    list3 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\3*')
    list4 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\4*')
    list5 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\5*')
    list6 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\6*')
    list7 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\7*')
    list8 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\8*')
    list9 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\9*')
    list10 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\10*')
    list11 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\11*')
    list12 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\12*')

    listF = [list0[0], list1[0], list2[0], list3[0], list4[0], list5[0], list6[0], list7[0], list8[0], list9[0], list10[0], list11[0], list12[0]]

    #5
    try:
        contents = os.listdir(listF[5])
        if len(contents) > 0:
            os.rename(listF[5], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\5 Протоколы СИ (+)')
        elif len(contents) == 0:
            os.rename(listF[5], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\5 Протоколы СИ (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[5]}' не существует.")

    #0
    try:
        contents = os.listdir(listF[0])
        if len(contents) > 1:
            os.rename(listF[0], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\0 Заявка и приложение (+)')
        elif len(contents) > 0:
            os.rename(listF[0],rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\0 Заявка и приложение (+—)')
        elif len(contents) == 0:
            os.rename(listF[0], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\0 Заявка и приложение (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[0]}' не существует.")

    #1
    try:
        contents = os.listdir(listF[1])
        if len(contents) > 0:
            os.rename(listF[1], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\1 Распоряжение по заявке (+)')
        elif len(contents) == 0:
            os.rename(listF[1], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\1 Распоряжение по заявке (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[1]}' не существует.")

    #2
    try:
        contents = os.listdir(listF[2])
        if len(contents) > 0:
            os.rename(listF[2], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\2 Решение по заявке (+)')
        elif len(contents) == 0:
            os.rename(listF[2], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\2 Решение по заявке (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[2]}' не существует.")

    #3
    try:
        contents = os.listdir(listF[3])
        if len(contents) > 3:
            os.rename(listF[3], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\3 Заключения по ОМД и ТД (+)')
        elif len(contents) == 3:
            os.rename(listF[3], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\3 Заключения по ОМД и ТД (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[3]}' не существует.")

    #4
    try:
        contents = os.listdir(listF[4])
        if len(contents) > 0:
            os.rename(listF[4], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\4 Акт выбора ПК (+)')
        elif len(contents) == 0:
            os.rename(listF[4], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\4 Акт выбора ПК (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[4]}' не существует.")

    #5
    try:
        contents = os.listdir(listF[5])
        if len(contents) > 0:
            os.rename(listF[5], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\5 Протоколы СИ (+)')
        elif len(contents) == 0:
            os.rename(listF[5], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\5 Протоколы СИ (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[5]}' не существует.")

    #6
    try:
        contents = os.listdir(listF[6])
        if len(contents) > 0:
            os.rename(listF[6], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\6 Заключение СИ (+)')
        elif len(contents) == 0:
            os.rename(listF[6], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\6 Заключение СИ (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[6]}' не существует.")

    #7
    try:
        contents = os.listdir(listF[7])
        if len(contents) > 0:
            os.rename(listF[7], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\7 Программа проверки произв (+)')
        elif len(contents) == 0:
            os.rename(listF[7], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\7 Программа проверки произв (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[7]}' не существует.")

    #8
    try:
        contents = os.listdir(listF[8])
        if len(contents) > 0:
            os.rename(listF[8], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\8 Акт ПП (+)')
        elif len(contents) == 0:
            os.rename(listF[8], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\8 Акт ПП (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[8]}' не существует.")

    #9
    try:
        contents = os.listdir(listF[9])
        if len(contents) > 0:
            os.rename(listF[9], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\9 Распоряжение на анализ (+)')
        elif len(contents) == 0:
            os.rename(listF[9], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\9 Распоряжение на анализ (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[9]}' не существует.")

    #10
    try:
        contents = os.listdir(listF[10])
        if len(contents) > 1:
            os.rename(listF[10], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\10 Решение о выдаче (+)')
        elif len(contents) > 0:
            os.rename(listF[10],rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\10 Решение о выдаче (+—)')
        elif len(contents) == 0:
            os.rename(listF[10], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\10 Решение о выдаче (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[10]}' не существует.")

    #11
    try:
        contents = os.listdir(listF[11])
        if len(contents) > 1:
            os.rename(listF[11], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\11 Сертификат (+)')
        elif len(contents) > 0:
            os.rename(listF[11],rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\11 Сертификат (+—)')
        elif len(contents) == 0:
            os.rename(listF[11], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\11 Сертификат (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[11]}' не существует.")

    #12
    try:
        contents = os.listdir(listF[12])
        if len(contents) > 0:
            os.rename(listF[12], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\12 Доп.материалы (+)')
        elif len(contents) == 0:
            os.rename(listF[12], rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{name}\12 Доп.материалы (—)')
    except FileNotFoundError:
        print(f"Папка '{listF[12]}' не существует.")

print('Successful!')