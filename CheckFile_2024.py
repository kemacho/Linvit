import os
import glob
from tqdm import tqdm

# Номера сертификатов
Num_SERT =  [265, 266, 267, 268, 269, 270, 271, 272, 273, 274, 275, 276, 277, 278, 279,
    280, 281, 282, 283, 284, 285, 286, 287, 288, 289, 290, 291, 292, 293, 294,
    295, 296, 297, 298, 299, 300, 301, 353, 354, 355, 357]

worry = [] # непонятки

def worrymessage(pt, nm1, nm2):
    worry.append(pt)
    worry.append(nm1)
    worry.append(nm2)

for i in tqdm(range(0, len(Num_SERT)), desc='Идет проверка папок'):
    name = str(Num_SERT[i])
    name2 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ')
    name3 = name2[0]

    list0 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\0*')
    list1 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\1*')
    list2 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\2*')
    list3 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\3*')
    list4 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\4*')
    list5 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\5*')
    list6 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\6*')
    list7 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\7*')
    list8 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\8*')
    list9 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\9*')
    list10 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\10*')
    list11 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\11*')
    list12 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}*\СИ\12*')


    dest0 = [rf'{name3}\0 Заявка и приложение (+)',
             rf'{name3}\0 Заявка и приложение (+—)',
             rf'{name3}\0 Заявка и приложение (—)']
    dest1 = [rf'{name3}\1 Распоряжение по заявке (+)',
             rf'{name3}\1 Распоряжение по заявке (—)']
    dest2 = [rf'{name3}\2 Решение по заявке (+)',
             rf'{name3}\2 Решение по заявке (—)']
    dest3 = [rf'{name3}\3 Заключения по ОМД и ТД (+)',
             rf'{name3}\3 Заключения по ОМД и ТД (—)']
    dest4 = [rf'{name3}\4 Акт выбора ПК (+)',
             rf'{name3}\4 Акт выбора ПК (—)']
    dest5 = [rf'{name3}\5 Протоколы СИ (+)',
             rf'{name3}\5 Протоколы СИ (—)']
    dest6 = [rf'{name3}\6 Заключение СИ (+)',
             rf'{name3}\6 Заключение СИ (—)']
    dest7 = [rf'{name3}\7 Программа проверки произв (+)',
             rf'{name3}\7 Программа проверки произв (—)']
    dest8 = [rf'{name3}\8 Акт ПП (+)',
             rf'{name3}\8 Акт ПП (—)']
    dest9 = [rf'{name3}\9 Распоряжение на анализ (+)',
             rf'{name3}\9 Распоряжение на анализ (—)']
    dest10 = [rf'{name3}\10 Решение о выдаче (+)',
              rf'{name3}\10 Решение о выдаче (—)']
    dest11 = [rf'{name3}\11 Сертификат (+)',
              rf'{name3}\11 Сертификат (—)']
    dest12 = [rf'{name3}\12 Доп.материалы (+)',
              rf'{name3}\12 Доп.материалы (—)']
    
    listF = [list0[0], list1[0], list2[0], list3[0], list4[0], list5[0], list6[0], list7[0], list8[0], list9[0],
             list10[0], list11[0], list12[0]]

    listDest = [dest0, dest1, dest2, dest3, dest4, dest5, dest6, dest7, dest8, dest9, dest10, dest11, dest12]


    for j in range(0, 13):
        try:
            contents = os.listdir(listF[j])
            # Для папок с 1 файлом
            if len(listDest[j]) == 2 and j != 3 and j != 5:
                if len(contents) == 1 and listF[j] != listDest[j][0]:
                    os.rename(listF[j], listDest[j][0])
                elif len(contents) == 0 and listF[j] != listDest[j][1]:
                    os.rename(listF[j], listDest[j][1])
                elif len(contents) > 1 and j != 12:
                    worrymessage(listF[j], len(contents), '1')

            # Для папки 3 отдельно
            elif j == 3:
                if len(contents) == 4 and listF[j] != listDest[j][0]:
                    os.rename(listF[j], listDest[j][0])
                elif len(contents) == 3 and listF[j] != listDest[j][1]:
                    os.rename(listF[j], listDest[j][1])
                elif len(contents) > 4:
                    worrymessage(listF[j], len(contents), '1')

            # Для папки 5 отдельно
            elif j == 5:
                if len(contents) > 0 and listF[j] != listDest[j][0]:
                    os.rename(listF[j], listDest[j][0])
                elif len(contents) == 0 and listF[j] != listDest[j][1]:
                    os.rename(listF[j], listDest[j][1])
                elif len(contents) > 20:
                    worrymessage(listF[j], len(contents), 'не так много')

            # Для папок с 2 файлами
            elif len(listDest[j]) == 3:
                if len(contents) == 2 and listF[j] != listDest[j][0]:
                    os.rename(listF[j], listDest[j][0])
                elif len(contents) == 1 and listF[j] != listDest[j][1]:
                    os.rename(listF[j], listDest[j][1])
                elif len(contents) == 0 and listF[j] != listDest[j][2]:
                    os.rename(listF[j], listDest[j][2])
                elif len(contents) > 2 and j != 3:
                    worrymessage(listF[j], len(contents), '2')
        except FileNotFoundError:
            print(f"Папка '{listF[j]}' не существует.")

print('Проверка завершена!')
for i in range(0, len(worry), 3):
    print('Пожалуйста проверьте папку:', worry[i], 'там находится ', worry[i + 1], 'файла, вместо', worry[i + 2])