import os
import glob
from tqdm import tqdm

# ПУТЬ К КОРНЕВОЙ ПАПКЕ
INpath = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_РЖД'

# Названия папок
folder_names = os.listdir(INpath)

worry = [] # непонятки

FileQNT = [2, 1, 1, 4, 1, 'Any', 1, 1, 1, 1, 2, 2, 'Any', 1, 1, 1]

def worrymessage(pt, nm1, nm2):
    worry.append(pt)
    worry.append(nm1)
    worry.append(nm2)

for i in tqdm(range(0, 2), desc='Идет проверка папок'):

    name = folder_names[i]

    pathSI = rf'{INpath}\{name}\0. СИ'
    pathIK1 = rf'{INpath}\{name}\1. ИК-1'
    pathIK2 = rf'{INpath}\{name}\2. ИК-2'


    path0 = glob.glob(rf'{pathSI}\0 Заявка и приложение*')
    path1 = glob.glob(rf'{pathSI}\1 Распоряжение по заявке*')
    path2 = glob.glob(rf'{pathSI}\2 Решение по заявке*')
    path3 = glob.glob(rf'{pathSI}\3 Заключения по ОМД и ТД*')
    path4 = glob.glob(rf'{pathSI}\4 Акт выбора ПК*')
    path5 = glob.glob(rf'{pathSI}\5 Протоколы СИ*')
    path6 = glob.glob(rf'{pathSI}\6 Заключение СИ*')
    path7 = glob.glob(rf'{pathSI}\7 Программа проверки произ*')
    path8 = glob.glob(rf'{pathSI}\8 Акт ПП*')
    path9 = glob.glob(rf'{pathSI}\9 Распоряжение на анализ*')
    path10 = glob.glob(rf'{pathSI}\10 Решение о выдаче*')
    path11 = glob.glob(rf'{pathSI}\11 Сертификат*')
    path12 = glob.glob(rf'{pathSI}\12 Доп.материалы*')

    path3_1 = glob.glob(rf'{pathSI}\3 Заключения по ОМД и ТД*\3.1 Заключение ОМД*')
    path3_2 = glob.glob(rf'{pathSI}\3 Заключения по ОМД и ТД*\3.2 Заключение ТД и РПН*')
    path3_3 = glob.glob(rf'{pathSI}\3 Заключения по ОМД и ТД*\3.3 Заключение ПМ*')

    pathIK1_0 = glob.glob(rf'{pathIK1}\0 Распоряжение*')
    pathIK1_1 = glob.glob(rf'{pathIK1}\1 Письмо-уведомление*')
    pathIK1_2 = glob.glob(rf'{pathIK1}\2 Программа ИК*')
    pathIK1_3 = glob.glob(rf'{pathIK1}\3 Программа проверки произ*')
    pathIK1_4 = glob.glob(rf'{pathIK1}\4 Акт выбора ПК*')
    pathIK1_5 = glob.glob(rf'{pathIK1}\5 Акт проверки производства*')
    pathIK1_6 = glob.glob(rf'{pathIK1}\6 Протоколы ИК*')
    pathIK1_7 = glob.glob(rf'{pathIK1}\7 Акт по результатам ИК*')
    pathIK1_8 = glob.glob(rf'{pathIK1}\8 Распоряжение на анализ*')
    pathIK1_9 = glob.glob(rf'{pathIK1}\9 Решение по ИК*')
    pathIK1_10 = glob.glob(rf'{pathIK1}\10 Доп. материалы*')

    pathIK2_0 = glob.glob(rf'{pathIK2}\0 Распоряжение*')
    pathIK2_1 = glob.glob(rf'{pathIK2}\1 Письмо-уведомление*')
    pathIK2_2 = glob.glob(rf'{pathIK2}\2 Программа ИК*')
    pathIK2_3 = glob.glob(rf'{pathIK2}\3 Программа проверки произ*')
    pathIK2_4 = glob.glob(rf'{pathIK2}\4 Акт выбора ПК*')
    pathIK2_5 = glob.glob(rf'{pathIK2}\5 Акт проверки производства*')
    pathIK2_6 = glob.glob(rf'{pathIK2}\6 Протоколы ИК*')
    pathIK2_7 = glob.glob(rf'{pathIK2}\7 Акт по результатам ИК*')
    pathIK2_8 = glob.glob(rf'{pathIK2}\8 Распоряжение на анализ*')
    pathIK2_9 = glob.glob(rf'{pathIK2}\9 Решение по ИК*')
    pathIK2_10 = glob.glob(rf'{pathIK2}\10 Доп. материалы*')


    dest0 = rf'{pathSI}\0 Заявка и приложение'
    dest1 = rf'{pathSI}\1 Распоряжение по заявке'
    dest2 = rf'{pathSI}\2 Решение по заявке'
    dest3 = rf'{pathSI}\3 Заключения по ОМД и ТД'
    dest4 = rf'{pathSI}\4 Акт выбора ПК'
    dest5 = rf'{pathSI}\5 Протоколы СИ'
    dest6 = rf'{pathSI}\6 Заключение СИ'
    dest7 = rf'{pathSI}\7 Программа проверки произв'
    dest8 = rf'{pathSI}\8 Акт ПП'
    dest9 = rf'{pathSI}\9 Распоряжение на анализ'
    dest10 = rf'{pathSI}\10 Решение о выдаче'
    dest11 = rf'{pathSI}\11 Сертификат'
    dest12 = rf'{pathSI}\12 Доп.материалы'
    
    dest3_1 = rf'{pathSI}\3 Заключения по ОМД и ТД\3.1 Заключение ОМД'
    dest3_2 = rf'{pathSI}\3 Заключения по ОМД и ТД\3.2 Заключение ТД и РПН'
    dest3_3 = rf'{pathSI}\3 Заключения по ОМД и ТД\3.3 Заключение ПМ'
    

    listF = [path0[0], path1[0], path2[0], path3[0], path4[0], path5[0], path6[0], path7[0], path8[0], path9[0],
             path10[0], path11[0], path12[0]]

    ListF1 = os.listdir(pathSI)
    print(listF, ListF1)

    listDest = [dest0, dest1, dest2, dest3, dest4, dest5, dest6, dest7, dest8, dest9, dest10, dest11, dest12]

    for j in range(0, len(listF)):
        try:
            contents = os.listdir(listF[j])

            File = listF[j]
            Pos_dest = str(listDest[j]) + ' (+)'
            Norm_dest = str(listDest[j]) + ' (+—)'
            Neg_dest = str(listDest[j]) + ' (—)'

            # Для папок с 1 файлом
            if FileQNT[j] == 1:
                if len(contents) == 1 and File != Pos_dest:
                    os.rename(File, Pos_dest)
                elif len(contents) == 0 and File != Neg_dest:
                    os.rename(File, Neg_dest)
                elif len(contents) > 1:
                    worrymessage(File, len(contents), '1')

            # Для папки 3 отдельно
            elif FileQNT[j] == 4:
                if len(contents) == 4 and File != Pos_dest:
                    os.rename(File, Pos_dest)
                elif len(contents) < 4 and File != Neg_dest:
                    os.rename(File, Neg_dest)
                elif len(contents) > 4:
                    worrymessage(File, len(contents), '4')

            # Для папки 5 и 12 отдельно
            elif FileQNT[j] == 'Any':
                if len(contents) > 1 and File != Pos_dest:
                    os.rename(File, Pos_dest)
                elif len(contents) == 0 and File != Neg_dest:
                    os.rename(File, Neg_dest)
                elif len(contents) > 20:
                    worrymessage(File, len(contents), 'не так много (возможно)')

            # Для папок с 2 файлами
            elif FileQNT[j] == 2:
                if len(contents) == 2 and File != Pos_dest:
                    os.rename(File, Pos_dest)
                elif len(contents) == 1 and File != Norm_dest:
                    os.rename(File, Norm_dest)
                elif len(contents) == 0 and File != Neg_dest:
                    os.rename(File, Neg_dest)
                elif len(contents) > 2:
                    worrymessage(File, len(contents), '2')
        except FileNotFoundError:
            print(f"Папка '{listF[j]}' не существует.")

print('Проверка завершена!')
for i in range(0, len(worry), 3):
    print('Пожалуйста проверьте папку:', worry[i], 'там находится ', worry[i + 1], 'файла, вместо', worry[i + 2])