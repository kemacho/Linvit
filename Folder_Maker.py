import os
import glob
from tqdm import tqdm

a = 329 # Первая папка
b = 347 # Последняя папка
INpath = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024'
namef = []

for i in tqdm(range(a, b+1)):
    name = str(i) + ' - ' + namef[i]
    path = os.path.join(INpath, name)

    path0 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\0 Заявка и приложение'
    path1 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\1 Распоряжение по заявке'
    path2 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\2 Решение по заявке'
    path3 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\3 Заключения по ОМД и ТД'
    path4 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\4 Акт выбора ПК'
    path5 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\5 Протоколы СИ'
    path6 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\6 Заключение СИ'
    path7 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\7 Программа проверки произ'
    path8 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\8 Акт ПП'
    path9 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\9 Распоряжение на анализ'
    path10 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\10 Решение о выдаче'
    path11 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\11 Сертификат'
    path12 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\12 Доп.материалы'

    pathList = [path0, path1, path2, path3, path4, path5, path6, path7, path8, path9, path10, path11, path12]


    if not os.path.exists(path):

        path3_1 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\3 Заключения по ОМД и ТД\3.1 Заключение ОМД'
        path3_2 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\3 Заключения по ОМД и ТД\3.2 Заключение ТД и РПН'
        path3_3 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\3 Заключения по ОМД и ТД\3.3 Заключение ПМ'
        list3List = [path3_1, path3_2, path3_3]

        pathList.append(path3_1)
        pathList.append(path3_2)
        pathList.append(path3_3)

        for j in range(-1, 16):
            if not os.path.exists(pathList[j]):
                os.makedirs(pathList[j])
        # for j in range(-1, 3):
        #     if os.path.exists(list3List[j]):
        #         os.makedirs(list3List[j])

    else:
        list0 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\0*')
        list1 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\1*')
        list2 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\2*')
        list3 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\3*')
        list4 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\4*')
        list5 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\5*')
        list6 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\6*')
        list7 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\7*')
        list8 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\8*')
        list9 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\9*')
        list10 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\10*')
        list11 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\11*')
        list12 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\12*')

        listList = [list0, list1, list2, list3, list4, list5, list6, list7, list8, list9, list10, list11, list12]

        dir3_name = list3[0].split('\\')
        path3_1 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\{dir3_name[-1]}\3.1 Заключение ОМД'
        path3_2 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\{dir3_name[-1]}\3.2 Заключение ТД и РПН'
        path3_3 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД1\{name}\{dir3_name[-1]}\3.3 Заключение ПМ'
        list3List = [path3_1, path3_2, path3_3]

        for j in range(-1, 13):
            if not os.path.exists(listList[j][0]):
                os.makedirs(pathList[j])

        for j in range(-1, 3):
            if not os.path.exists(list3List[j]):
                os.makedirs(list3List[j])
        
print('Successful!')