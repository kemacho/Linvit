import os
from tqdm import tqdm

# Сертификаты 2021 года
Num_SERT =  [122, 123, 124, 125, 126, 127, 128, 129, 130, 131, 132, 133, 134, 135, 136, 137, 138, 139, 140, 141, 142,
             143, 144, 145, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163]

folder_names = ["122 - Дагэнерго", "123 - КРЭМЗ-Тула (22.01.21)", "124 - Ленэнерго ГТЭС (30.04.21)",
                "125 - Алексин (30.04.21)", "126 - БашЭн-ЦЭС (31.05.21)", "127 - БашЭн-ОЭС (31.05.21)",
                "128 - БашЭн-ОЭС (31.05.21)", "129 - БашЭн-БЭС (31.05.21)", "130 - БашЭн-УГС (29.06.21)",
                "131 - БашЭн-ОЭС (29.06.21)", "132 - БашЭн-ОЭС (29.06.21)", "133 - БашЭн-ЦЭС (29.06.21)",
                "134 - Энергосеть (23.07.21)", "135 - Агроимпульс (23.07.21)", "136 - Барнаул (23.07.21)",
                "137 - БашЭн-НЭС (30.07.21)", "138 - БашЭн-НЭС (30.07.21)", "139 - Байконурэнерго (30.07.21)",
                "140 - ВолжВУЭС (31.07.21)", "141 - ВолжЧЭС (25.08.21)", "142 - БашЭн-БЭС (30.08.21)",
                "143 - Тяжпромарматура (30.08.21)", "144 - БашЭн БЭС (30.09.21)", "145 - БашЭн ОЭС (30.09.21)",
                "146 - БашЭн ОЭС (30.09.21)", "147 - БашЭн НЭС (30.09.21)", "148 - БашЭн СЭС (30.09.21)",
                "149 - ВолжВЭС (30.09.21)", "150 - БашЭн СВЭС (недействителен. PCA 28.08.2024)",
                "151 - БашЭн ЦЭС (недействителен. PCA 28.08.2024)", "152 - БашЭн ЦЭС (недействителен. PCA 28.08.2024)",
                "153 - БашЭн ОЭС (30.10.21) (недействителен. PCA 28.08.2024)",
                "154 - БашЭн КЭС (29.11.21) (недействителен. PCA 28.08.2024)",
                "155 - БашЭн ЦЭС (29.11.21) (недействителен. PCA 28.08.2024)",
                "156 - БашЭн ЦЭС (29.11.21) (недействителен. PCA 28.08.2024)",
                "157 - Северное ПМЭС ФСК (20.12.21) (недействителен. PCA 28.08.2024)",
                "158 - Архэнерго АЭС (20.12.21) (недействителен. PCA 28.08.2024)",
                "159 - Ленэнерго-КЭС (недействителен. ОС ЛИНВИТ 27.12.2022)",
                "160 - ЭСК ЧТПЗ Первоуральск (недействителен. РСА 28.08.2024)",
                "161 - ЧГКЭС Чернушка (20.12.21) (недействителен. РСА 28.08.2024)",
                "162 - Магаданэнергосеть (27.12.21) (недействителен. РСА 28.08.2024)", "163 - Тульские ГЭС (27.12.21)"]

for i in tqdm(range(0, len(Num_SERT))):

    name = folder_names[i]
    pathSI  = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2021\{name}\СИ'
    pathIK1 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2021\{name}\ИК1'
    pathIK2 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2021\{name}\ИК2'
    pathIK3 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2021\{name}\ИК3'

    path0 = rf'{pathSI}\0 Заявка и приложение'
    path1 = rf'{pathSI}\1 Распоряжение по заявке'
    path2 = rf'{pathSI}\2 Решение по заявке'
    path3 = rf'{pathSI}\3 Заключения по ОМД и ТД'
    path4 = rf'{pathSI}\4 Акт выбора ПК'
    path5 = rf'{pathSI}\5 Протоколы СИ'
    path6 = rf'{pathSI}\6 Заключение СИ'
    path7 = rf'{pathSI}\7 Программа проверки произ'
    path8 = rf'{pathSI}\8 Акт ПП'
    path9 = rf'{pathSI}\9 Распоряжение на анализ'
    path10 = rf'{pathSI}\10 Решение о выдаче'
    path11 = rf'{pathSI}\11 Сертификат'
    path12 = rf'{pathSI}\12 Доп.материалы'

    path3_1 = rf'{pathSI}\3 Заключения по ОМД и ТД\3.1 Заключение ОМД'
    path3_2 = rf'{pathSI}\3 Заключения по ОМД и ТД\3.2 Заключение ТД и РПН'
    path3_3 = rf'{pathSI}\3 Заключения по ОМД и ТД\3.3 Заключение ПМ'

    pathIK1_0 = rf'{pathIK1}\0'
    pathIK1_1 = rf'{pathIK1}\1'
    pathIK1_2 = rf'{pathIK1}\2'
    pathIK1_3 = rf'{pathIK1}\3'
    pathIK1_4 = rf'{pathIK1}\4'
    pathIK1_5 = rf'{pathIK1}\5'
    pathIK1_6 = rf'{pathIK1}\6'
    pathIK1_7 = rf'{pathIK1}\7'
    pathIK1_8 = rf'{pathIK1}\8'
    
    pathIK2_0 = rf'{pathIK2}\0'
    pathIK2_1 = rf'{pathIK2}\1'
    pathIK2_2 = rf'{pathIK2}\2'
    pathIK2_3 = rf'{pathIK2}\3'
    pathIK2_4 = rf'{pathIK2}\4'
    pathIK2_5 = rf'{pathIK2}\5'
    pathIK2_6 = rf'{pathIK2}\6'
    pathIK2_7 = rf'{pathIK2}\7'
    pathIK2_8 = rf'{pathIK2}\8'
    
    pathIK3_0 = rf'{pathIK3}\0'
    pathIK3_1 = rf'{pathIK3}\1'
    pathIK3_2 = rf'{pathIK3}\2'
    pathIK3_3 = rf'{pathIK3}\3'
    pathIK3_4 = rf'{pathIK3}\4'
    pathIK3_5 = rf'{pathIK3}\5'
    pathIK3_6 = rf'{pathIK3}\6'
    pathIK3_7 = rf'{pathIK3}\7'
    pathIK3_8 = rf'{pathIK3}\8'

    pathListSI = [path0, path1, path2, path3, path4, path5, path6, path7, path8, path9, path10, path11, path12]
    pathList_3 = [path3_1, path3_2, path3_3]
    pathlistIK1 = [pathIK1_0, pathIK1_1, pathIK1_2, pathIK1_3, pathIK1_4, pathIK1_5, pathIK1_6, pathIK1_7, pathIK1_8]
    pathlistIK2 = [pathIK2_0, pathIK2_1, pathIK2_2, pathIK2_3, pathIK2_4, pathIK2_5, pathIK2_6, pathIK2_7, pathIK2_8]
    pathlistIK3 = [pathIK3_0, pathIK3_1, pathIK3_2, pathIK3_3, pathIK3_4, pathIK3_5, pathIK3_6, pathIK3_7, pathIK3_8]

    # Создание папки СИ и ее содержимого
    if not os.path.exists(pathSI):
        for j in range(0, len(pathListSI)):
            if j == 3 and not os.path.exists(pathListSI[j]):
                os.makedirs(pathListSI[j])
                os.makedirs(pathList_3[0])
                os.makedirs(pathList_3[1])
                os.makedirs(pathList_3[2])
            elif not os.path.exists(pathListSI[j]):
                os.makedirs(pathListSI[j])

   #  # Создание папки ИК1 и ее содержимого
   #  if not os.path.exists(pathIK1):
   #      for j in range(0, len(pathlistIK1)):
   #          if not os.path.exists(pathlistIK1[j]):
   #              os.makedirs(pathlistIK1[j])
   #
   #  # Создание папки ИК2 и ее содержимого
   #  if not os.path.exists(pathIK2):
   #      for j in range(0, len(pathlistIK2)):
   #          if not os.path.exists(pathlistIK2[j]):
   #              os.makedirs(pathlistIK2[j])
   #
   # # Создание папки ИК3 и ее содержимого
   #  if not os.path.exists(pathIK3):
   #      for j in range(0, len(pathlistIK3)):
   #          if not os.path.exists(pathlistIK3[j]):
   #              os.makedirs(pathlistIK3[j])

        
print('Successful!')