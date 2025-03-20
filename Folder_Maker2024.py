import os
import glob
from tqdm import tqdm


# Сертификаты 2024 года
Num_SERT =  [265, 266, 267, 268, 269, 270, 271, 272, 273, 274,
    275, 276, 277, 278, 279, 280, 281, 282, 283, 284,
    285, 286, 287, 288, 289, 290, 291, 292, 293, 294,
    295, 296, 297, 298, 299, 300, 301, 353, 354, 355,
    357]

folder_names = ["265 - Башкирэнерго-КЭС (29.03.2024)",
    "266 - ОЭСК (Омск) (16.05.2024)",
    "267 - Башкирэнерго-БЭС (30.05.2024)",
    "268 - Башкирэнерго-КЭС (30.05.2024)",
    "269 - Башкирэнерго-ОЭС (30.05.2024)",
    "270 - Башкирэнерго-ОЭС (30.05.2024)",
    "271 - Башкирэнерго-ЦЭС (30.05.2024)",
    "272 - ОборонЭнерго-Хаб. (30.05.2024) нет сертификата",
    "273 - ОборонЭнерго-ЕАО (30.05.2024)",
    "274 - ОборонЭнерго-Якутск (30.05.2024)",
    "275 - Барнаул",
    "276 - Ленэнерго-СПбВС (03.06.24)",
    "277 - Дагэнерго (28.06.2024)",
    "278 - Свердловэнерго (24.12.2024)",
    "279 - Башкирэнерго-ОЭС (28.06.2024)",
    "280 - Башкирэнерго-ОЭС (28.06.2024)",
    "281 - Башкирэнерго-УГЭС (28.06.2024)",
    "282 - Башкирэнерго-УГЭС (28.06.2024)",
    "283 - Башкирэнерго-ЦЭС (28.06.2024)",
    "284 - Башкирэнерго-НЭС (30.07.2024)",
    "285 - Башкирэнерго-НЭС (30.07.2024)",
    "286 - Башкирэнерго-КЭС (30.07.2024)",
    "287 - Алексинская ЭСК (31.07.2024)",
    "288 - Башкирэнерго БЭС (28.08.2024)",
    "289 - ОборонЭнерго Камч. (27.09.2024)",
    "290 - Башкирэнерго-БЭС (27.09.2024)",
    "291 - Башкирэнерго-СЭС (27.09.2024)",
    "292 - Башкирэнерго-ОЭС (27.09.2024)",
    "293 - Башкирэнерго-ОЭС (27.09.2024)",
    "294 - Башкирэнерго-НЭС (27.09.2024)",
    "295 - Северное ПМЭС ФСК (27.09.2024)",
    "296 - Ленэнерго-ЮЭС (17.10.2024)",
    "297 - Ленэнерго-СЭС (04.12.2024)",
    "298 - Башкирэнерго-СВЭС (29.10.2024)",
    "299 - Башкирэнерго-ОЭС (29.10.2024)",
    "300 - Башкирэнерго-ЦЭС (29.10.2024)",
    "301 - Башкирэнерго-ЦЭС (29.10.2024)",
    "353 - Башкирэнерго-ЦЭС (29.11.2024)",
    "354 - Башкирэнерго-ЦЭС (29.11.2024)",
    "355 - Череповец (06.12.2024)",
    "357 - Ленэнерго-КС (13.12.2024)"]

namef = []

for i in tqdm(range(0, len(Num_SERT))):

    name = folder_names[i]
    pathSI = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}\СИ'
    pathIK1 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}\ИК1'
    pathIK2 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}\ИК2'
    pathIK3 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2024\{name}\ИК3'

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

    # Создание папки ИК1 и ее содержимого
    if not os.path.exists(pathIK1):
        for j in range(0, len(pathlistIK1)):
            if not os.path.exists(pathlistIK1[j]):
                os.makedirs(pathlistIK1[j])

    # Создание папки ИК2 и ее содержимого
    if not os.path.exists(pathIK2):
        for j in range(0, len(pathlistIK2)):
            if not os.path.exists(pathlistIK2[j]):
                os.makedirs(pathlistIK2[j])

   # Создание папки ИК3 и ее содержимого
    if not os.path.exists(pathIK3):
        for j in range(0, len(pathlistIK3)):
            if not os.path.exists(pathlistIK3[j]):
                os.makedirs(pathlistIK3[j])

        
print('Successful!')