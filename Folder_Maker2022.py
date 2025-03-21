import os
from tqdm import tqdm

# Сертификаты 2024 года
Num_SERT =  [164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178,
    179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193,
    194, 195, 196, 197, 198, 199, 200]

folder_names = ["164 - Снежногорск (31.03.22)",
    "165 - КГД-Гортранс (31.03.22)",
    "166 - Свердловэнерго (13.04.22)",
    "167- РИР-Краснокаменск (19.05.22)",
    "168 - Ленэнерго - СПбВЭС (03.06.22)",
    "169 - Ленэнерго - ГТЭС (03.06.22)",
    "170 - НТЭК-Норильск (09.06.22)",
    "171 - НТЭК-Дудинка (09.06.22)",
    "172 - Магаданэнерго (15.06.2022)",
    "173 - Агроимпульс (29.06.22)",
    "174 - Энергосеть (29.06.22)",
    "175 - Уральские ЭС (30.06.22)",
    "176 - Уральские ЭС (30.06.22)",
    "177 - Ленэнерго - ТхЭС (05.07.22)",
    "178 - Ленэнерго НлЭС (05.07.22)",
    "179 - Никольская ЭСК (07.07.22)",
    "180 - Архэнерго - КЭС (19.08.22)",
    "181 - Орехово-Зуевская электросеть (01.09.22)",
    "182 - Радугагорэнерго (02.09.2022)",
    "183 - Горэлектросеть г. Тутаев (09.09.2022)",
    "184 - ЮЭС Камчатки (29.09.2022)",
    "185 - Архэнерго - ВЭС (30.09.2022)",
    "186 - Ленэнерго - КС (10.10.2022)",
    "187 - Ленэнерго ВЭС (10.10.2022)",
    "188 - Архэнерго - ПЭС (31.10.2022)",
    "189 - МУП Электросеть Череповец (11.11.2022)",
    "190 - Волгодонск (28.11.22)",
    "191 - Архэнерго - АЭС (30.11.2022)",
    "192-123 AP3 (13.12.22)",
    "193 - ТаймырАльянсТрейдинг (13.12.22)",
    "194 - Оборонэнерго-Арх (14.12.22)",
    "195 - Ленэнерго ЮЭС (20.12.2022)",
    "196 - Ленэнерго СЭС (20.12.2022)",
    "197 - СевКавказЭнерго (23.12.22)",
    "198 - Ленэнерго - КнЭС (27.12.2022)",
    "199 - Энергосеть-Узловая (27.12.22)",
    "200 - Оборонэнерго-Коми (29.12.22)"]

for i in tqdm(range(0, len(Num_SERT))):

    name = folder_names[i]
    pathSI = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}\СИ'
    pathIK1 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}\ИК1'
    pathIK2 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}\ИК2'
    pathIK3 = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}\ИК3'

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