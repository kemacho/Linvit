import os
import glob
from tqdm import tqdm

# Сертификаты 2022 года
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

worry = []  # непонятки


def worrymessage(pt, nm1, nm2):
    worry.append(pt)
    worry.append(nm1)
    worry.append(nm2)


for i in tqdm(range(0, len(Num_SERT)), desc='Идет проверка папок'):
    name = str(Num_SERT[i])
    name2 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ')
    name3 = name2[0]

    list0 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\0*')
    list1 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\1*')
    list2 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\2*')
    list3 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\3*')
    list4 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\4*')
    list5 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\5*')
    list6 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\6*')
    list7 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\7*')
    list8 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\8*')
    list9 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\9*')
    list10 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\10*')
    list11 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\11*')
    list12 = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}*\СИ\12*')

    dest0 = [rf'{name3}\0 Заявка и приложение (+)',
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
            if len(listDest[j]) == 2 and j != 3 and j != 5 and j != 0 and j != 12:
                if len(contents) == 1 and listF[j] != listDest[j][0]:
                    os.rename(listF[j], listDest[j][0])
                elif len(contents) == 0 and listF[j] != listDest[j][1]:
                    os.rename(listF[j], listDest[j][1])
                elif len(contents) > 1:
                    os.rename(listF[j], listDest[j][0])
                    worrymessage(listF[j], len(contents), '1')

            # Для папки 0 отдельно
            if j == 0:
                if len(contents) == 0 and listF[j] != listDest[j][1]:
                    os.rename(listF[j], listDest[j][1])
                elif len(contents) == 2 and listF[j] != listDest[j][0]:
                    os.rename(listF[j], listDest[j][0])
                elif len(contents) == 1 and listF[j] != listDest[j][1]:
                    os.rename(listF[j], listDest[j][1])
                    worrymessage(listF[j], contents, '1')

            # Для папки 3 отдельно
            elif j == 3:
                if len(contents) == 4 and listF[j] != listDest[j][0]:
                    os.rename(listF[j], listDest[j][0])
                elif len(contents) == 3 and listF[j] != listDest[j][1]:
                    os.rename(listF[j], listDest[j][1])
                elif len(contents) > 4:
                    worrymessage(listF[j], len(contents), '1')

            # Для папки 5 и 12 отдельно
            elif j == 5 or j == 12:
                if len(contents) > 0 and listF[j] != listDest[j][0]:
                    os.rename(listF[j], listDest[j][0])
                elif len(contents) == 0 and listF[j] != listDest[j][1]:
                    os.rename(listF[j], listDest[j][1])
                elif len(contents) > 20:
                    worrymessage(listF[j], len(contents), 'не так много')

        except FileNotFoundError:
            print(f"Папка '{listF[j]}' не существует.")

print('Проверка завершена!')
for i in range(0, len(worry), 3):
    print('Пожалуйста проверьте папку:', worry[i], 'там находится ', worry[i + 1], 'файла, вместо', worry[i + 2])