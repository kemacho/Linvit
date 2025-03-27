import os
import shutil
import glob
from tqdm import tqdm

# Номера заявок
Num_Apl = ['01', '02', '03', '04', '05', '06', '07', '08', '09', '10',
    '11', '12', '13', '14', '15', '16', '17', '18', '19', '20',
    '21', '22', '23', '24', '25', '26', '27', '28', '29', '30',
    '33', '34', '35', '36', '37', '38', '39', '40', '41', '42',
    '43', '44', '45', '46', '31', '47', '48', '49', '50', '32',
    '51', '52', '53', '54', '55', '56', '57', '58', '59', '60',
    '61', '62', '63', '64', '65', '66', '67', '68', '69', '70',
    '71', '72', '73', '74', '75', '76', '77', '78', '79', '80',
    '81', '82', '83', '84', '85', '86', '87', '88', '89', '90',
    '91', '92']


# Номер ОС
Num_OS = [
    "ОС-500/24", "ОС-501/24", "ОС-502/24", "ОС-503/24", "ОС-504/24",
    "ОС-505/24", "ОС-496/24", "ОС-484/24", "ОС-485/24", "ОС-486/24",
    "ОС-487/24", "ОС-464/24", "ОС-465/24", "ОС-466/24", "ОС-467/24",
    "ОС-468/24", "ОС-469/24", "ОС-443/24", "ОС-444/24", "ОС-445/24",
    "ОС-446/24", "ОС-447/24", "ОС-448/24", "ОС-449/24", "ОС-450/24",
    "ОС-452/24", "ОС-453/24", "ОС-454/24", "ОС-434/24", "ОС-435/24",
    "ОС-416/24", "ОС-417/24", "ОС-418/24", "ОС-419/24", "ОС-420/24",
    "ОС-421/24", "ОС-422/24", "ОС-423/24", "ОС-424/24", "ОС-479/24",
    "ОС-480/24", "ОС-481/24", "ОС-482/24", "ОС-483/24", "ОС-488/24",
    "ОС-489/24", "ОС-490/24", "ОС-491/24", "ОС-492/24", "ОС-498/24",
    "ОС-499/24", "ОС-493/24", "ОС-494/24", "ОС-495/24", "ОС-470/24",
    "ОС-471/24", "ОС-507/24", "ОС-472/24", "ОС-473/24", "ОС-474/24",
    "ОС-475/24", "ОС-476/24", "ОС-477/24", "ОС-478/24", "ОС-425/24",
    "ОС-426/24", "ОС-427/24", "ОС-428/24", "ОС-429/24", "ОС-430/24",
    "ОС-431/24", "ОС-432/24", "ОС-433/24", "ОС-436/24", "ОС-451/24",
    "ОС-437/24", "ОС-438/24", "ОС-439/24", "ОС-440/24", "ОС-441/24",
    "ОС-442/24", "ОС-506/24", "ОС-455/24", "ОС-456/24", "ОС-457/24",
    "ОС-458/24", "ОС-459/24", "ОС-460/24", "ОС-461/24", "ОС-462/24",
    "ОС-463/24", "ОС-497/24"
        ]

# ПУТЬ К КОРНЕВОЙ ПАПКЕ
INpath = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_РЖД'





# Названия папок
folder_names = os.listdir(INpath)

for i in tqdm(range(0, len(folder_names))):
    source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\Распоряжения\*{Num_OS}-РЗ*')
    dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{folder_names[i]}\0. СИ\1 Распоряжение по заявке*')
    file1 = source[0]
    destf = dest[0]

    contents = os.listdir(destf)
    if len(contents) == 0:
        shutil.copy(file1, destf)
        print(file1, ' => ', destf)



# # заявка
# for i in range(0, len(Num_SERT)):
#     source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\01. Заявки РЖД\*\*{Num_Apl[i]}*')
#     dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\0 Заявка и приложение*')
#     file1 = source[0]
#     destf = dest[0]
#     contents = os.listdir(destf)
#     if len(contents) == 0:
#         shutil.copy(file1, destf)
#         print(file1, ' => ', destf)

# # приложение к заявке
# for i in range(0, len(Num_SERT)):
#     source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\Приложения к заявкам\*{Num_Apl[i]}*')
#     dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\0 Заявка и приложение*')
#     file1 = source[0]
#     destf = dest[0]
#     contents = os.listdir(destf)
#     if len(contents) == 0 or len(contents) == 1:
#         shutil.copy(file1, destf)
#         print(file1, ' => ', destf)
#
# # сертификаты
# for i in range(0, len(Num_SERT)):
#     source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\Сертификаты\{Num_SERT[i]}*\№  РОСС RU*')
#     dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\11 Сертификат*')
#     file1 = source[0]
#     file2 = source[1]
#     destf = dest[0]
#     contents = os.listdir(destf)
#     if len(contents) == 0:
#         shutil.copy(file1, destf)
#         shutil.copy(file2, destf)
#         print(file1, file2, ' => ', destf)
#
# # Акт ПП
# for i in range(0, len(Num_SERT)):
#     source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\Сертификаты\{Num_SERT[i]}*\*АСП*')
#     dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\8 Акт ПП*')
#     file1 = source[0]
#     destf = dest[0]
#     contents = os.listdir(destf)
#     if len(contents) == 0:
#         shutil.copy(file1, destf)
#         print(file1, ' => ', destf)

# # Заключения по ОМД и ТД
# for i in range(0, len(Num_SERT)):
#     source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\Сертификаты\{Num_SERT[i]}*\*ОМД*')
#     dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\3 Заключения по ОМД и ТД*')
#     file1 = source[0]
#     destf = dest[0]
#     contents = os.listdir(destf)
#     if len(contents) == 3:
#         shutil.copy(file1, destf)
#         print(file1, ' => ', destf)

# # Акт выбора ПК
# for i in range(0, len(Num_SERT)):
#     source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\Сертификаты\{Num_SERT[i]}*\*ПК*')
#     dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\4 Акт выбора ПК*')
#     file1 = source[0]
#     destf = dest[0]
#     contents = os.listdir(destf)
#     if len(contents) == 0:
#         shutil.copy(file1, destf)
#         print(file1, ' => ', destf)
#
# # Заключение СИ
# for i in range(0, len(Num_SERT)):
#     source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\Сертификаты\{Num_SERT[i]}*\*ЭЗС*')
#     dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\6 Заключение СИ*')
#     file1 = source[0]
#     destf = dest[0]
#     contents = os.listdir(destf)
#     if len(contents) == 0:
#         shutil.copy(file1, destf)
#         print(file1, ' => ', destf)
#
# # Решение по заявке
# for i in range(0, len(Num_SERT)):
#     source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\Сертификаты\{Num_SERT[i]}*\*{Num_OS}*Р1*')
#     dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\2 Решение по заявке*')
#     file1 = source[0]
#     destf = dest[0]
#     contents = os.listdir(destf)
#     if len(contents) == 0:
#         shutil.copy(file1, destf)
#         print(file1, ' => ', destf)
#
# # Решение о выдаче
# for i in range(0, len(Num_SERT)):
#     source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\Сертификаты\{Num_SERT[i]}*\*{Num_OS}*Р2*(*)*')
#     dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\10 Решение о выдаче*')
#     destf = dest[0]
#     contents = os.listdir(destf)
#     if len(contents) == 0:
#         for j in range(0, len(source)):
#             file1 = source[j]
#             shutil.copy(source[j], destf)
#             shutil.copy(source[j], destf)
#             print(source[j], ' => ', destf)


# listST = [  336, 306, 312, 318, 327, 303, 297, 182, 173, 167, 162, 77, 83, 88, 101, 99,
#             93, 4, 10, 13, 18, 24, 30, 38, 41, 150, 133, 158, 109, 105, 99, 47, 58,
#             48, 76, 71, 53, 62, 50, 215, 221, 194, 201, 207, 232, 249, 237, 223, 240,
#             340, 352, 264, 270, 280, 119, 127, 128, 133, 137, 139, 141, 143, 147, 149,
#             150, 154, 157, 163, 167, 169, 171, 175, 179, 182, 14, 206, 198, 194, 190,
#             200, 214, 210, 121, 122, 141, 147, 129, 154, 110, 119, 112, 289]
#
# listFN = [  339, 311, 317, 326, 335, 305, 302, 193, 181, 172, 166, 82, 87, 92, 109,
#             100, 98, 9, 12, 13, 23, 29, 37, 40, 46, 153, 140, 161, 114, 108, 104,
#             47, 61, 49, 76, 75, 57, 70, 52, 220, 222, 200, 206, 214, 236, 263, 239,
#             231, 248, 351, 363, 269, 279, 288, 126, 127, 132, 136, 138, 140, 142,
#             146, 148, 149, 153, 156, 162, 166, 168, 170, 174, 178, 181, 185, 17,
#             209, 199, 197, 193, 205, 216, 213, 121, 128, 146, 149, 132, 157, 111,
#             120, 118, 296]
#
# # ЭнерТест / Ростест
# listNM = ["ЭнерТест", "ЭнерТест","ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест",
#         "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест",
#         "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест",
#         "ЭнерТест", "Ростест", "Ростест", "Ростест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест", "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "ЭнерТест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "Ростест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест",
# "ЭнерТест"]
#
# # Протоколы СИ
# for i in range(0, len(Num_SERT)):
#     dest = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\5 Протоколы СИ*')
#     destCheck = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Папки РЖД\{Num_SERT[i]} - РЖД\5 Протоколы СИ*\*')
#     destf = dest[0]
#     contents = os.listdir(destf)
#     if listNM[i] == "ЭнерТест" :
#         for j in range(listST[i], listFN[i] + 1):
#             name2024 = '2024-' + str(j)
#             source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\06. Протоколы СИ\ЭнерТест\*\*{name2024}*')
#             file1 = source[0]
#             contents = os.listdir(destf)
#
#             shutil.copy(file1, destf)
#             print(file1, ' => ', destf)
#
#     elif listNM[i] == "Ростест":
#         for j in range(listST[i], listFN[i] + 1):
#             namej = str(j)
#             name533 = 'Протокол № 553'
#             source = glob.glob(rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\06. Протоколы СИ\Ростест\*\*{name533}*{namej}*')
#             file1 = source[0]
#             contents = os.listdir(destf)
#
#             shutil.copy(file1, destf)
#             print(file1, ' => ', destf)

print('Successful!')