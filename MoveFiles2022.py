import os
import glob
from tqdm import tqdm
import shutil


# Сертификаты 2022 года
# Num_SERT =  [164, 165, 166, 167, 168, 169, 170, 171, 172, 173, 174, 175, 176, 177, 178,
#     179, 180, 181, 182, 183, 184, 185, 186, 187, 188, 189, 190, 191, 192, 193,
#     194, 195, 196, 197, 198, 199, 200]

Num_SERT =  [192, 193,
    194, 195, 196, 197, 198, 199, 200]

folder_names = [
    "196 - Ленэнерго - СЭС (20.12.2022)",
    "197 - СевКавказЭнерго (23.12.22)",
    "198 - Ленэнерго - КнЭС (27.12.2022)",
    "199 - Энергосеть-Узловая (27.12.22)",
    "200 - Оборонэнерго-Коми (29.12.22)"]

for i in tqdm(range(0, len(Num_SERT))):

    name = folder_names[i]
    pathSI = rf'\\192.168.34.9\линвит\ПОЛЬЗОВАТЕЛИ\USER49\!Сертификаты_2022\{name}\СИ'
    pathSRC = rf'\\192.168.34.9\линвит\ОБЩИЕ ДОКУМЕНТЫ\ЛИНВИТ\12. Обмен файлами\1. Чурсанов Андрей\Сертификаты ЛИНВИТ\_Сертификаты-2022 (164-200)\{name}'

    contents = os.listdir(pathSRC)

    for contents in contents:
        source_path = os.path.join(pathSRC, contents)
        dest_path = os.path.join(pathSI, contents)

        if os.path.isfile(source_path):
            shutil.copy2(source_path, dest_path)
            print(f"Файл скопирован: {source_path} -> {dest_path}")
        elif os.path.isdir(source_path):
            shutil.copytree(source_path, dest_path, dirs_exist_ok=True)
            print(f"Папка скопирована: {source_path} -> {dest_path}")
        else:
            print(f"Неизвестный тип файла: {source_path}")
