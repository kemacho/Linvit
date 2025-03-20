import tkinter as tk
from tkinter import filedialog, ttk, messagebox
import openpyxl
import os
from datetime import datetime
import threading
from tkinterdnd2 import DND_FILES, TkinterDnD
from openpyxl.utils import get_column_letter

# Функция для форматирования даты
def format_date(date_value):
    if isinstance(date_value, datetime):
        return date_value.strftime("%d.%m.%Y")
    else:
        return str(date_value or "")

# Функция для форматирования времени
def format_time(date_value):
    if isinstance(date_value, datetime):
        return date_value.strftime("%H:%M")
    else:
        return str(date_value or "")

def process_files(action):
    """
    Обрабатывает выбранные файлы, используя заранее известный индекс листа.
    """
    update_status("Начало обработки файлов...")  # Обновление статуса

    if not file_paths:
        messagebox.showinfo("Внимание", "Файлы не выбраны.")
        update_status("Ожидание выбора файлов.")
        return

    if action == "existing":
        save_path = filedialog.askopenfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            update_status("Обработка прервана пользователем.")
            return

        try:
            output_workbook = openpyxl.load_workbook(save_path)
            output_sheet = output_workbook.active
        except FileNotFoundError:
            messagebox.showerror("Ошибка", f"Файл не найден: {save_path}")
            update_status(f"Ошибка: Файл не найден - {save_path}")
            return
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при открытии файла {save_path}: {e}")
            update_status(f"Ошибка: Не удалось открыть файл - {save_path}. Ошибка: {e}")
            return

    elif action == "new":
        save_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not save_path:
            update_status("Обработка прервана пользователем.")
            return

        output_workbook = openpyxl.Workbook()
        output_sheet = output_workbook.active
        output_sheet['A1'] = "Номер Протокола"
        output_sheet['B1'] = "Адрес"
        output_sheet['C1'] = "Дата начала испытаний"
        output_sheet['D1'] = "Дата окончания испытаний"
        output_sheet['E1'] = "Прибор 1"
        output_sheet['F1'] = "Заводской номер Прибора 1"
        output_sheet['G1'] = "Прибор 1"
        output_sheet['H1'] = "Заводской номер Прибора 2"
    else:
        messagebox.showerror("Ошибка", "Некорректное действие.")
        update_status("Ошибка: Некорректное действие.")
        return

    row_index = output_sheet.max_row + 1
    total_files = len(file_paths)
    progress_bar['maximum'] = total_files
    progress_bar['value'] = 0

    for i, file_path in enumerate(file_paths):
        try:
            update_status(f"Обработка файла {i+1} из {total_files}: {os.path.basename(file_path)}")
            workbook = openpyxl.load_workbook(file_path, data_only=True)
            sheet_names = workbook.sheetnames
            sheet1 = workbook[sheet_names[1]]
            sheet2 = workbook[sheet_names[2]]

            cell_value1 = sheet1['BC29'].value
            cell_value2 = sheet2['R21'].value
            cell_value3 = format_date(sheet2['M26'].value)
            cell_value3_1 = format_time(sheet2['AK26'].value)
            cell_value4 = format_date(sheet2['M27'].value)
            cell_value4_1 = format_time(sheet2['AK27'].value)
            cell_value5 = sheet2['AG35'].value
            cell_value6 = int(sheet2['BE35'].value)
            cell_value7 = sheet2['AG36'].value
            cell_value8 = int(sheet2['BE36'].value)

            output_sheet[f'A{row_index}'] = cell_value1
            output_sheet[f'B{row_index}'] = cell_value2
            output_sheet[f'C{row_index}'] = cell_value3 + ' (' + cell_value3_1 + ')'
            output_sheet[f'D{row_index}'] = cell_value4 + ' (' + cell_value4_1 + ')'
            output_sheet[f'E{row_index}'] = cell_value5
            output_sheet[f'F{row_index}'] = cell_value6
            output_sheet[f'G{row_index}'] = cell_value7
            output_sheet[f'H{row_index}'] = cell_value8

            row_index += 1

        except FileNotFoundError:
            messagebox.showerror("Ошибка", f"Файл не найден: {file_path}")
            update_status(f"Ошибка: Файл не найден - {os.path.basename(file_path)}")
        except Exception as e:
            messagebox.showerror("Ошибка", f"Ошибка при обработке файла {os.path.basename(file_path)}: {e}")
            update_status(f"Ошибка при обработке {os.path.basename(file_path)}: {e}")

        # Обновление прогресс-бара (основной поток)
        root.after(0, update_progress)

    try:
        update_status("Сохранение файла...")
        if action == "new": # Только для новых файлов
            adjust_column_width(output_sheet) #  <----  ВЫЗОВ ФУНКЦИИ

        output_workbook.save(save_path)
        messagebox.showinfo("Успех", f"Данные успешно записаны в файл: {save_path}")
        update_status(f"Данные успешно записаны в файл: {save_path}")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Ошибка при сохранении файла: {e}")
        update_status(f"Ошибка при сохранении файла: {e}")
    finally:
        output_workbook.close()
        update_status("Обработка завершена.")

    root.after(0, processing_complete)  # Сообщение о завершении


def adjust_column_width(sheet):
    for column_cells in sheet.columns:
        length = max(len(str(cell.value)) for cell in column_cells)
        sheet.column_dimensions[get_column_letter(column_cells[0].column)].width = length + 2  # +2 для небольшого отступа

def update_progress():
    """Обновляет значение прогресс-бара."""
    progress_bar['value'] += 1


def processing_complete():
    """Сбрасывает прогресс-бар после завершения."""
    progress_bar['value'] = 0

def update_status(message):
    """Обновляет текст в метке статуса."""
    status_label.config(text=message)
    root.update_idletasks()  # Обновление интерфейса для отображения сообщения


def choose_files():
    """Выбор файлов."""
    global file_paths
    new_file_paths = filedialog.askopenfilenames(filetypes=[("Excel files", "*.xlsx")])
    if new_file_paths:
        file_paths.extend(new_file_paths)
        update_file_list()
    update_status("Файлы выбраны.")


def delete_selected_file():
    """Удаление выбранного файла."""
    try:
        selected_index = file_list.curselection()[0]
        del file_paths[selected_index]
        update_file_list()
    except IndexError:
        messagebox.showinfo("Внимание", "Файл не выбран для удаления.")
    update_status("Файл удален из списка.")


def clear_file_list():
    """Очистка списка файлов."""
    global file_paths
    file_paths = []
    update_file_list()
    update_status("Список файлов очищен.")


def update_file_list():
    """Обновление списка файлов."""
    file_list.delete(0, tk.END)
    for file_path in file_paths:
        file_list.insert(tk.END, os.path.basename(file_path))


def drop(event):
    """Обработчик события Drop."""
    global file_paths
    files = root.tk.splitlist(event.data)
    file_paths.extend(files)
    update_file_list()
    update_status("Файлы добавлены перетаскиванием.")


def start_processing_thread(action):
    """Запуск обработки в отдельном потоке."""
    thread = threading.Thread(target=process_files, args=(action,))
    thread.start()





# == UI Setup ==
root = TkinterDnD.Tk()  # Инициализируем TkinterDnD
root.title("Excel File Transfer")

# Стиль ttk
style = ttk.Style()
style.configure("TButton", padding=5, relief="raised")
style.configure("TLabel", padding=5)

# Рамка для элементов
frame = ttk.Frame(root, padding=10)
frame.pack(expand=True, fill="both")

# Кнопка выбора файлов
choose_button = ttk.Button(frame, text="Выберите файлы или перетащите их в область ниже", command=choose_files, style="TButton")
choose_button.pack(pady=10)

file_list = tk.Listbox(frame, width=60, height=10, selectmode=tk.SINGLE)
file_list.pack(pady=5, expand=True, fill="both")

# Привязка события Drop к Listbox
file_list.drop_target_register(DND_FILES)
file_list.dnd_bind('<<Drop>>', drop)


# Фрейм для кнопок "Удалить" и "Очистить"
button_frame = ttk.Frame(frame)
button_frame.pack(pady=5)

# Кнопки управления списком
delete_button = ttk.Button(button_frame, text="Удалить выбранный файл", command=delete_selected_file, style="TButton")
delete_button.pack(side="left", padx=5)  # side="left" для размещения в ряд

clear_button = ttk.Button(button_frame, text="Очистить список файлов", command=clear_file_list, style="TButton")
clear_button.pack(side="left", padx=5)  # side="left" для размещения в ряд

# Фрейм для кнопок действий
action_button_frame = ttk.Frame(frame)
action_button_frame.pack(pady=5, fill="x")
button_width = 13  # Минимальная ширина, можно настроить

# Кнопки действий
new_button = ttk.Button(action_button_frame, text="Создать новую таблицу", command=lambda: start_processing_thread("new"), style="TButton")
new_button.pack(expand=True, fill="x")
new_button.config(width=button_width)

existing_button = ttk.Button(action_button_frame, text="Добавить в существующую таблицу", command=lambda: start_processing_thread("existing"), style="TButton")
existing_button.pack(expand=True, fill="x")
existing_button.config(width=button_width)

# Прогресс-бар
progress_bar = ttk.Progressbar(frame, orient="horizontal", length=300, mode="determinate")
progress_bar.pack(pady=10)

# Метка статуса
status_label = ttk.Label(frame, text="Ожидание выбора файлов.", anchor="center")
status_label.pack(pady=5, fill="x")


# Переменная списка файлов
file_paths = []

# Инициализация статуса
update_status("Ожидание выбора файлов.")

# Запуск UI
root.mainloop()
