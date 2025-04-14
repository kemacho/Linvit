import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import docx
import openpyxl
from openpyxl.styles import Alignment
import os
import threading


class WordToExcelConverter(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        self.title("Word → Excel Converter")
        self.geometry("700x550")
        self.configure(bg="#f0f0f0")
        self.file_paths = []

        # Стили
        self.style = ttk.Style()
        self.style.configure("TButton", padding=6, font=('Segoe UI', 10))
        self.style.configure("TLabel", font=('Segoe UI', 9))
        self.style.configure("Title.TLabel", font=('Segoe UI', 12, 'bold'))

        self.create_widgets()

    def create_widgets(self):
        # Основные фреймы
        header_frame = ttk.Frame(self, padding=10)
        header_frame.pack(fill=tk.X)

        main_frame = ttk.Frame(self, padding=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        ttk.Label(header_frame, text="Конвертер Word в Excel", style="Title.TLabel").pack()

        # Область файлов
        file_frame = ttk.LabelFrame(main_frame, text=" Файлы для обработки ", padding=10)
        file_frame.pack(fill=tk.BOTH, expand=True)

        # Список файлов с прокруткой
        self.file_list = tk.Listbox(
            file_frame,
            selectmode=tk.EXTENDED,
            height=12,
            bg="white",
            fg="#333",
            selectbackground="#4CAF50",
            selectforeground="white",
            font=('Segoe UI', 9),
            relief="flat",
            highlightthickness=0
        )
        self.file_list.pack(fill=tk.BOTH, expand=True, pady=5)

        # Настройка drag and drop
        self.file_list.drop_target_register(DND_FILES)
        self.file_list.dnd_bind('<<Drop>>', self.add_dropped_files)

        # Кнопки управления файлами
        btn_frame = ttk.Frame(file_frame)
        btn_frame.pack(fill=tk.X, pady=(5, 0))

        ttk.Button(btn_frame, text="Добавить", command=self.add_files).pack(side=tk.LEFT, fill=tk.X, expand=True,
                                                                            padx=2)
        ttk.Button(btn_frame, text="Удалить", command=self.remove_selected).pack(side=tk.LEFT, fill=tk.X, expand=True,
                                                                                 padx=2)
        ttk.Button(btn_frame, text="Очистить", command=self.clear_list).pack(side=tk.LEFT, fill=tk.X, expand=True,
                                                                             padx=2)

        # Кнопки обработки
        process_frame = ttk.Frame(main_frame)
        process_frame.pack(fill=tk.X, pady=(10, 0))

        ttk.Button(process_frame, text="Новая таблица",
                   command=lambda: self.start_processing("new")).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)
        ttk.Button(process_frame, text="Добавить в файл",
                   command=lambda: self.start_processing("existing")).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=2)

        # Прогресс и статус
        self.progress = ttk.Progressbar(main_frame, orient="horizontal", mode="determinate")
        self.progress.pack(fill=tk.X, pady=(15, 5))

        self.status = ttk.Label(main_frame, text="Готов к работе")
        self.status.pack(fill=tk.X)

    def add_dropped_files(self, event):
        files = self.tk.splitlist(event.data)
        added = 0
        for f in files:
            if f.lower().endswith(('.doc', '.docx')) and f not in self.file_paths:
                self.file_paths.append(f)
                added += 1
        self.update_file_list()
        self.update_status(f"Добавлено {added} файлов" if added else "Нет новых файлов для добавления")

    def add_files(self):
        files = filedialog.askopenfilenames(
            filetypes=[("Word Files", "*.docx;*.doc"), ("All files", "*.*")],
            title="Выберите файлы Word"
        )
        if files:
            for f in files:
                if f not in self.file_paths:
                    self.file_paths.append(f)
            self.update_file_list()
            self.update_status(f"Добавлено {len(files)} файлов")

    def remove_selected(self):
        selected = self.file_list.curselection()
        if selected:
            for i in reversed(selected):
                del self.file_paths[i]
            self.update_file_list()
            self.update_status(f"Удалено {len(selected)} файлов")
        else:
            self.update_status("Не выбраны файлы для удаления")

    def clear_list(self):
        if self.file_paths:
            self.file_paths = []
            self.update_file_list()
            self.update_status("Список файлов очищен")
        else:
            self.update_status("Список файлов уже пуст")

    def update_file_list(self):
        self.file_list.delete(0, tk.END)
        for f in self.file_paths:
            self.file_list.insert(tk.END, os.path.basename(f))

    def update_status(self, message):
        self.status.config(text=message)
        self.update_idletasks()

    def update_progress(self, value):
        self.progress['value'] = value
        self.update_idletasks()

    def start_processing(self, action):
        if not self.file_paths:
            messagebox.showwarning("Внимание", "Сначала выберите файлы!")
            return

        thread = threading.Thread(target=self.process_files, args=(action,), daemon=True)
        thread.start()

    def extract_data_from_word(self, filepath):
        try:
            doc = docx.Document(filepath)
            z_data = []
            i_data = []

            for table in doc.tables:
                for i, row in enumerate(table.rows):
                    for cell in row.cells:
                        text = cell.text.upper()
                        if "ЗАЯВИТЕЛЬ" in text:
                            for j in range(1, 4):
                                if i + j < len(table.rows):
                                    z_data.append(table.rows[i + j].cells[0].text.strip())
                                else:
                                    z_data.append("")

                        if "ИЗГОТОВИТЕЛЬ" in text:
                            for j in range(1, 4):
                                if i + j < len(table.rows):
                                    i_data.append(table.rows[i + j].cells[0].text.strip())
                                else:
                                    i_data.append("")

            return z_data[:3], i_data[:3]

        except Exception as e:
            self.update_status(f"Ошибка при обработке {os.path.basename(filepath)}: {str(e)}")
            return None, None

    def process_files(self, action):
        self.update_status("Начало обработки...")
        self.progress['maximum'] = len(self.file_paths)
        self.progress['value'] = 0

        try:
            if action == "new":
                wb = openpyxl.Workbook()
                ws = wb.active
                row = 1
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel Files", "*.xlsx")],
                    title="Сохранить новую таблицу как"
                )
            else:
                save_path = filedialog.askopenfilename(
                    filetypes=[("Excel Files", "*.xlsx")],
                    title="Выберите существующий файл Excel"
                )
                if not save_path:
                    self.update_status("Обработка отменена")
                    return
                wb = openpyxl.load_workbook(save_path)
                ws = wb.active
                row = ws.max_row + 2

            if not save_path:
                self.update_status("Обработка отменена")
                return

            for i, filepath in enumerate(self.file_paths):
                z_data, i_data = self.extract_data_from_word(filepath)

                if z_data and i_data:
                    # Запись данных после "ЗАЯВИТЕЛЬ" в столбец A
                    for j in range(3):
                        ws.cell(row=row + j, column=1, value=z_data[j] if j < len(z_data) else "")
                        ws.cell(row=row + j, column=1).alignment = Alignment(wrap_text=True)

                    # Запись данных после "ИЗГОТОВИТЕЛЬ" в столбец B
                    for j in range(3):
                        ws.cell(row=row + j, column=2, value=i_data[j] if j < len(i_data) else "")
                        ws.cell(row=row + j, column=2).alignment = Alignment(wrap_text=True)

                    row += 3

                self.update_progress(i + 1)
                self.update_status(f"Обработан файл {i + 1}/{len(self.file_paths)}")

            # Настройка ширины столбцов
            ws.column_dimensions['A'].width = 50
            ws.column_dimensions['B'].width = 50

            wb.save(save_path)
            self.update_status(f"Данные сохранены в {save_path}")
            messagebox.showinfo("Готово", f"Данные успешно сохранены в:\n{save_path}")

        except Exception as e:
            self.update_status(f"Ошибка: {str(e)}")
            messagebox.showerror("Ошибка", f"Произошла ошибка:\n{str(e)}")
        finally:
            self.progress['value'] = 0


if __name__ == "__main__":
    app = WordToExcelConverter()
    app.mainloop()