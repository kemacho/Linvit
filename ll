import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinterdnd2 import TkinterDnD, DND_FILES
import docx
import openpyxl
from openpyxl.styles import Alignment
import os
import threading
import re


class WordToExcelConverter(TkinterDnD.Tk):
    def __init__(self):
        super().__init__()
        # ... (остальной код инициализации без изменений)

    def create_widgets(self):
        # ... (без изменений)

    # ... (остальные методы без изменений до extract_data_from_word)

    def split_applicant_text(self, text):
        """Разделяет текст заявителя на 4 части по заданным правилам."""
        # Приводим текст к единому формату
        text = ' '.join(text.split())
        pattern = re.compile(
            r'^(.*?)(?=\s*место\s*нахождения:)'
            r'\s*место\s*нахождения:\s*(.*?)(?=\s*ОГРН)'
            r'\s*ОГРН.*?тел\.\s*(.*?)(?=\s*e-mail:)'
            r'\s*e-mail:\s*(.*)$',
            re.IGNORECASE | re.DOTALL
        )
        match = pattern.search(text)
        if match:
            return (
                match.group(1).strip(),
                match.group(2).strip(),
                match.group(3).strip(),
                match.group(4).strip()
            )
        else:
            return (text, '', '', '')  # Возвращаем исходный текст, если не нашли шаблон

    def extract_data_from_word(self, filepath):
        try:
            doc = docx.Document(filepath)
            z_parts = ['', '', '', '']  # Части данных заявителя
            i_data = []

            for table in doc.tables:
                for row_idx, row in enumerate(table.rows):
                    for cell_idx, cell in enumerate(row.cells):
                        text = cell.text.strip().lower()

                        if "заявитель" in text:
                            # Собираем весь текст из текущей и следующих строк
                            applicant_text = []
                            # Текущая строка
                            current_row = row
                            applicant_text.append(' '.join(
                                [c.text.strip() for c in current_row.cells[cell_idx+1:]]
                            ))
                            # Следующие 3 строки (или меньше, если таблица заканчивается)
                            for j in range(1, 4):
                                if row_idx + j < len(table.rows):
                                    next_row = table.rows[row_idx + j]
                                    applicant_text.append(' '.join(
                                        [c.text.strip() for c in next_row.cells]
                                    ))
                            full_text = ' '.join(applicant_text)
                            z_parts = self.split_applicant_text(full_text)

                        elif "изготовитель" in text:
                            # Старая логика для изготовителя
                            for j in range(1, 5):
                                if row_idx + j < len(table.rows):
                                    if table.rows[row_idx + j].cells:
                                        i_data.append(
                                            table.rows[row_idx + j].cells[0].text.strip()
                                        )

            # Объединяем данные изготовителя
            i_result = "\n".join(i_data[:4]) if i_data else ""
            return z_parts, i_result

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
                # Добавляем заголовки
                headers = ["Часть1 Заявитель", "Часть2 Адрес", "Часть3 Телефон",
                          "Часть4 Email", "Изготовитель"]
                for col, header in enumerate(headers, 1):
                    ws.cell(row=1, column=col, value=header)
                row = 2
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
                row = ws.max_row + 1

            if not save_path:
                self.update_status("Обработка отменена")
                return

            for i, filepath in enumerate(self.file_paths):
                z_parts, i_result = self.extract_data_from_word(filepath)

                if z_parts and i_result is not None:
                    # Запись данных заявителя
                    for col_idx, part in enumerate(z_parts, 1):
                        ws.cell(row=row, column=col_idx, value=part)
                        ws.cell(row=row, column=col_idx).alignment = Alignment(
                            wrap_text=True,
                            vertical='top'
                        )
                    # Запись изготовителя в 5-й столбец
                    ws.cell(row=row, column=5, value=i_result).alignment = Alignment(
                        wrap_text=True,
                        vertical='top'
                    )
                    row += 1

                self.update_progress(i + 1)
                self.update_status(f"Обработан файл {i + 1}/{len(self.file_paths)}")

            # Настройка ширины столбцов
            for col in 'ABCDE':
                ws.column_dimensions[col].width = 30

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