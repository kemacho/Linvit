import docx
import openpyxl
from tkinter import Tk, filedialog

def extract_and_save_to_excel():
    """
    Позволяет пользователю выбрать файл Word, извлекает текст между "заявитель" и "изготовитель"
    и сохраняет его в файл Excel.
    """

    # 1. Выбор файла Word
    root = Tk()
    root.withdraw()  # Скрываем основное окно Tkinter
    file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])

    if not file_path:
        print("Файл не выбран.")
        return

    # 2. Извлечение текста
    try:
        doc = docx.Document(file_path)
        full_text = []
        for paragraph in doc.paragraphs:
            full_text.append(paragraph.text)
        full_text = '\n'.join(full_text)  # Объединяем параграфы в одну строку

        start_word = "ЗАЯВИТЕЛЬ"
        end_word = "ИЗГОТОВИТЕЛЬ"

        try:
            start_index = full_text.lower().index(start_word.lower()) + len(start_word)
            end_index = full_text.lower().index(end_word.lower(), start_index) # ищем второе слово после первого
            extracted_text = full_text[start_index:end_index].strip()
        except ValueError:
            print(f"Не удалось найти слова '{start_word}' или '{end_word}' в документе.")
            return

    except Exception as e:
        print(f"Ошибка при чтении файла Word: {e}")
        return

    # 3. Сохранение в Excel
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet['A1'] = "Извлеченный текст"
        sheet['A2'] = extracted_text

        excel_file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])

        if not excel_file_path:
            print("Сохранение отменено.")
            return

        workbook.save(excel_file_path)
        print(f"Текст успешно извлечен и сохранен в {excel_file_path}")

    except Exception as e:
        print(f"Ошибка при сохранении в Excel: {e}")


if __name__ == "__main__":
    extract_and_save_to_excel()
