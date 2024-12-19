import openpyxl
import re

def copy_data(input_file, output_file):
    # Открываем Excel файл
    wb = openpyxl.load_workbook(input_file)
    sheet = wb.active

    # Проходим по всем строкам в столбце E
    for row in range(2, sheet.max_row + 1):  # Предполагаем, что первая строка — заголовок
        value = sheet[f"E{row}"].value
        if value:
            # Извлекаем данные до первого текста после цифр
            match = re.match(r"(\d{1,3}(?:, \d{1,3})*)", value)
            if match:
                sheet[f"F{row}"].value = match.group(1)

    # Сохраняем изменения в новый файл
    wb.save(output_file)

# Пример использования
input_file = "ochkoslona.xlsx"  # Имя входного файла
output_file = "kogotbobra.xlsx"  # Имя выходного файла
copy_data(input_file, output_file)