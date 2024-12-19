import pandas as pd
import re

# Загружаем данные из файла
file_path = 'your_output_file.xlsx'
df = pd.read_excel(file_path)

# Выводим список столбцов, чтобы проверить их имена
print(df.columns)

# Определяем ключевые слова для поиска
keywords = r'\b(строка|строки|строках|строк|строке|строкам)\b'

# Функция для удаления текста до ключевого слова (включая его)
def remove_before_and_keyword(text):
    match = re.search(keywords, text, flags=re.IGNORECASE)
    if match:
        # Оставляем текст после ключевого слова
        after_keyword = text[match.end():].strip()
        return after_keyword
    return ""  # Если ключевое слово не найдено, возвращаем пустую строку

# Замените 'Текст абзаца.1' на корректное имя столбца
df['E'] = df['Текст абзаца.1'].apply(lambda x: remove_before_and_keyword(str(x)))

# Сохраняем изменения в новый файл
output_file = 'ochkoslona.xlsx'
df.to_excel(output_file, index=False)
