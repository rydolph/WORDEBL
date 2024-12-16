from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import re

def underline_text_in_docx(docx_path, output_path):
    # Открываем документ
    doc = Document(docx_path)

    # Регулярное выражение для поиска ключевых слов и чисел
    pattern = re.compile(r'(\bграфах\b|\bграфе\b|\bГрафах\b|\bГрафе\b)(\s*\d+(?:[-\u2013]\d+)?(?:,\s*\d+)*)?')

    for paragraph in doc.paragraphs:
        matches = list(pattern.finditer(paragraph.text))
        if matches:
            # Создаем новый список для обновлённого текста с подчёркиванием
            runs = []
            last_index = 0

            for match in matches:
                # Добавляем текст до совпадения без изменений
                if match.start() > last_index:
                    runs.append((paragraph.text[last_index:match.start()], False))

                # Добавляем совпавший текст с подчёркиванием
                runs.append((match.group(0), True))
                last_index = match.end()

            # Добавляем оставшийся текст без изменений
            if last_index < len(paragraph.text):
                runs.append((paragraph.text[last_index:], False))

            # Очищаем параграф и добавляем обновлённые части
            paragraph.clear()
            for text, underline in runs:
                run = paragraph.add_run(text)
                if underline:
                    run.font.underline = True

    # Сохраняем документ
    doc.save(output_path)

# Пример использования
input_path = "processed_VPO1Second.docx"  # Исходный документ
output_path = "VPO1output.docx"  # Документ с подчеркнутыми совпадениями
underline_text_in_docx(input_path, output_path)
