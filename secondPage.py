from docx import Document
import re

def remove_text_in_brackets(docx_path, output_path):
    # Открываем документ

    doc = Document(docx_path)

    # Регулярное выражение для поиска текста в формате (из строки (номер строки))
    bracket_pattern = re.compile(r"\((из строки \(\d+\)|из строк \(\d+[\-,\d\s]*\)|Из строки \(\d+\)|Из строк \(\d+[\-,\d\s]*\))\)")

    # Проходим по каждому абзацу документа
    for paragraph in doc.paragraphs:
        if paragraph.text:  # Если текст не пустой
            # Удаляем текст, который соответствует шаблону
            new_text = bracket_pattern.sub("", paragraph.text)
            paragraph.text = new_text

    # Сохраняем отредактированный файл
    doc.save(output_path)

# Использование

