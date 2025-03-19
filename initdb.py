import os
import sqlite3
from docx import Document

def init_db():
    conn = sqlite3.connect('C:\\ChatMetodicBot\\database.db')  # Указываем путь к базе данных
    c = conn.cursor()
    c.execute('''
        CREATE TABLE IF NOT EXISTS topics (
            id INTEGER PRIMARY KEY,
            name TEXT,
            content TEXT
        )
    ''')
    conn.commit()
    conn.close()

def extract_text_from_word(file_path):
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    return '\n'.join(full_text)

def add_topic_from_word(name, content):
    conn = sqlite3.connect('C:\\ChatMetodicBot\\database.db')
    c = conn.cursor()
    c.execute("INSERT INTO topics (name, content) VALUES (?, ?)", (name, content))
    conn.commit()
    conn.close()

def load_files_from_directory(directory):
    for filename in os.listdir(directory):
        if filename.endswith(".docx"):
            file_path = os.path.join(directory, filename)
            content = extract_text_from_word(file_path)
            topic_name = os.path.splitext(filename)[0]  # Используем имя файла как имя темы
            add_topic_from_word(topic_name, content)
            print(f"Файл {filename} добавлен в базу данных.")

if __name__ == "__main__":
    init_db()
    load_files_from_directory("C:\\ChatMetodicBot")
