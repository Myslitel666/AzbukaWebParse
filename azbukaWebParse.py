import os
import subprocess
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENTATION
import requests
from bs4 import BeautifulSoup

import requests
from bs4 import BeautifulSoup

# URL страницы
url = "https://azbyka.ru/otechnik/Ioann_Kassian_Rimljanin/pisaniya_k_desyati/"

# Заголовки, чтобы имитировать браузер (некоторые сайты блокируют ботов)
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

# Получаем HTML страницы
response = requests.get(url, headers=headers)
response.encoding = 'utf-8'  # Явно указываем кодировку

# Парсим HTML
soup = BeautifulSoup(response.text, 'html.parser')

# Ищем все ссылки, которые ведут на главы (href начинается с './' или '#')
# По HTML-коду видно, что ссылки имеют формат href="./1", href="./1_1", href="#0_1" и т.д.
links = soup.find_all('a', href=True)

# Выводим все ссылки, которые содержат в href признаки глав
print("=" * 80)
print("Ссылки на главы:")
print("=" * 80)

for link in links:
    href = link.get('href', '')
    # Проверяем, что это ссылка на главу (начинается с ./ или # и содержит цифры)
    if (href.startswith('./') and href[2:].isdigit()) or \
       (href.startswith('./') and '_' in href and href[2:].split('_')[0].isdigit()) or \
       (href.startswith('#') and href[1:].isdigit()):
        
        # Получаем текст ссылки (название главы)
        text = link.get_text(strip=True)
        if text:  # Если текст не пустой
            print(f"{href} -> {text}")

print("=" * 80)
print("Всего найдено ссылок:", len([l for l in links if (l.get('href', '').startswith('./') or l.get('href', '').startswith('#')) and l.get_text(strip=True)]))

# ========== ПРОВЕРКА И ЗАКРЫТИЕ ФАЙЛА ==========
file_name = 'оформленный_документ.docx'

# Проверяем, существует ли файл
if os.path.exists(file_name):
    try:
        # Пытаемся удалить файл (если он открыт, будет ошибка)
        os.remove(file_name)
        print(f'Старый файл "{file_name}" удалён')
    except PermissionError:
        # Если файл открыт, просим пользователя закрыть его
        print(f'Файл "{file_name}" открыт!')
        print('Пожалуйста, закройте файл в Word и нажмите Enter...')
        input()
        # После закрытия пробуем удалить снова
        os.remove(file_name)
        print(f'Файл "{file_name}" удалён')

# ========== СОЗДАЁМ ДОКУМЕНТ ==========
print('Создаю документ...')

# Создаём документ
doc = Document()

# ========== НАСТРОЙКА ПОЛЕЙ ==========
section = doc.sections[0]
section.top_margin = Inches(1)      # Верхнее поле 1 дюйм
section.bottom_margin = Inches(1)   # Нижнее поле 1 дюйм
section.left_margin = Inches(1)     # Левое поле 1 дюйм
section.right_margin = Inches(1)    # Правое поле 1 дюйм

# ========== ДОБАВЛЯЕМ ЗАГОЛОВКИ ==========

# Заголовок 1 (с центрированием)
heading1 = doc.add_heading('Главный заголовок', level=1)
heading1.alignment = WD_ALIGN_PARAGRAPH.CENTER

# Настройка шрифта для заголовка
for run in heading1.runs:
    run.font.size = Pt(24)
    run.font.name = 'Arial Narrow'
    run.font.bold = True

# Заголовок 2 (с центрированием)
heading2 = doc.add_heading('Подзаголовок', level=2)
heading2.alignment = WD_ALIGN_PARAGRAPH.CENTER

for run in heading2.runs:
    run.font.size = Pt(18)
    run.font.name = 'Arial Narrow'
    run.font.bold = True

# ========== ДОБАВЛЯЕМ ОБЫЧНЫЙ ТЕКСТ ==========

# Обычный абзац
paragraph = doc.add_paragraph('Э1то пример обычного текста документа. Здесь можно написать любой текст, который должен быть в документе.')

# Настройка шрифта для обычного текста
for run in paragraph.runs:
    run.font.size = Pt(12)
    run.font.name = 'Arial Narrow'

# Ещё один абзац с выравниванием по ширине
paragraph2 = doc.add_paragraph('Это второй абзац с дополнительным текстом. Он будет выровнен по ширине для более аккуратного вида.')
paragraph2.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY

for run in paragraph2.runs:
    run.font.size = Pt(12)
    run.font.name = 'Arial Narrow'

# ========== СОХРАНЯЕМ ДОКУМЕНТ ==========
doc.save(file_name)

print(f'Документ "{file_name}" успешно создан!')

# ========== ОТКРЫВАЕМ ФАЙЛ ==========
try:
    # Для Windows
    if os.name == 'nt':
        os.startfile(file_name)
    # Для Mac
    elif os.name == 'posix':
        subprocess.call(['open', file_name])
    # Для Linux
    else:
        subprocess.call(['xdg-open', file_name])
    print(f'Файл "{file_name}" открыт')
except Exception as e:
    print(f'Не удалось открыть файл автоматически: {e}')
    print(f'Вы можете открыть его вручную: {os.path.abspath(file_name)}')