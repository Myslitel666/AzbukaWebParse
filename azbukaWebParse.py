import os
import subprocess
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENTATION
import requests
from bs4 import BeautifulSoup

# URL страницы
url = "https://azbyka.ru/otechnik/Ioann_Kassian_Rimljanin/pisaniya_k_desyati/"

# Заголовки, чтобы имитировать браузер
headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
}

# Получаем HTML страницы
print("Загружаю страницу...")
response = requests.get(url, headers=headers)
response.encoding = 'utf-8'

# Парсим HTML
soup = BeautifulSoup(response.text, 'html.parser')

# Ищем все ссылки на главы
links = soup.find_all('a', href=True)

# Собираем ссылки в список с определением уровня
chapters = []
for link in links:
    href = link.get('href', '')
    # Получаем класс ссылки
    span = link.find('span')
    if span:
        span_class = span.get('class', [])
        span_class = span_class[0] if span_class else ''
    else:
        span_class = ''
    
    text = link.get_text(strip=True)
    
    # Проверяем, что это ссылка на главу
    if (href.startswith('./') and href[2:].isdigit()) or \
       (href.startswith('./') and '_' in href and href[2:].split('_')[0].isdigit()) or \
       (href.startswith('#') and href[1:].isdigit()):
        
        if text:
            # Определяем уровень заголовка
            if span_class == 'h2o':
                level = 1  # Заголовок 1 уровня (основные собеседования)
            elif span_class == 'h3o':
                level = 2  # Заголовок 2 уровня (главы внутри собеседований)
            else:
                level = 0  # Другое (например, введение)
            
            chapters.append({
                'href': href,
                'title': text,
                'level': level,
                'class': span_class
            })

print(f"Найдено элементов: {len(chapters)}")
print(f"Из них:")
print(f"  - Заголовки 1 уровня (h2o): {len([c for c in chapters if c['level'] == 1])}")
print(f"  - Заголовки 2 уровня (h3o): {len([c for c in chapters if c['level'] == 2])}")
print(f"  - Прочее: {len([c for c in chapters if c['level'] == 0])}")

# ========== ПРОВЕРКА И ЗАКРЫТИЕ ФАЙЛА ==========
file_name = 'оформленный_документ.docx'

if os.path.exists(file_name):
    try:
        os.remove(file_name)
        print(f'Старый файл "{file_name}" удалён')
    except PermissionError:
        print(f'Файл "{file_name}" открыт!')
        print('Пожалуйста, закройте файл в Word и нажмите Enter...')
        input()
        os.remove(file_name)
        print(f'Файл "{file_name}" удалён')

# ========== СОЗДАЁМ ДОКУМЕНТ ==========
print('Создаю документ...')

# Создаём документ
doc = Document()

# ========== НАСТРОЙКА ПОЛЕЙ ==========
section = doc.sections[0]
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)
section.left_margin = Inches(1)
section.right_margin = Inches(1)

# ========== ДОБАВЛЯЕМ ЗАГОЛОВОК ДОКУМЕНТА ==========
main_title = doc.add_heading('Содержание книги "Писания к десяти"', level=1)
main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER

for run in main_title.runs:
    run.font.size = Pt(24)
    run.font.name = 'Arial Narrow'
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)  # Чёрный цвет

# ========== ДОБАВЛЯЕМ СПИСОК ГЛАВ ==========

# Добавляем пустую строку для отступа
doc.add_paragraph()

# Перебираем все найденные главы и добавляем их в документ
for chapter in chapters:
    title = chapter['title']
    level = chapter['level']
    
    if level == 1:
        # Заголовок 1 уровня (основное собеседование) - h2o
        heading = doc.add_heading(title, level=1)
        for run in heading.runs:
            run.font.size = Pt(18)
            run.font.name = 'Arial Narrow'
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)  # Чёрный цвет
        
    elif level == 2:
        # Заголовок 2 уровня (глава внутри собеседования) - h3o
        heading = doc.add_heading(title, level=2)
        for run in heading.runs:
            run.font.size = Pt(14)
            run.font.name = 'Arial Narrow'
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)  # Чёрный цвет
            
    else:
        # Прочее (например, введение) - заголовок 2 уровня с курсивом
        heading = doc.add_heading(title, level=2)
        for run in heading.runs:
            run.font.size = Pt(14)
            run.font.name = 'Arial Narrow'
            run.font.bold = True
            run.font.italic = True
            run.font.color.rgb = RGBColor(0, 0, 0)  # Чёрный цвет

# Добавляем информацию в конце
doc.add_paragraph()
footer_paragraph = doc.add_paragraph(f'Всего добавлено элементов: {len(chapters)}')
footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in footer_paragraph.runs:
    run.font.size = Pt(10)
    run.font.name = 'Arial Narrow'
    run.font.italic = True
    run.font.color.rgb = RGBColor(0, 0, 0)  # Чёрный цвет

# ========== СОХРАНЯЕМ ДОКУМЕНТ ==========
doc.save(file_name)

print(f'Документ "{file_name}" успешно создан!')
print(f'Добавлено {len(chapters)} элементов')

# ========== ОТКРЫВАЕМ ФАЙЛ ==========
try:
    if os.name == 'nt':
        os.startfile(file_name)
    elif os.name == 'posix':
        subprocess.call(['open', file_name])
    else:
        subprocess.call(['xdg-open', file_name])
    print(f'Файл "{file_name}" открыт')
except Exception as e:
    print(f'Не удалось открыть файл автоматически: {e}')
    print(f'Вы можете открыть его вручную: {os.path.abspath(file_name)}')