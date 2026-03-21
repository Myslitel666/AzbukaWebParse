import os
import subprocess
import time
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
import requests
from bs4 import BeautifulSoup

# URL страницы с оглавлением
url = "https://azbyka.ru/otechnik/Ioann_Kassian_Rimljanin/pisaniya_k_desyati/"

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
}

print("Загружаю оглавление...")
response = requests.get(url, headers=headers, timeout=30)
response.encoding = 'utf-8'
soup = BeautifulSoup(response.text, 'html.parser')

# Находим все ссылки на собеседования (с классом h2o)
conversations = []
for span in soup.find_all('span', class_='h2o'):
    link = span.find_parent('a')
    if link and link.get('href', '').startswith('./'):
        href = link.get('href')
        title = span.get_text(strip=True)
        full_url = f"https://azbyka.ru/otechnik/Ioann_Kassian_Rimljanin/pisaniya_k_desyati/{href[2:]}"
        conversations.append({
            'number': href[2:],
            'title': title,
            'url': full_url
        })

print(f"Найдено собеседований: {len(conversations)}")

# ========== ФУНКЦИЯ ДЛЯ ПАРСИНГА ОДНОГО СОБЕСЕДОВАНИЯ ==========
def parse_conversation(conv_url):
    """Загружает страницу собеседования и парсит все главы с текстом"""
    try:
        print(f"  Загружаю: {conv_url}")
        response = requests.get(conv_url, headers=headers, timeout=30)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        
        chapters = []
        
        # Находим все h2 с классом text-center (заголовки глав)
        headings = soup.find_all('h2', class_='text-center')
        print(f"    Найдено заголовков h2: {len(headings)}")
        
        for i, heading in enumerate(headings):
            chapter_title = heading.get_text(strip=True).replace('\n', ' ').strip()
            print(f"    Заголовок {i+1}: {chapter_title[:60]}...")
            
            # Ищем следующий div (без класса) после заголовка
            # В структуре: <h2>...</h2> -> <div> -> <p class="txt">...</p>
            next_div = heading.find_next_sibling('div')
            
            chapter_text = []
            if next_div:
                # Ищем внутри div все p с классом txt
                paragraphs = next_div.find_all('p', class_='txt')
                if paragraphs:
                    for p in paragraphs:
                        text = p.get_text(strip=True)
                        if text:
                            chapter_text.append(text)
                else:
                    # Если нет p, берем текст из div
                    text = next_div.get_text(strip=True)
                    if text:
                        chapter_text.append(text)
            
            if chapter_text:
                chapters.append({
                    'title': chapter_title,
                    'content': chapter_text
                })
                print(f"      Добавлена глава, абзацев: {len(chapter_text)}")
            else:
                print(f"      Нет текста для этой главы")
        
        return chapters
        
    except Exception as e:
        print(f"    Ошибка при загрузке {conv_url}: {e}")
        return []

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
        print(f'Старый файл удалён')

# ========== СОЗДАЁМ ДОКУМЕНТ ==========
print('\nСоздаю документ...')
doc = Document()

# ========== НАСТРОЙКА ПОЛЕЙ ==========
section = doc.sections[0]
section.top_margin = Inches(1)
section.bottom_margin = Inches(1)
section.left_margin = Inches(1)
section.right_margin = Inches(1)

# ========== ЗАГОЛОВОК ДОКУМЕНТА ==========
main_title = doc.add_heading('Писания к десяти', level=1)
main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
main_title.paragraph_format.page_break_before = False

for run in main_title.runs:
    run.font.size = Pt(24)
    run.font.name = 'Arial Narrow'
    run.font.bold = True
    run.font.color.rgb = RGBColor(0, 0, 0)

# Подзаголовок
subtitle = doc.add_heading('Собеседования преподобного Иоанна Кассиана Римлянина', level=2)
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitle.paragraph_format.page_break_before = False

for run in subtitle.runs:
    run.font.size = Pt(14)
    run.font.name = 'Arial Narrow'
    run.font.italic = True
    run.font.color.rgb = RGBColor(0, 0, 0)

doc.add_paragraph()

# ========== ПАРСИМ КАЖДОЕ СОБЕСЕДОВАНИЕ ==========
total_chapters = 0

for conv in conversations:
    print(f"\nОбрабатываю: {conv['title']}")
    
    # Заголовок собеседования (уровень 1)
    conv_heading = doc.add_heading(conv['title'], level=1)
    conv_heading.paragraph_format.page_break_before = False
    
    for run in conv_heading.runs:
        run.font.size = Pt(18)
        run.font.name = 'Arial Narrow'
        run.font.bold = True
        run.font.color.rgb = RGBColor(0, 0, 0)
    
    # Парсим главы этого собеседования
    chapters = parse_conversation(conv['url'])
    print(f"  Найдено глав с текстом: {len(chapters)}")
    
    # Добавляем каждую главу в документ
    for chapter in chapters:
        # Заголовок главы (уровень 2)
        chapter_heading = doc.add_heading(chapter['title'], level=2)
        chapter_heading.paragraph_format.page_break_before = False
        
        for run in chapter_heading.runs:
            run.font.size = Pt(14)
            run.font.name = 'Arial Narrow'
            run.font.bold = True
            run.font.color.rgb = RGBColor(0, 0, 0)
        
        # Текст главы
        for para_text in chapter['content']:
            paragraph = doc.add_paragraph(para_text)
            paragraph.paragraph_format.first_line_indent = Inches(0.3)
            for run in paragraph.runs:
                run.font.size = Pt(12)
                run.font.name = 'Arial Narrow'
                run.font.color.rgb = RGBColor(0, 0, 0)
        
        total_chapters += 1
    
    time.sleep(0.5)

# ========== ИТОГО ==========
doc.add_paragraph()
footer_paragraph = doc.add_paragraph(f'Всего собеседований: {len(conversations)}, глав: {total_chapters}')
footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER

for run in footer_paragraph.runs:
    run.font.size = Pt(10)
    run.font.name = 'Arial Narrow'
    run.font.italic = True
    run.font.color.rgb = RGBColor(0, 0, 0)

# ========== СОХРАНЯЕМ ==========
doc.save(file_name)

print(f'\nДокумент "{file_name}" успешно создан!')
print(f'Собеседований: {len(conversations)}')
print(f'Всего глав с текстом: {total_chapters}')

# ========== ОТКРЫВАЕМ ==========
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