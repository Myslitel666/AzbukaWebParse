import os
import subprocess
import time
import json
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.section import WD_ORIENTATION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import requests
from bs4 import BeautifulSoup

# ========== ЗАГРУЗКА КОНФИГУРАЦИИ ==========
def load_config():
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

config = load_config()
FONT_NAME = config['font_name']  # Общее название шрифта для всего документа

# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========
def apply_font(run, font_config):
    """Применяет настройки шрифта к run"""
    run.font.name = FONT_NAME
    run.font.size = Pt(font_config['size_pt'])
    run.font.bold = font_config['bold']
    run.font.italic = font_config['italic']
    rgb = font_config['color_rgb']
    run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])

def setup_footer(doc):
    """Настраивает нижний колонтитул с нумерацией страниц"""
    if not config['footer']['enabled']:
        return
    
    section = doc.sections[0]
    section.header.is_linked_to_previous = False
    section.footer.is_linked_to_previous = False
    
    footer = section.footer
    footer_paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_paragraph.clear()
    
    footer_config = config['fonts']['footer']
    
    run_left = footer_paragraph.add_run(config['footer']['left_symbol'])
    apply_font(run_left, footer_config)
    
    run_page = footer_paragraph.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run_page._r.append(fldChar1)
    
    instrText = OxmlElement('w:instrText')
    instrText.text = "PAGE"
    run_page._r.append(instrText)
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run_page._r.append(fldChar2)
    apply_font(run_page, footer_config)
    
    run_right = footer_paragraph.add_run(config['footer']['right_symbol'])
    apply_font(run_right, footer_config)

def setup_styles(doc):
    """Настраивает стили заголовков и стили TOC для оглавления"""
    # Стиль для заголовка 1 уровня (собеседования)
    style_h1 = doc.styles['Heading 1']
    font = style_h1.font
    font.name = FONT_NAME
    font.size = Pt(config['fonts']['conversation']['size_pt'])
    font.bold = config['fonts']['conversation']['bold']
    font.italic = config['fonts']['conversation']['italic']
    rgb = config['fonts']['conversation']['color_rgb']
    font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
    
    # Стиль для заголовка 2 уровня (главы)
    style_h2 = doc.styles['Heading 2']
    font = style_h2.font
    font.name = FONT_NAME
    font.size = Pt(config['fonts']['chapter']['size_pt'])
    font.bold = config['fonts']['chapter']['bold']
    font.italic = config['fonts']['chapter']['italic']
    rgb = config['fonts']['chapter']['color_rgb']
    font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
    
    # Настраиваем стили TOC для оглавления
    text_config = config['fonts']['text']
    
    # TOC 1 - для заголовков уровня 1 в оглавлении
    try:
        style_toc1 = doc.styles['TOC 1']
        font = style_toc1.font
        font.name = FONT_NAME
        font.size = Pt(text_config['size_pt'])
        font.bold = text_config['bold']
        font.italic = text_config['italic']
        rgb = text_config['color_rgb']
        font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
    except:
        pass
    
    # TOC 2 - для заголовков уровня 2 в оглавлении
    try:
        style_toc2 = doc.styles['TOC 2']
        font = style_toc2.font
        font.name = FONT_NAME
        font.size = Pt(text_config['size_pt'])
        font.bold = text_config['bold']
        font.italic = text_config['italic']
        rgb = text_config['color_rgb']
        font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
    except:
        pass

def add_table_of_contents(doc):
    """Добавляет автоматическое оглавление с заметкой"""
    # Обычный абзац "Содержание" (не заголовок)
    toc_paragraph = doc.add_paragraph()
    toc_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    run_title = toc_paragraph.add_run('Содержание')
    run_title.font.size = Pt(12)
    run_title.font.name = FONT_NAME
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0, 0, 0)
    
    # Пустая строка после заголовка
    doc.add_paragraph()
    
    # Поле TOC
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    
    # Начало поля
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)
    
    # Инструкция TOC (включаем заголовки уровней 1 и 2)
    instrText = OxmlElement('w:instrText')
    instrText.text = 'TOC \\o "1-2" \\h \\z \\u'
    run._r.append(instrText)
    
    # Конец поля
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar2)
    
    # Заметка для пользователя
    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_note = note.add_run('(Нажмите Ctrl+A, затем F9 для обновления содержания)')
    run_note.font.size = Pt(10)
    run_note.font.italic = True
    run_note.font.color.rgb = RGBColor(128, 128, 128)

# ========== ПАРСИНГ САЙТА ==========
def get_conversations():
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    }
    
    print("Загружаю оглавление...")
    response = requests.get(config['url'], headers=headers, timeout=30)
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')
    
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
    
    return conversations

def parse_conversation(conv_url):
    """Загружает страницу собеседования и парсит все главы с текстом"""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
    }
    
    try:
        print(f"  Загружаю: {conv_url}")
        response = requests.get(conv_url, headers=headers, timeout=30)
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        
        chapters = []
        headings = soup.find_all('h2', class_='text-center')
        
        for heading in headings:
            chapter_title = heading.get_text(strip=True).replace('\n', ' ').strip()
            next_div = heading.find_next_sibling('div')
            
            chapter_text = []
            if next_div:
                # Ищем все параграфы внутри div
                paragraphs = next_div.find_all('p', class_='txt')
                
                for p in paragraphs:
                    # Заменяем <br> на пробелы
                    for br in p.find_all('br'):
                        br.replace_with(' ')
                    
                    # Получаем текст, вставляя пробелы между элементами
                    # Используем get_text() с разделителем пробелом
                    # Но сначала удаляем все лишние пробелы
                    text = p.get_text()
                    # Заменяем множественные пробелы на один
                    text = ' '.join(text.split())
                    if text:
                        chapter_text.append(text)
            
            if chapter_text:
                chapters.append({
                    'title': chapter_title,
                    'content': chapter_text
                })
        
        return chapters
        
    except Exception as e:
        print(f"    Ошибка при загрузке {conv_url}: {e}")
        return []

# ========== ОСНОВНОЙ КОД ==========
file_name = config['output_file']
if os.path.exists(file_name):
    try:
        os.remove(file_name)
        print(f'Старый файл "{file_name}" удалён')
    except PermissionError:
        print(f'Файл "{file_name}" открыт!')
        print('Пожалуйста, закройте файл в Word и нажмите Enter...')
        input()
        os.remove(file_name)

print('\nСоздаю документ...')
doc = Document()

# Настройка полей
section = doc.sections[0]
section.top_margin = Cm(config['margins']['top_cm'])
section.bottom_margin = Cm(config['margins']['bottom_cm'])
section.left_margin = Cm(config['margins']['left_cm'])
section.right_margin = Cm(config['margins']['right_cm'])

# Настройка колонтитулов
setup_footer(doc)

# ========== НАСТРОЙКА СТИЛЕЙ ==========
setup_styles(doc)

# ========== НАЗВАНИЕ КНИГИ (обычный абзац, не заголовок) ==========
book_title = doc.add_paragraph()
book_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_title = book_title.add_run(config['headers']['main_title'])
apply_font(run_title, config['fonts']['main_title'])

# Авторство (обычный абзац, не заголовок)
author = doc.add_paragraph()
author.alignment = WD_ALIGN_PARAGRAPH.CENTER
run_author = author.add_run(config['headers']['subtitle'])
apply_font(run_author, config['fonts']['subtitle'])

doc.add_paragraph()

# ========== ДОБАВЛЯЕМ ОГЛАВЛЕНИЕ ==========
add_table_of_contents(doc)

# Получаем список собеседований
conversations = get_conversations()
print(f"Найдено собеседований: {len(conversations)}")

# Парсим каждое собеседование
total_chapters = 0

for conv in conversations:
    print(f"\nОбрабатываю: {conv['title']}")
    
    # Заголовок собеседования (уровень 1)
    conv_heading = doc.add_heading(conv['title'], level=1)
    conv_heading.paragraph_format.page_break_before = False
    conv_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Принудительно применяем шрифт к каждому run в заголовке
    for run in conv_heading.runs:
        apply_font(run, config['fonts']['conversation'])
    
    # Парсим главы
    chapters = parse_conversation(conv['url'])
    print(f"  Найдено глав с текстом: {len(chapters)}")
    
    # Добавляем главы
    for chapter in chapters:
        # Заголовок главы (уровень 2)
        chapter_heading = doc.add_heading(chapter['title'], level=2)
        chapter_heading.paragraph_format.page_break_before = False
        chapter_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        # Принудительно применяем шрифт к каждому run в заголовке главы
        for run in chapter_heading.runs:
            apply_font(run, config['fonts']['chapter'])
        
        # Текст главы
        text_config = config['fonts']['text']
        for para_text in chapter['content']:
            paragraph = doc.add_paragraph(para_text)
            paragraph.paragraph_format.first_line_indent = Cm(text_config.get('first_line_indent_cm', 0.76))
            for run in paragraph.runs:
                apply_font(run, text_config)
        
        total_chapters += 1
    
    time.sleep(0.5)

# Сохраняем
doc.save(file_name)

print(f'\nДокумент "{file_name}" успешно создан!')
print(f'Собеседований: {len(conversations)}')
print(f'Всего глав с текстом: {total_chapters}')

# Открываем файл
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