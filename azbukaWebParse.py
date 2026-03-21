import os
import subprocess
import time
import json
import asyncio
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_BREAK
from docx.enum.section import WD_ORIENTATION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import httpx
from bs4 import BeautifulSoup

# ========== ЗАГРУЗКА КОНФИГУРАЦИИ ==========
def load_config():
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

config = load_config()
FONT_NAME = config['font_name']

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
    style_h1 = doc.styles['Heading 1']
    font = style_h1.font
    font.name = FONT_NAME
    font.size = Pt(config['fonts']['conversation']['size_pt'])
    font.bold = config['fonts']['conversation']['bold']
    font.italic = config['fonts']['conversation']['italic']
    rgb = config['fonts']['conversation']['color_rgb']
    font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
    
    style_h2 = doc.styles['Heading 2']
    font = style_h2.font
    font.name = FONT_NAME
    font.size = Pt(config['fonts']['chapter']['size_pt'])
    font.bold = config['fonts']['chapter']['bold']
    font.italic = config['fonts']['chapter']['italic']
    rgb = config['fonts']['chapter']['color_rgb']
    font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])
    
    text_config = config['fonts']['text']
    
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
    toc_paragraph = doc.add_paragraph()
    toc_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    run_title = toc_paragraph.add_run('Содержание')
    run_title.font.size = Pt(12)
    run_title.font.name = FONT_NAME
    run_title.font.bold = True
    run_title.font.color.rgb = RGBColor(0, 0, 0)
    
    doc.add_paragraph()
    
    paragraph = doc.add_paragraph()
    run = paragraph.add_run()
    
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)
    
    instrText = OxmlElement('w:instrText')
    instrText.text = 'TOC \\o "1-2" \\h \\z \\u'
    run._r.append(instrText)
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar2)
    
    note = doc.add_paragraph()
    note.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run_note = note.add_run('(Нажмите Ctrl+A, затем F9 для обновления содержания)')
    run_note.font.size = Pt(10)
    run_note.font.italic = True
    run_note.font.color.rgb = RGBColor(128, 128, 128)
    
    doc.add_page_break()

# ========== ФУНКЦИЯ ДЛЯ ФОРМАТИРОВАННОГО ТЕКСТА ==========
def add_formatted_paragraph(doc, p_element, text_config):
    """Добавляет параграф в документ с сохранением форматирования (жирный/курсив)"""
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.first_line_indent = Cm(text_config.get('first_line_indent_cm', 0.76))
    
    # Устанавливаем выравнивание из конфига
    alignment = text_config.get('alignment', 'justify')
    if alignment == 'justify':
        paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    elif alignment == 'center':
        paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    elif alignment == 'left':
        paragraph.alignment = WD_ALIGN_PARAGRAPH.LEFT
    elif alignment == 'right':
        paragraph.alignment = WD_ALIGN_PARAGRAPH.RIGHT
    
    # Обрабатываем все дочерние элементы
    for child in p_element.children:
        if isinstance(child, str):
            # Обычный текст - добавляем как есть, сохраняя пробелы
            if child:
                run = paragraph.add_run(child)
                apply_font(run, text_config)
        elif child.name == 'b':
            # Жирный текст
            text = child.get_text()
            if text:
                run = paragraph.add_run(text)
                apply_font(run, text_config)
                run.font.bold = True
        elif child.name == 'span':
            # Проверяем класс для цитат
            classes = child.get('class', [])
            if 'quote' in classes or 'synodal' in classes:
                # Цитата курсивом
                text = child.get_text()
                if text:
                    run = paragraph.add_run(text)
                    apply_font(run, text_config)
                    run.font.italic = True
            else:
                # Обычный span
                text = child.get_text()
                if text:
                    run = paragraph.add_run(text)
                    apply_font(run, text_config)
        elif child.name == 'a':
            # Ссылки (берём только текст)
            text = child.get_text()
            if text:
                run = paragraph.add_run(text)
                apply_font(run, text_config)
        elif child.name == 'br':
            paragraph.add_run('\n')
        else:
            # Для любых других тегов - берём текст
            text = child.get_text()
            if text:
                run = paragraph.add_run(text)
                apply_font(run, text_config)

# ========== АСИНХРОННЫЙ ПАРСИНГ ==========
async def fetch_conversation(client, conv, text_config, doc):
    """Асинхронно загружает одно собеседование и парсит главы"""
    try:
        print(f"  Загружаю: {conv['title']}")
        response = await client.get(conv['url'])
        response.encoding = 'utf-8'
        soup = BeautifulSoup(response.text, 'html.parser')
        
        chapters = []
        headings = soup.find_all('h2', class_='text-center')
        
        for heading in headings:
            chapter_title = heading.get_text(strip=True).replace('\n', ' ').strip()
            next_div = heading.find_next_sibling('div')
            
            if next_div:
                paragraphs = next_div.find_all('p', class_='txt')
                
                if paragraphs:
                    chapters.append({
                        'title': chapter_title,
                        'paragraphs': paragraphs
                    })
        
        return conv, chapters
        
    except Exception as e:
        print(f"    Ошибка при загрузке {conv['url']}: {e}")
        return conv, []

async def fetch_all_conversations(conversations, text_config, doc):
    """Асинхронно загружает все собеседования"""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    }
    
    async with httpx.AsyncClient(headers=headers, timeout=30.0) as client:
        tasks = [fetch_conversation(client, conv, text_config, doc) for conv in conversations]
        results = await asyncio.gather(*tasks)
        return results

def get_conversations():
    """Получает список собеседований с главной страницы"""
    headers = {
        'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36',
        'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8',
    }
    
    print("Загружаю оглавление...")
    response = httpx.get(config['url'], headers=headers, timeout=30)
    response.encoding = 'utf-8'
    soup = BeautifulSoup(response.text, 'html.parser')
    
    # Получаем базовый URL
    base_url = config['url']
    if base_url.endswith('/'):
        base_url = base_url[:-1]
    
    conversations = []
    for span in soup.find_all('span', class_='h2o'):
        link = span.find_parent('a')
        if link and link.get('href', '').startswith('./'):
            href = link.get('href')
            title = span.get_text(strip=True)
            # Формируем полный URL на основе базового
            full_url = f"{base_url}{href[1:]}"  # href начинается с './', убираем точку
            conversations.append({
                'number': href[2:],
                'title': title,
                'url': full_url
            })
    
    return conversations

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

# Настройка стилей
setup_styles(doc)

# Название книги (обычный абзац, не заголовок)
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

# Добавляем оглавление
add_table_of_contents(doc)

# Получаем список собеседований
conversations = get_conversations()
print(f"Найдено собеседований: {len(conversations)}")

# Асинхронно загружаем все собеседования
print("\nЗагружаю собеседования параллельно...")
text_config = config['fonts']['text']

# Запускаем асинхронную загрузку
results = asyncio.run(fetch_all_conversations(conversations, text_config, doc))

# Добавляем главы в документ
total_chapters = 0

for conv, chapters in results:
    print(f"\nОбрабатываю: {conv['title']}")
    print(f"  Найдено глав с текстом: {len(chapters)}")
    
    # Заголовок собеседования (уровень 1)
    conv_heading = doc.add_heading(conv['title'], level=1)
    conv_heading.paragraph_format.page_break_before = False
    conv_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    for run in conv_heading.runs:
        apply_font(run, config['fonts']['conversation'])
    
    # Добавляем главы
    for chapter in chapters:
        # Заголовок главы (уровень 2)
        chapter_heading = doc.add_heading(chapter['title'], level=2)
        chapter_heading.paragraph_format.page_break_before = False
        chapter_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
        
        for run in chapter_heading.runs:
            apply_font(run, config['fonts']['chapter'])
        
        # Добавляем текст главы с форматированием
        for p in chapter['paragraphs']:
            add_formatted_paragraph(doc, p, text_config)
        
        total_chapters += 1

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