import os
import json
import re
from concurrent.futures import ThreadPoolExecutor, as_completed

from docx import Document
from docx.shared import Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

import requests
from bs4 import BeautifulSoup

# ========== ЗАГРУЗКА КОНФИГУРАЦИИ ==========
def load_config():
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

config = load_config()
FONT_NAME = config['font_name']

# ========== HTTP ==========
session = requests.Session()

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/122.0.0.0 Safari/537.36",
    "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
    "Connection": "keep-alive",
}

# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========
def apply_font(run, font_config):
    run.font.name = FONT_NAME
    run.font.size = Pt(font_config['size_pt'])
    run.font.bold = font_config['bold']
    run.font.italic = font_config['italic']
    rgb = font_config['color_rgb']
    run.font.color.rgb = RGBColor(rgb[0], rgb[1], rgb[2])

def setup_footer(doc):
    if not config['footer']['enabled']:
        return
    
    section = doc.sections[0]
    section.footer.is_linked_to_previous = False
    
    footer = section.footer
    p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    p.clear()
    
    cfg = config['fonts']['footer']
    
    run = p.add_run(config['footer']['left_symbol'])
    apply_font(run, cfg)
    
    run_page = p.add_run()
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run_page._r.append(fldChar1)
    
    instrText = OxmlElement('w:instrText')
    instrText.text = "PAGE"
    run_page._r.append(instrText)
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run_page._r.append(fldChar2)
    
    apply_font(run_page, cfg)
    
    run = p.add_run(config['footer']['right_symbol'])
    apply_font(run, cfg)

def setup_styles(doc):
    h1 = doc.styles['Heading 1'].font
    h1.name = FONT_NAME
    h1.size = Pt(config['fonts']['conversation']['size_pt'])
    h1.bold = config['fonts']['conversation']['bold']
    h1.italic = config['fonts']['conversation']['italic']
    h1.color.rgb = RGBColor(*config['fonts']['conversation']['color_rgb'])
    
    h2 = doc.styles['Heading 2'].font
    h2.name = FONT_NAME
    h2.size = Pt(config['fonts']['chapter']['size_pt'])
    h2.bold = config['fonts']['chapter']['bold']
    h2.italic = config['fonts']['chapter']['italic']
    h2.color.rgb = RGBColor(*config['fonts']['chapter']['color_rgb'])

def add_table_of_contents(doc):
    p = doc.add_paragraph()
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    run = p.add_run('Содержание')
    run.font.size = Pt(12)
    run.font.name = FONT_NAME
    run.font.bold = True
    
    doc.add_paragraph()
    
    p = doc.add_paragraph()
    run = p.add_run()
    
    fldChar1 = OxmlElement('w:fldChar')
    fldChar1.set(qn('w:fldCharType'), 'begin')
    run._r.append(fldChar1)
    
    instrText = OxmlElement('w:instrText')
    instrText.text = 'TOC \\o "1-2" \\h \\z \\u'
    run._r.append(instrText)
    
    fldChar2 = OxmlElement('w:fldChar')
    fldChar2.set(qn('w:fldCharType'), 'end')
    run._r.append(fldChar2)
    
    doc.add_page_break()

def process_footnotes_in_text(element, text_config):
    """Обрабатывает элемент и возвращает список фрагментов с форматированием"""
    fragments = []
    
    def process_node(node, default_format='normal'):
        if isinstance(node, str):
            if node.strip():
                fragments.append((node, default_format, None))
            return
        
        # Определяем форматирование для текущего узла
        current_format = default_format
        
        # Обработка ссылки на сноску
        if node.name == 'a' and node.get('href', '').startswith('#note'):
            note_text = node.get_text(strip=True)
            match = re.search(r'(\d+)', note_text)
            if match:
                note_number = match.group(1)
                fragments.append((note_number, 'superscript', None))
            return
        
        # Обработка sup тега
        if node.name == 'sup':
            for child in node.children:
                if child.name == 'a' and child.get('href', '').startswith('#note'):
                    note_text = child.get_text(strip=True)
                    match = re.search(r'(\d+)', note_text)
                    if match:
                        note_number = match.group(1)
                        fragments.append((note_number, 'superscript', None))
                else:
                    text = child if isinstance(child, str) else child.get_text()
                    if text:
                        fragments.append((text, 'superscript', None))
            return
        
        # Обработка span с цитатами - ТОЛЬКО ЭТИ СПАНЫ ДЕЛАЕМ КУРСИВОМ
        if node.name == 'span':
            classes = node.get('class', [])
            if 'quote' in classes or 'synodal' in classes:
                current_format = 'italic'
            # Церковнославянские цитаты тоже курсивом
            if 'church' in classes:
                current_format = 'italic'
        
        # Обработка жирного текста
        if node.name == 'b':
            current_format = 'bold'
        
        # Рекурсивно обрабатываем всех детей
        for child in node.children:
            process_node(child, current_format)
    
    process_node(element)
    return fragments

def add_heading_with_footnotes(doc, element, heading_level, font_config):
    """Добавляет заголовок с возможными сносками"""
    if heading_level == 1:
        heading = doc.add_heading(level=1)
    else:
        heading = doc.add_heading(level=2)
    
    heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    fragments = process_footnotes_in_text(element, font_config)
    
    for text, fmt, note_num in fragments:
        if not text:
            continue
        
        run = heading.add_run(text)
        apply_font(run, font_config)
        
        if fmt == 'bold':
            run.bold = True
        elif fmt == 'italic':
            run.italic = True
        elif fmt == 'superscript':
            run.font.superscript = True

def add_formatted_paragraph(doc, p_element, text_config):
    """Добавляет параграф с обработкой сносок"""
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.first_line_indent = Cm(text_config.get('first_line_indent_cm', 0.76))
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    fragments = process_footnotes_in_text(p_element, text_config)
    
    # Добавляем пробелы между фрагментами там, где нужно
    last_text = ""
    for text, fmt, note_num in fragments:
        if not text:
            continue
        
        # Добавляем пробел между разными span'ами, если их склеило
        if last_text and last_text[-1].isalpha() and text[0].isalpha():
            # Между буквами добавляем пробел
            run = paragraph.add_run(" ")
            apply_font(run, text_config)
        
        run = paragraph.add_run(text)
        apply_font(run, text_config)
        
        if fmt == 'bold':
            run.bold = True
        elif fmt == 'italic':
            run.italic = True
        elif fmt == 'superscript':
            run.font.superscript = True
        
        last_text = text

def add_notes_section(doc, notes):
    """Добавляет раздел с примечаниями"""
    if not notes:
        return
    
    # Разделитель со звёздочками
    p = doc.add_paragraph()
    run = p.add_run('★ ★ ★')
    run.font.size = Pt(config['fonts']['text']['size_pt'])  # Используем обычный размер шрифта
    run.font.name = FONT_NAME
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    
    # Заголовок "Примечания"
    h = doc.add_heading('Примечания', 2)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in h.runs:
        apply_font(run, config['fonts']['chapter'])
    
    # Добавляем каждое примечание
    for note in notes:
        p = doc.add_paragraph()
        p.paragraph_format.first_line_indent = Cm(0)
        p.paragraph_format.left_indent = Cm(0.5)
        
        # Номер сноски (надстрочный)
        run = p.add_run(f"{note['number']}")
        run.font.superscript = True
        apply_font(run, config['fonts']['text'])
        
        # Пробел после номера
        run = p.add_run(" ")
        apply_font(run, config['fonts']['text'])
        
        # Текст примечания
        for fragment in note['fragments']:
            run = p.add_run(fragment['text'])
            apply_font(run, config['fonts']['text'])
            if fragment.get('italic'):
                run.italic = True
            if fragment.get('bold'):
                run.bold = True

# ========== ПАРСИНГ ==========
def fetch_conversation(conv, text_config):
    try:
        print(f"  Загружаю: {conv['title']}")
        
        r = session.get(conv['url'], headers=HEADERS, timeout=30)
        r.encoding = 'utf-8'
        
        soup = BeautifulSoup(r.text, 'html.parser')
        
        chapters = []
        is_fallback = False

        h1 = soup.find('h1')
        if not h1:
            return conv, [], False, []

        node = h1

        current_chapter = None
        intro_paragraphs = []
        notes = []  # Собираем примечания для этой главы

        while True:
            node = node.find_next()
            if not node:
                break

            # ===== ВСТРЕТИЛИ H2 =====
            if node.name == 'h2' and 'text-center' in node.get('class', []):
                # Сохраняем элемент целиком для обработки сносок
                current_chapter = {
                    'title': node.get_text(strip=True),
                    'element': node,
                    'paragraphs': []
                }
                chapters.append(current_chapter)
                continue

            # ===== ПАРАГРАФ =====
            if node.name == 'p' and 'txt' in node.get('class', []):
                if current_chapter is None:
                    intro_paragraphs.append(node)
                else:
                    current_chapter['paragraphs'].append(node)
                continue
            
            # ===== ОБРАБОТКА ПРИМЕЧАНИЙ =====
            # Ищем разделитель * * *
            if node.name == 'p' and 'after-text-vignette' in node.get('class', []):
                # Это разделитель перед примечаниями
                continue
            
            # Ищем заголовок "Примечания"
            if node.name == 'p' and node.get('class') == ['h2'] and node.get_text(strip=True) == 'Примечания':
                continue
            
            # Ищем блоки с примечаниями
            if node.name == 'div' and 'note' in node.get('class', []):
                # Находим номер сноски
                sup_link = node.find('sup')
                note_number = None
                if sup_link:
                    sup_text = sup_link.get_text(strip=True)
                    match = re.search(r'(\d+)', sup_text)
                    if match:
                        note_number = match.group(1)
                
                # Находим текст примечания
                note_p = node.find('p', class_='txt')
                if note_p and note_number:
                    # Обрабатываем текст примечания с возможным форматированием
                    note_fragments = []
                    for child in note_p.children:
                        if isinstance(child, str):
                            if child.strip():
                                note_fragments.append({'text': child, 'italic': False, 'bold': False})
                        elif child.name == 'i' or (child.name == 'span' and 'quote' in child.get('class', [])):
                            note_fragments.append({'text': child.get_text(), 'italic': True, 'bold': False})
                        elif child.name == 'b':
                            note_fragments.append({'text': child.get_text(), 'italic': False, 'bold': True})
                        else:
                            note_fragments.append({'text': child.get_text(), 'italic': False, 'bold': False})
                    
                    notes.append({
                        'number': note_number,
                        'fragments': note_fragments
                    })
                continue

        if intro_paragraphs:
            chapters.insert(0, {
                'title': '',
                'element': None,
                'paragraphs': intro_paragraphs
            })

        return conv, chapters, is_fallback, notes
    
    except Exception as e:
        print(f"Ошибка: {e}")
        return conv, [], False, []

def fetch_all(conversations, text_config):
    results = [None] * len(conversations)
    
    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_index = {
            executor.submit(fetch_conversation, conv, text_config): i
            for i, conv in enumerate(conversations)
        }
        
        for future in as_completed(future_to_index):
            index = future_to_index[future]
            results[index] = future.result()
    
    return results

def get_conversations():
    print("Загружаю оглавление...")
    
    r = session.get(config['url'], headers=HEADERS, timeout=30)
    r.encoding = 'utf-8'
    
    soup = BeautifulSoup(r.text, 'html.parser')
    
    base = config['url'].rstrip('/')
    
    conversations = []
    
    for span in soup.find_all('span', class_='h2o'):
        link = span.find_parent('a')
        if link and link.get('href', '').startswith('./'):
            href = link['href']
            
            conversations.append({
                'title': span.get_text(strip=True),
                'url': f"{base}{href[1:]}"
            })
    
    return conversations

# ========== ОСНОВНОЙ КОД ==========
file_name = config['output_file']

while True:
    if not os.path.exists(file_name):
        break
    try:
        os.remove(file_name)
        print(f'Старый файл "{file_name}" удалён')
        break
    except PermissionError:
        print(f'Файл "{file_name}" открыт!')
        print('Пожалуйста, закройте файл в Word и нажмите Enter...')
        input()

print('\nСоздаю документ...')
doc = Document()

section = doc.sections[0]
section.top_margin = Cm(config['margins']['top_cm'])
section.bottom_margin = Cm(config['margins']['bottom_cm'])
section.left_margin = Cm(config['margins']['left_cm'])
section.right_margin = Cm(config['margins']['right_cm'])

setup_footer(doc)
setup_styles(doc)

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
apply_font(p.add_run(config['headers']['main_title']), config['fonts']['main_title'])

p = doc.add_paragraph()
p.alignment = WD_ALIGN_PARAGRAPH.CENTER
apply_font(p.add_run(config['headers']['subtitle']), config['fonts']['subtitle'])

doc.add_paragraph()
add_table_of_contents(doc)

conversations = get_conversations()
print(f"Найдено: {len(conversations)}")

print("\nЗагрузка...")
results = fetch_all(conversations, config['fonts']['text'])

total = 0
total_notes = 0

for conv, chapters, is_fallback, notes in results:
    # Добавляем заголовок H1
    h = doc.add_heading(conv['title'], 1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in h.runs:
        apply_font(run, config['fonts']['conversation'])

    if is_fallback:
        for ch in chapters:
            for p in ch['paragraphs']:
                add_formatted_paragraph(doc, p, config['fonts']['text'])
            total += 1
    else:
        for ch in chapters:
            # Добавляем заголовок H2 со сносками
            if ch['element']:
                add_heading_with_footnotes(doc, ch['element'], 2, config['fonts']['chapter'])
            
            for p in ch['paragraphs']:
                add_formatted_paragraph(doc, p, config['fonts']['text'])
            
            total += 1
    
    # Добавляем примечания в конец главы
    if notes:
        add_notes_section(doc, notes)
        total_notes += len(notes)

doc.save(file_name)

print(f"\nГотово: {file_name}")
print(f"Глав: {total}")
print(f"Примечаний: {total_notes}")

try:
    os.startfile(file_name)
except:
    pass