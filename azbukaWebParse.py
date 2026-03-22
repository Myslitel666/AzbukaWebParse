import os
import subprocess
import time
import json
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

def add_formatted_paragraph(doc, p_element, text_config):
    paragraph = doc.add_paragraph()
    paragraph.paragraph_format.first_line_indent = Cm(text_config.get('first_line_indent_cm', 0.76))
    paragraph.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
    
    for child in p_element.children:
        text = child if isinstance(child, str) else child.get_text()
        if not text:
            continue
        
        run = paragraph.add_run(text)
        apply_font(run, text_config)
        
        if getattr(child, 'name', None) == 'b':
            run.bold = True
        
        if getattr(child, 'name', None) == 'span':
            classes = child.get('class', [])
            if 'quote' in classes or 'synodal' in classes:
                run.italic = True

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
            return conv, [], False

        node = h1

        current_chapter = None
        intro_paragraphs = []

        while True:
            node = node.find_next()
            if not node:
                break

            # ===== ВСТРЕТИЛИ H2 =====
            if node.name == 'h2' and 'text-center' in node.get('class', []):
                title = node.get_text(strip=True)

                current_chapter = {
                    'title': title,
                    'paragraphs': []
                }
                chapters.append(current_chapter)
                continue

            # ===== ПАРАГРАФ =====
            if node.name == 'p' and 'txt' in node.get('class', []):
                
                # если ещё не было ни одного h2 → это intro
                if current_chapter is None:
                    intro_paragraphs.append(node)
                else:
                    current_chapter['paragraphs'].append(node)

        # ===== ЕСЛИ БЫЛ INTRO =====
        if intro_paragraphs:
            chapters.insert(0, {
                'title': '',  # без заголовка
                'paragraphs': intro_paragraphs
            })

        return conv, chapters, is_fallback
    
    except Exception as e:
        print(f"Ошибка: {e}")
        return conv, [], False

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

for conv, chapters, is_fallback in results:
    h = doc.add_heading(conv['title'], 1)
    h.alignment = WD_ALIGN_PARAGRAPH.CENTER

    for run in h.runs:
        apply_font(run, config['fonts']['conversation'])

    # ===== ЕСЛИ fallback =====
    if is_fallback:
        for ch in chapters:
            # НЕ создаём h2
            for p in ch['paragraphs']:
                add_formatted_paragraph(doc, p, config['fonts']['text'])
            total += 1

    # ===== ЕСЛИ нормальная структура =====
    else:
        for ch in chapters:
            h2 = doc.add_heading(ch['title'], 2)
            h2.alignment = WD_ALIGN_PARAGRAPH.CENTER
            
            for run in h2.runs:
                apply_font(run, config['fonts']['chapter'])
            
            for p in ch['paragraphs']:
                add_formatted_paragraph(doc, p, config['fonts']['text'])
            
            total += 1

doc.save(file_name)

print(f"\nГотово: {file_name}")
print(f"Глав: {total}")

try:
    os.startfile(file_name)
except:
    pass