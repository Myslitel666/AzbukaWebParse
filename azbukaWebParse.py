import os
import subprocess
import time
import json
from docx import Document
from docx.shared import Inches, Pt, RGBColor, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.section import WD_ORIENTATION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import requests
from bs4 import BeautifulSoup

# ========== ЗАГРУЗКА КОНФИГУРАЦИИ ==========
def load_config():
    """Загружает настройки из config.json"""
    with open('config.json', 'r', encoding='utf-8') as f:
        return json.load(f)

config = load_config()

# ========== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ==========
def apply_font(run, font_config):
    """Применяет настройки шрифта к run"""
    run.font.name = font_config['name']
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
    
    # Левое тире
    run_left = footer_paragraph.add_run(config['footer']['left_symbol'])
    apply_font(run_left, footer_config)
    
    # Поле номера страницы
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
    
    # Правое тире
    run_right = footer_paragraph.add_run(config['footer']['right_symbol'])
    apply_font(run_right, footer_config)

# ========== ПАРСИНГ САЙТА ==========
def get_conversations():
    """Получает список собеседований с главной страницы"""
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
                paragraphs = next_div.find_all('p', class_='txt')
                if paragraphs:
                    for p in paragraphs:
                        text = p.get_text(strip=True)
                        if text:
                            chapter_text.append(text)
                else:
                    text = next_div.get_text(strip=True)
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
# Проверка и удаление старого файла
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
        print(f'Старый файл удалён')

# Создаём документ
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

# Главный заголовок
main_title = doc.add_heading(config['headers']['main_title'], level=1)
main_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
main_title.paragraph_format.page_break_before = False
for run in main_title.runs:
    apply_font(run, config['fonts']['main_title'])

# Подзаголовок
subtitle = doc.add_heading(config['headers']['subtitle'], level=2)
subtitle.alignment = WD_ALIGN_PARAGRAPH.CENTER
subtitle.paragraph_format.page_break_before = False
for run in subtitle.runs:
    apply_font(run, config['fonts']['subtitle'])

doc.add_paragraph()

# Получаем список собеседований
conversations = get_conversations()
print(f"Найдено собеседований: {len(conversations)}")

# Парсим каждое собеседование
total_chapters = 0

for conv in conversations:
    print(f"\nОбрабатываю: {conv['title']}")
    
    # Заголовок собеседования
    conv_heading = doc.add_heading(conv['title'], level=1)
    conv_heading.paragraph_format.page_break_before = False
    conv_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
    for run in conv_heading.runs:
        apply_font(run, config['fonts']['conversation'])
    
    # Парсим главы
    chapters = parse_conversation(conv['url'])
    print(f"  Найдено глав с текстом: {len(chapters)}")
    
    # Добавляем главы
    for chapter in chapters:
        # Заголовок главы
        chapter_heading = doc.add_heading(chapter['title'], level=2)
        chapter_heading.paragraph_format.page_break_before = False
        chapter_heading.alignment = WD_ALIGN_PARAGRAPH.CENTER
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

# Итоговая информация
doc.add_paragraph()
footer_paragraph = doc.add_paragraph(f'Всего собеседований: {len(conversations)}, глав: {total_chapters}')
footer_paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
for run in footer_paragraph.runs:
    run.font.size = Pt(10)
    run.font.name = 'Arial Narrow'
    run.font.italic = True
    run.font.color.rgb = RGBColor(0, 0, 0)

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