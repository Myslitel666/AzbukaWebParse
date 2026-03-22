import re
from concurrent.futures import ThreadPoolExecutor, as_completed

from bs4 import BeautifulSoup

from .http_client import session, HEADERS
from .config_loader import config

def parse_note(node, notes):
    """Парсит примечание"""
    sup_link = node.find('sup')
    note_number = None
    if sup_link:
        sup_text = sup_link.get_text(strip=True)
        match = re.search(r'(\d+)', sup_text)
        if match:
            note_number = match.group(1)
    
    note_p = node.find('p', class_='txt')
    if note_p and note_number:
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

def is_navigation_paragraph(p):
    """Проверяет, является ли параграф навигационным (ссылки на письма, оглавление)"""
    classes = p.get('class', [])
    if 'h2o' in classes or 'h3o' in classes:
        return True
    if p.find('span', class_='h2o') or p.find('span', class_='h3o'):
        return True
    text = p.get_text(strip=True)
    if re.match(r'^[\d\s]+$', text):
        return True
    return False

def is_note_section_element(node):
    """Проверяет, является ли элемент частью раздела примечаний"""
    if node.name == 'p':
        classes = node.get('class', [])
        if 'after-text-vignette' in classes:
            return True
        if 'h2' in classes and node.get_text(strip=True) == 'Примечания':
            return True
    if node.name == 'div' and 'note' in node.get('class', []):
        return True
    return False

def collect_paragraphs_until_next_h2(start_node, chapters_h2, notes):
    """Собирает параграфы от start_node до следующего h2"""
    paragraphs = []
    node = start_node.find_next()
    while node:
        if node.name == 'h2':
            break
        
        # Пропускаем примечания
        if node.name == 'div' and 'note' in node.get('class', []):
            parse_note(node, notes)
            node = node.find_next_sibling()
            continue
        
        # Пропускаем элементы примечаний
        if is_note_section_element(node):
            node = node.find_next()
            continue
        
        # Пропускаем навигационные параграфы
        if node.name == 'p' and is_navigation_paragraph(node):
            node = node.find_next()
            continue
        
        # Собираем обычные параграфы
        if node.name == 'p':
            text = node.get_text(strip=True)
            if text and not re.match(r'^[\d\s]+$', text):
                paragraphs.append(node)
        
        node = node.find_next()
    
    return paragraphs

def is_intro_header(h2):
    """Проверяет, является ли заголовок вступительным (От издателей, Вместо предисловия)"""
    text = h2.get_text(strip=True)
    intro_titles = ['От издателей', 'Вместо предисловия']
    return text in intro_titles

def fetch_single_page_conversation(conv, text_config):
    """Парсит страницу, где все главы на одной странице (письма)"""
    try:
        print(f"  Загружаю: {conv['title'] if conv['title'] else 'одностраничник'}")
        
        r = session.get(conv['url'], headers=HEADERS, timeout=30)
        r.encoding = 'utf-8'
        
        # Удаляем все переносы строк из HTML
        html_content = r.text
        html_content = html_content.replace('\n', ' ')
        html_content = re.sub(r' +', ' ', html_content)
        
        soup = BeautifulSoup(html_content, 'html.parser')
        
        chapters = []
        notes = []
        
        # Находим контейнер с содержимым
        book_div = soup.find('div', class_='book')
        if not book_div:
            return conv, [], False, []
        
        # Находим все заголовки h2 в порядке их следования
        all_h2 = book_div.find_all('h2')
        
        # Проходим по всем заголовкам
        for h2 in all_h2:
            text = h2.get_text(strip=True)
            
            # Пропускаем служебные заголовки
            if text == 'Содержание':
                continue
            
            # Создаём главу
            current_chapter = {
                'title': text,
                'element': h2,
                'paragraphs': []
            }
            
            # Собираем параграфы
            current_chapter['paragraphs'] = collect_paragraphs_until_next_h2(h2, all_h2, notes)
            
            # Добавляем главу
            chapters.append(current_chapter)
        
        return conv, chapters, False, notes
    
    except Exception as e:
        print(f"Ошибка в fetch_single_page_conversation: {e}")
        import traceback
        traceback.print_exc()
        return conv, [], False, []
    
def fetch_conversation(conv, text_config):
    """Парсит страницу (для многостраничников - собеседований)"""
    try:
        print(f"  Загружаю: {conv['title']}")
        
        r = session.get(conv['url'], headers=HEADERS, timeout=30)
        r.encoding = 'utf-8'
        
        # Удаляем все переносы строк из HTML
        html_content = r.text
        html_content = html_content.replace('\n', ' ')
        html_content = re.sub(r' +', ' ', html_content)
        
        soup = BeautifulSoup(html_content, 'html.parser')
        
        chapters = []
        is_fallback = False

        h1 = soup.find('h1')
        if not h1:
            return conv, [], False, []

        node = h1

        current_chapter = None
        intro_paragraphs = []
        notes = []

        while True:
            node = node.find_next()
            if not node:
                break

            if node.name == 'h2' and 'text-center' in node.get('class', []):
                current_chapter = {
                    'title': node.get_text(strip=True),
                    'element': node,
                    'paragraphs': []
                }
                chapters.append(current_chapter)
                continue

            # Пропускаем div с примечаниями - они будут обработаны отдельно
            if node.name == 'div' and 'note' in node.get('class', []):
                parse_note(node, notes)
                continue

            # Собираем только обычные параграфы (не из примечаний)
            if node.name == 'p' and ('txt' in node.get('class', []) or 'h6cc' in node.get('class', [])):
                # Проверяем, не находится ли этот параграф внутри div.note
                parent = node.find_parent('div', class_='note')
                if not parent:  # Если не внутри примечания
                    if current_chapter is None:
                        intro_paragraphs.append(node)
                    else:
                        current_chapter['paragraphs'].append(node)
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
    """Универсальная загрузка: определяет тип сайта и использует нужный парсер"""
    results = [None] * len(conversations)
    
    with ThreadPoolExecutor(max_workers=5) as executor:
        future_to_index = {}
        for i, conv in enumerate(conversations):
            # Если в URL есть якорь (#) или это единственная страница
            if '#' in conv['url'] or conv.get('is_single_page', False):
                future = executor.submit(fetch_single_page_conversation, conv, text_config)
            else:
                future = executor.submit(fetch_conversation, conv, text_config)
            future_to_index[future] = i
        
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
    
    # 1. Пробуем найти ссылки на другие страницы (многостраничник)
    for span in soup.find_all('span', class_='h2o'):
        link = span.find_parent('a')
        if link and link.get('href', '').startswith('./'):
            href = link['href']
            conversations.append({
                'title': span.get_text(strip=True),
                'url': f"{base}{href[1:]}",
                'is_single_page': False
            })
    
    # 2. Если нашли ссылки - это многостраничник
    if conversations:
        print(f"Найдено {len(conversations)} страниц (многостраничник)")
        return conversations
    
    # 3. Если ссылок нет - проверяем одностраничник
    book_div = soup.find('div', class_='book')
    if book_div:
        # Считаем количество h2 с классом text-center
        chapters_count = len(book_div.find_all('h2', class_='text-center'))
        
        if chapters_count > 0:
            print(f"Найдено {chapters_count} глав на одной странице (одностраничник)")
            # Получаем заголовок книги
            h1 = book_div.find('h1')
            title = h1.get_text(strip=True) if h1 else config.get('book_title', '')
            
            conversations.append({
                'title': title,
                'url': base,
                'is_single_page': True  # флаг для использования другого парсера
            })
            return conversations
    
    # 4. Fallback - если ничего не нашли
    print("Не удалось определить структуру, используется fallback")
    conversations.append({
        'title': '',  # Пустой заголовок
        'url': base,
        'is_single_page': True  # ← этот флаг
    })
    
    return conversations