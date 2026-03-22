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
        
        # Находим все заголовки h2
        all_h2 = book_div.find_all('h2')
        
        # Отделяем служебные заголовки (От издателей) от основных (text-center)
        service_h2 = []
        chapters_h2 = []
        
        for h2 in all_h2:
            classes = h2.get('class', [])
            if 'title' in classes and 'h2' in classes:
                service_h2.append(h2)
            elif 'text-center' in classes:
                chapters_h2.append(h2)
        
        # 1. Обрабатываем "От издателей"
        if service_h2:
            izdateli = service_h2[0]
            izdateli_chapter = {
                'title': izdateli.get_text(strip=True),
                'element': izdateli,
                'paragraphs': []
            }
            
            node = izdateli.find_next()
            while node:
                if node in chapters_h2:
                    break
                # Пропускаем примечания
                if node.name == 'div' and 'note' in node.get('class', []):
                    parse_note(node, notes)
                    node = node.find_next()
                    continue
                # Собираем ВСЕ параграфы
                if node.name == 'p':
                    text = node.get_text(strip=True)
                    if text and not re.match(r'^[\d\s]+$', text):
                        izdateli_chapter['paragraphs'].append(node)
                node = node.find_next()
            
            if izdateli_chapter['paragraphs']:
                chapters.append(izdateli_chapter)
        
        # 2. Обрабатываем письма
        for i, h2 in enumerate(chapters_h2):
            current_chapter = {
                'title': h2.get_text(strip=True),
                'element': h2,
                'paragraphs': []
            }
            
            # Собираем ВСЕ параграфы, пропуская примечания
            node = h2.find_next()
            while node:
                if node.name == 'h2':
                    break
                
                # Если нашли блок примечаний - пропускаем его целиком
                if node.name == 'div' and 'note' in node.get('class', []):
                    parse_note(node, notes)
                    node = node.find_next()
                    continue
                
                # Собираем ВСЕ параграфы (любые p)
                if node.name == 'p':
                    text = node.get_text(strip=True)
                    if text and not re.match(r'^[\d\s]+$', text):
                        current_chapter['paragraphs'].append(node)
                
                node = node.find_next()
            
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