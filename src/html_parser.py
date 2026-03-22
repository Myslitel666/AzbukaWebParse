import re
from concurrent.futures import ThreadPoolExecutor, as_completed

from bs4 import BeautifulSoup

from .http_client import session, HEADERS
from .config_loader import config

def fetch_conversation(conv, text_config):
    try:
        print(f"  Загружаю: {conv['title']}")
        
        r = session.get(conv['url'], headers=HEADERS, timeout=30)
        r.encoding = 'utf-8'
        
        # Удаляем все переносы строк из HTML
        html_content = r.text
        # Заменяем переносы строк на пробелы
        html_content = html_content.replace('\n', ' ')
        # Убираем множественные пробелы
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

            if node.name == 'p' and 'txt' in node.get('class', []):
                if current_chapter is None:
                    intro_paragraphs.append(node)
                else:
                    current_chapter['paragraphs'].append(node)
                continue
            
            if node.name == 'div' and 'note' in node.get('class', []):
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