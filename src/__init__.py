from .config_loader import load_config, config
from .http_client import session, HEADERS
from .docx_helpers import apply_font, setup_footer, setup_styles, add_table_of_contents
from .docx_builder import (
    add_heading_with_footnotes, 
    add_formatted_paragraph, 
    add_notes_section,
    process_footnotes_in_text,
    create_document  # Добавляем create_document
)
from .html_parser import get_conversations, fetch_all

__all__ = [
    'config',
    'load_config',
    'session',
    'HEADERS',
    'apply_font',
    'setup_footer',
    'setup_styles',
    'add_table_of_contents',
    'add_heading_with_footnotes',
    'add_formatted_paragraph',
    'add_notes_section',
    'get_conversations',
    'fetch_all',
    'process_footnotes_in_text',
    'create_document'  # Добавляем в __all__
]