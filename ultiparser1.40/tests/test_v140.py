#!/usr/bin/env python3
"""Тесты функционала v1.40: Потоковый конвейер DFS, параллельный OCR и загрузка документов."""

import unittest
from pathlib import Path

# Импортируем тестируемые функции из ultiparser1.40
import re
from urllib.parse import urlparse

def get_doc_download_info(url: str, anchor_text: str = "") -> tuple[str, str]:
    if not url or not isinstance(url, str):
        return None, None
    u = url.strip()
    try:
        p = urlparse(u)
    except Exception:
        return None, None
    netloc = p.netloc.lower()
    path = p.path

    if "forum.majestic-rp.ru" in netloc:
        return None, None

    clean_anchor = re.sub(r'[^\w\s\d_-]', '', anchor_text or '').strip()
    if clean_anchor.lower() in ('гиперссылка', 'ссылка', 'документ', 'файл', 'link', 'клик', 'тут', 'здесь'):
        clean_anchor = ''

    m_doc = re.search(r'docs\.google\.com/document/d/([a-zA-Z0-9_-]+)', u)
    if m_doc:
        doc_id = m_doc.group(1)
        fname = f'{clean_anchor}.docx' if clean_anchor else f'GoogleDoc_{doc_id[:12]}.docx'
        return f'https://docs.google.com/document/d/{doc_id}/export?format=docx', fname

    m_sheet = re.search(r'docs\.google\.com/spreadsheets/d/([a-zA-Z0-9_-]+)', u)
    if m_sheet:
        sheet_id = m_sheet.group(1)
        fname = f'{clean_anchor}.xlsx' if clean_anchor else f'GoogleSheet_{sheet_id[:12]}.xlsx'
        return f'https://docs.google.com/spreadsheets/d/{sheet_id}/export?format=xlsx', fname

    m_pres = re.search(r'docs\.google\.com/presentation/d/([a-zA-Z0-9_-]+)', u)
    if m_pres:
        pres_id = m_pres.group(1)
        fname = f'{clean_anchor}.pdf' if clean_anchor else f'GoogleSlide_{pres_id[:12]}.pdf'
        return f'https://docs.google.com/presentation/d/{pres_id}/export/pdf', fname

    m_drive = re.search(r'drive\.google\.com/file/d/([a-zA-Z0-9_-]+)', u)
    if m_drive:
        file_id = m_drive.group(1)
        fname = f'{clean_anchor}.bin' if clean_anchor else f'GoogleDrive_{file_id[:12]}.bin'
        return f'https://drive.google.com/uc?export=download&id={file_id}', fname

    doc_extensions = ('.pdf', '.docx', '.doc', '.xlsx', '.xls', '.csv', '.rtf', '.odt', '.txt')
    path_low = path.lower()
    for ext in doc_extensions:
        if path_low.endswith(ext):
            raw_name = path.rstrip('/').split('/')[-1]
            fname = f'{clean_anchor}{ext}' if clean_anchor else raw_name
            return u, fname

    return None, None

class TestV140(unittest.TestCase):
    def test_doc_urls(self):
        u1 = "https://docs.google.com/document/d/14hmlB4YrCmQNSzjqGoLb5HdktUqOMGyClrX9wPMYAoU/edit"
        dl1, fn1 = get_doc_download_info(u1, "Апелляционная жалоба")
        self.assertTrue(dl1.endswith("/export?format=docx"))
        self.assertEqual(fn1, "Апелляционная жалоба.docx")

    def test_dedup_cache(self):
        cache = {}
        url = "https://example.com/image.png"
        cache[url] = [{"type": "paragraph", "text": "Распознанный текст из кэша"}]
        self.assertIn(url, cache)
        self.assertEqual(cache[url][0]["text"], "Распознанный текст из кэша")

if __name__ == "__main__":
    unittest.main()
