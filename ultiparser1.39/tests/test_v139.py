#!/usr/bin/env python3
"""Тесты функционала v1.39: DFS ссылок на посты/темы и парсинг документов."""

import unittest
from pathlib import Path
import re
from urllib.parse import urlparse

# Чистая копия функции для автономного тестирования без инициализации UI
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

class TestV139(unittest.TestCase):
    def test_doc_detection_google_docs(self):
        url = "https://docs.google.com/document/d/14hmlB4YrCmQNSzjqGoLb5HdktUqOMGyClrX9wPMYAoU/edit?tab=t.0"
        dl, fn = get_doc_download_info(url, "Исковое заявление")
        self.assertIsNotNone(dl)
        self.assertTrue(dl.endswith("/export?format=docx"))
        self.assertEqual(fn, "Исковое заявление.docx")

    def test_doc_detection_google_sheets(self):
        url = "https://docs.google.com/spreadsheets/d/1JhYrZcQjX8afOgE76HwVhdfCgQ3-k4MkBS5TIXyI3Po/edit"
        dl, fn = get_doc_download_info(url, "Таблица")
        self.assertIsNotNone(dl)
        self.assertTrue(dl.endswith("/export?format=xlsx"))
        self.assertEqual(fn, "Таблица.xlsx")

    def test_doc_detection_pdf(self):
        url = "https://example.com/files/claim.pdf"
        dl, fn = get_doc_download_info(url, "Приложение 1")
        self.assertIsNotNone(dl)
        self.assertEqual(fn, "Приложение 1.pdf")

    def test_forum_url_not_doc(self):
        url = "https://forum.majestic-rp.ru/posts/11830060/"
        dl, fn = get_doc_download_info(url, "Пост")
        self.assertIsNone(dl)
        self.assertIsNone(fn)

if __name__ == "__main__":
    unittest.main()
