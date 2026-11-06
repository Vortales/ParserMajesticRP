"""Тесты новой функциональности v1.31:
1. Удаление всех символов '━' (U+2501) из текста и документов.
2. Этап 4: Конвертация .docx -> .pdf и шифрование через encrypt_tool.py.
"""
import io
import re
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

from extract31 import SRC, TARGET, def_block, grab

HERE = Path(__file__).resolve().parent


class TestV131Features(unittest.TestCase):
    def setUp(self):
        # Подготавливаем namespace с функциями v1.31
        self.ns = {}
        fernet_helpers = grab("# ─── v1.31: Чистая реализация AES-128", "def get_encrypt_tool_function")
        code = (
            "import os\nimport time\nimport re\nimport subprocess\nimport zipfile\nimport xml.etree.ElementTree as ET\nimport sys\n"
            "from pathlib import Path\n"
            "def get_logger():\n"
            "    import logging\n"
            "    return logging.getLogger('test')\n"
            + def_block("strip_box_drawing_chars")
            + "\n"
            + def_block("normalize_text")
            + "\n"
            + def_block("extract_text_from_docx")
            + "\n"
            + def_block("text_to_pdf_bytes")
            + "\n"
            + def_block("convert_docx_to_pdf")
            + "\n"
            + fernet_helpers
            + "\n"
            + def_block("get_encrypt_tool_function")
            + "\n"
            + def_block("encrypt_file_with_tool")
        )
        exec(code, self.ns)

    def test_strip_box_drawing_chars(self):
        strip_fn = self.ns["strip_box_drawing_chars"]
        norm_fn = self.ns["normalize_text"]

        raw = "Раздел 1 ━━━━━━ Общие положения ━━━━━"
        cleaned = strip_fn(raw)
        self.assertNotIn("━", cleaned)
        self.assertEqual(cleaned, "Раздел 1  Общие положения ")

        normed = norm_fn(raw)
        self.assertNotIn("━", normed)
        self.assertEqual(normed, "Раздел 1 Общие положения")

    def test_docx_extract_and_pdf_generation(self):
        extract_fn = self.ns["extract_text_from_docx"]
        pdf_fn = self.ns["text_to_pdf_bytes"]
        conv_fn = self.ns["convert_docx_to_pdf"]

        # Создаём минимальный .docx в памяти
        with tempfile.TemporaryDirectory() as tmpdir:
            docx_path = Path(tmpdir) / "test_doc.docx"
            pdf_path = Path(tmpdir) / "test_doc.pdf"

            # Собираем минимальный валидный docx
            doc_xml = (
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
                '<w:body>'
                '<w:p><w:r><w:t>Заголовок документа ━━━━</w:t></w:r></w:p>'
                '<w:p><w:r><w:t>Статья 1.1. Правило сервера.</w:t></w:r></w:p>'
                '</w:body>'
                '</w:document>'
            )
            with zipfile.ZipFile(docx_path, "w") as z:
                z.writestr("word/document.xml", doc_xml.encode("utf-8"))

            # Проверяем извлечение текста
            txt = extract_fn(docx_path)
            self.assertNotIn("━", txt)
            self.assertIn("Заголовок документа", txt)
            self.assertIn("Статья 1.1", txt)

            # Проверяем конвертацию docx -> pdf
            ok = conv_fn(docx_path, pdf_path)
            self.assertTrue(ok)
            self.assertTrue(pdf_path.is_file())

            pdf_data = pdf_path.read_bytes()
            self.assertTrue(pdf_data.startswith(b"%PDF-1.4"))
            self.assertTrue(pdf_data.endswith(b"%%EOF\n"))
            self.assertGreater(len(pdf_data), 500)

    def test_encrypt_tool_integration(self):
        get_enc_fn = self.ns["get_encrypt_tool_function"]
        enc_file_fn = self.ns["encrypt_file_with_tool"]

        enc_fn = get_enc_fn()
        self.assertIsNotNone(enc_fn)

        sample_bytes = b"Hello, encrypted PDF content!"
        encrypted = enc_fn(sample_bytes)
        self.assertIsInstance(encrypted, bytes)
        self.assertNotEqual(encrypted, sample_bytes)
        self.assertGreater(len(encrypted), len(sample_bytes))

        # Тестируем encrypt_file_with_tool
        with tempfile.TemporaryDirectory() as tmpdir:
            src_pdf = Path(tmpdir) / "test.pdf"
            enc_file = Path(tmpdir) / "test.enc"
            src_pdf.write_bytes(b"%PDF-1.4 sample stream data")

            ok = enc_file_fn(src_pdf, enc_file)
            self.assertTrue(ok)
            self.assertTrue(enc_file.is_file())
            self.assertGreater(enc_file.stat().st_size, src_pdf.stat().st_size)

    def test_gui_buttons_presence(self):
        # Проверяем, что в ultiparser1.31.py зарегистрирована кнопка Этапа 4
        self.assertIn('btn_stage4', SRC)
        self.assertIn('Этап 4: PDF и зашифровать', SRC)
        self.assertIn('def _do_stage4_pdf_encrypt', SRC)


if __name__ == "__main__":
    unittest.main()
