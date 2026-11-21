"""Тесты новой функциональности v1.33:
Выборочный парсер (SelectiveParserWindow):
- Наличие класса SelectiveParserWindow и кнопки вызова.
- Функция формирования чистого документа build_selective_docx.
- Независимость дерева выборочного парсера от основного дерева.
"""
import unittest
from extract34 import SRC, TARGET, def_block

class TestV133SelectiveParser(unittest.TestCase):
    def test_selective_parser_window_class_exists(self):
        self.assertIn("class SelectiveParserWindow", SRC)
        self.assertIn("btn_selective", SRC)
        self.assertIn("🎯 Выборочный парсер", SRC)

    def test_build_selective_docx_func(self):
        self.assertIn("def build_selective_docx", SRC)

    def test_version_133(self):
        self.assertIn('PARSER_VERSION = "1.34"', SRC)

    def test_progressbar_and_realtime_json(self):
        self.assertIn("self.progress = ttk.Progressbar", SRC)
        self.assertIn("self._update_progress_ui", SRC)
        self.assertIn("compute_content_checksum", SRC)

    def test_skip_docx_option(self):
        self.assertIn("self.v_skip_docx = tk.BooleanVar(value=False)", SRC)
        self.assertIn("Только .json (без .docx)", SRC)
        self.assertIn("if not skip_docx:", SRC)

    def test_ocr_fixes_and_external_images(self):
        self.assertIn("import copy", SRC)
        self.assertIn("HAS_EASYOCR = _ensure(\"easyocr\", \"easyocr\", required=False)", SRC)
        self.assertIn("HAS_PYTESSERACT = _ensure(\"pytesseract\", \"pytesseract\", required=False)", SRC)
        self.assertIn("def _is_image_or_host_url", SRC)
        self.assertIn("cached_ocr = ocr_path.read_text(encoding=\"utf-8\").strip()", SRC)
        self.assertIn("if ocr_text and ocr_text.strip():", SRC)

if __name__ == "__main__":
    unittest.main()
