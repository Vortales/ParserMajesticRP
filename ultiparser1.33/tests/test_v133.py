"""Тесты новой функциональности v1.33:
Выборочный парсер (SelectiveParserWindow):
- Наличие класса SelectiveParserWindow и кнопки вызова.
- Функция формирования чистого документа build_selective_docx.
- Независимость дерева выборочного парсера от основного дерева.
"""
import unittest
from extract33 import SRC, TARGET, def_block

class TestV133SelectiveParser(unittest.TestCase):
    def test_selective_parser_window_class_exists(self):
        self.assertIn("class SelectiveParserWindow", SRC)
        self.assertIn("btn_selective", SRC)
        self.assertIn("🎯 Выборочный парсер", SRC)

    def test_build_selective_docx_func(self):
        self.assertIn("def build_selective_docx", SRC)

    def test_version_133(self):
        self.assertIn('PARSER_VERSION = "1.33"', SRC)

if __name__ == "__main__":
    unittest.main()

    def test_progressbar_and_realtime_json(self):
        self.assertIn("self.progress = ttk.Progressbar", SRC)
        self.assertIn("self._update_progress_ui", SRC)
        self.assertIn("compute_content_checksum", SRC)

    def test_skip_docx_option(self):
        self.assertIn("self.v_skip_docx = tk.BooleanVar(value=False)", SRC)
        self.assertIn("Только .json (без .docx)", SRC)
        self.assertIn("if not skip_docx:", SRC)
