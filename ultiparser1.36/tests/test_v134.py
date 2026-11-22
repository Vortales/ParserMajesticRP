"""Тесты новой функциональности v1.34:
- Многофункциональное логирование OCR (_ocr_log_event).
- Диагностика причин пропуска или нераспознавания текста (_ocr_image_bytes).
- Запись ocr_debug.log и ocr_events.jsonl в папку images/.
"""
import unittest
from extract36 import SRC, TARGET, def_block

class TestV134OcrLogging(unittest.TestCase):
    def test_version_134(self):
        self.assertIn('PARSER_VERSION = "1.36"', SRC)

    def test_ocr_log_event_exists(self):
        self.assertIn("def _ocr_log_event", SRC)
        self.assertIn("ocr_events.jsonl", SRC)
        self.assertIn("ocr_debug.log", SRC)

    def test_ocr_image_bytes_diag(self):
        self.assertIn("def _ocr_image_bytes(data: bytes, diag: dict = None)", SRC)
        self.assertIn("no_text_found_by_engines", SRC)
        self.assertIn("pillow_not_installed", SRC)

    def test_process_image_block_logs_events(self):
        self.assertIn("_ocr_log_event(images_dir, {", SRC)
        self.assertIn('"status": "download_failed"', SRC)
        self.assertIn('"status": "no_text"', SRC)
        self.assertIn('"status": "success"', SRC)

if __name__ == "__main__":
    unittest.main()
