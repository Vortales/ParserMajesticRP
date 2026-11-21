import unittest
from pathlib import Path
from extract35 import SRC

class TestV135SelectiveOcr(unittest.TestCase):
    def test_version_135(self):
        self.assertIn('PARSER_VERSION = "1.35"', SRC)

    def test_selective_window_uses_stage3_image_enrichment(self):
        # Проверяем, что SelectiveParserWindow использует правильные параметры и логику Stage 3
        self.assertIn("class SelectiveParserWindow", SRC)
        # Проверка проверки транспорта перед запуском парсера
        self.assertIn("self.app.http.fast_enabled", SRC)
        self.assertIn("self.app._ensure_fetch_browser()", SRC)
        # Проверка вызова enrich_thread_data_with_images с cookies, base_url, fetcher
        self.assertIn("fetcher=self.app._fetch_image_bytes", SRC)
        self.assertIn("base_url=data.get(\"url\", \"\") or url or FORUM_BASE_URL", SRC)
        # Проверка передачи ocr_min_global в threads_to_lines
        self.assertIn("ocr_min_global=ocr_min_global", SRC)
        # Проверка функции проверки потребности в OCR обогащении
        self.assertIn("def _thread_needs_image_enrichment", SRC)

if __name__ == "__main__":
    unittest.main()
