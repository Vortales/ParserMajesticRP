import unittest
from extract37 import SRC

class TestV135SelectiveOcr(unittest.TestCase):
    def test_version_135(self):
        self.assertIn('PARSER_VERSION = "1.37"', SRC)

    def test_selective_window_uses_stage3_image_enrichment(self):
        self.assertIn("class SelectiveParserWindow", SRC)
        self.assertIn("self.app.http.fast_enabled", SRC)
        self.assertIn("self.app._ensure_fetch_browser()", SRC)
        self.assertIn("fetcher=self.app._fetch_image_bytes", SRC)
        self.assertIn("base_url=data.get(\"url\", \"\") or url or FORUM_BASE_URL", SRC)
        self.assertIn("ocr_min_global=ocr_min_global", SRC)
        self.assertIn("def _thread_needs_image_enrichment", SRC)

    def test_pytesseract_removed(self):
        # Проверяем, что pytesseract полностью удалён из кода парсера
        self.assertNotIn("pytesseract", SRC.lower())

    def test_easyocr_gpu_cpu_fallback(self):
        # Проверяем, что EasyOCR имеет обработку и фолбэк на gpu=False
        self.assertIn("easyocr.Reader([\"ru\", \"en\"], verbose=False, gpu=True)", SRC)
        self.assertIn("easyocr.Reader([\"ru\", \"en\"], verbose=False, gpu=False)", SRC)

if __name__ == "__main__":
    unittest.main()

    def test_cuda_status_message(self):
        self.assertIn("def _check_cuda_status", SRC)
        self.assertIn("CUDA: доступна", SRC)
        self.assertIn("CUDA: не найдена", SRC)
        self.assertIn("[EasyOCR] Модель загружена", SRC)
