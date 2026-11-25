import unittest
from extract38 import SRC

class TestV137Optimizations(unittest.TestCase):
    def test_version_137(self):
        self.assertIn('PARSER_VERSION = "1.38"', SRC)

    def test_dead_ttl_and_faster_deadlines(self):
        self.assertIn("IMAGE_DEAD_TTL_DAYS = 5", SRC)
        self.assertIn("IMAGE_DOWNLOAD_DEADLINE = 8", SRC)
        self.assertIn("IMAGE_FALLBACK_TIMEOUT = 4", SRC)

    def test_dead_disk_cache_check(self):
        self.assertIn('.dead"', SRC)
        self.assertIn("IMAGE_DEAD_TTL_DAYS * 86400", SRC)

    def test_ocr_min_area_gate(self):
        self.assertIn("OCR_MIN_AREA = 2500", SRC)
        self.assertIn("(w * h) < OCR_MIN_AREA", SRC)

if __name__ == "__main__":
    unittest.main()
