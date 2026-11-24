import unittest
from extract38 import SRC

class TestV136UnifiedLogging(unittest.TestCase):
    def test_version_136(self):
        self.assertIn('PARSER_VERSION = "1.38"', SRC)

    def test_unified_telemetry_and_latest_log(self):
        self.assertIn("latest_unified_run.log", SRC)
        self.assertIn("unified_telemetry.jsonl", SRC)
        self.assertIn("def unified_telemetry_event", SRC)

    def test_browser_telemetry_integrated(self):
        self.assertIn("[BROWSER] [NAVIGATE]", SRC)
        self.assertIn("[BROWSER] [NAVIGATE_DONE]", SRC)
        self.assertIn("unified_telemetry_event(\"browser\", \"fetch_start\"", SRC)
        self.assertIn("unified_telemetry_event(\"browser\", \"fetch_done\"", SRC)

    def test_http_and_stage3_telemetry(self):
        self.assertIn("[HTTP] [FAST_OK]", SRC)
        self.assertIn("[HTTP] [BROWSER_OK]", SRC)
        self.assertIn("[STAGE3] [THREAD_FINISH]", SRC)
        self.assertIn("[STAGE3] [SUMMARY]", SRC)

if __name__ == "__main__":
    unittest.main()
