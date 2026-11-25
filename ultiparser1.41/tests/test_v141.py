#!/usr/bin/env python3
"""Тесты функционала v1.41: Последовательная очередь запросов форума с паузами и дедупликацией готовых тем."""

import unittest
from pathlib import Path

SRC_PATH = Path(__file__).resolve().parent.parent / "ultiparser1.41.py"
SRC = SRC_PATH.read_text(encoding="utf-8")

class TestV141(unittest.TestCase):
    def test_version_141(self):
        self.assertIn('PARSER_VERSION = "1.41"', SRC)

    def test_queue_and_copy_mechanism(self):
        # Проверяем наличие очереди и проверки существующих файлов
        self.assertIn("thread_queue = collections.deque()", SRC)
        self.assertIn("def _find_existing_topic_file_and_data(norm_u: str):", SRC)
        self.assertIn("shutil.copy2(existing_fp, json_path)", SRC)
        self.assertIn("PAGE_DELAY_MIN, PAGE_DELAY_MAX", SRC)

    def test_no_threadpool_for_forum_threads(self):
        # Проверяем, что многопоточный пул запросов к форуму убран
        self.assertNotIn("with ThreadPoolExecutor(max_workers=max(1, workers_count)) as pool:", SRC)
        self.assertNotIn("list(pool.map(parse_one, items))", SRC)

if __name__ == "__main__":
    unittest.main()
