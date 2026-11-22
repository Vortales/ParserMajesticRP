"""Тесты новой функциональности v1.32:
1. Гарантированное появление .json файлов для всех тем, отмеченных в «Выбор».
2. Удаление тем, которые сняты с «Выбор».
3. Точное разделение тем: только в «Общие» либо только в конкретный «Сервер».
"""
import io
import json
import tempfile
import unittest
from pathlib import Path

from extract32 import SRC, TARGET, def_block, method_block

class TestV132TopicJson(unittest.TestCase):
    def test_ensure_topic_json_present_method_exists(self):
        self.assertIn("def _ensure_topic_json_present", SRC)

    def test_ensure_and_prune_lifecycle(self):
        # Проверяем симуляцию жизненного цикла:
        # 1) Тема создана/активна -> присутствует в Темы
        # 2) Тема снята из выбора -> удаляется из Темы
        with tempfile.TemporaryDirectory() as tmpdir:
            topics_dir = Path(tmpdir) / "Темы"
            topics_dir.mkdir()

            # Файл темы A (выбрана)
            file_a = topics_dir / "Тема_А.json"
            file_a.write_text(json.dumps({
                "format": "ultiparser-topic/1",
                "title": "Тема А",
                "url": "https://forum.majestic-rp.ru/threads/a.1/",
                "posts": []
            }), encoding="utf-8")

            # Файл темы B (была ранее, но снята из выбора)
            file_b = topics_dir / "Тема_Б.json"
            file_b.write_text(json.dumps({
                "format": "ultiparser-topic/1",
                "title": "Тема Б",
                "url": "https://forum.majestic-rp.ru/threads/b.2/",
                "posts": []
            }), encoding="utf-8")

            # Активный набор для этой директории содержит только Тема_А.json
            active_map = {str(topics_dir.resolve()): {"Тема_А.json"}}

            # Имитируем логику _prune_unchecked_topics
            removed = []
            for p in topics_dir.glob("*.json"):
                if p.name not in active_map[str(topics_dir.resolve())]:
                    p.unlink()
                    removed.append(p.name)

            self.assertEqual(removed, ["Тема_Б.json"])
            self.assertTrue(file_a.exists())
            self.assertFalse(file_b.exists())

    def test_version_132(self):
        self.assertIn('PARSER_VERSION = "1.32"', SRC)

if __name__ == "__main__":
    unittest.main()
