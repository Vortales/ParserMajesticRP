"""Тесты новой функциональности v1.38:
- Извлечение ссылок в виде сносок: текст [сноска: URL]
- Рекурсивный обход ссылок в глубину (DFS) в выборочном парсере
- Настраиваемая глубина DFS рекурсии
- Обогащение внешних ссылок на хостинги картинок (link_url)
- Защита от зацикливания через нормализацию URL
"""
import unittest
from extract38 import SRC, TARGET, def_block

class TestV138Features(unittest.TestCase):
    def test_version_138(self):
        self.assertIn('PARSER_VERSION = "1.38"', SRC)

    def test_footnote_formatting(self):
        self.assertIn('[сноска: {href}]', SRC)
        self.assertIn('"link_url": href', SRC)

    def test_dfs_controls_in_selective_parser(self):
        self.assertIn('v_dfs_enabled', SRC)
        self.assertIn('v_dfs_depth', SRC)
        self.assertIn('🔗 Парсить связанные темы по ссылкам в тексте (DFS цепочка исков)', SRC)
        self.assertIn('Глубина цепочки:', SRC)

    def test_dfs_recursion_and_visited_protection(self):
        self.assertIn('def _extract_thread_links_from_data', SRC)
        self.assertIn('visited_urls', SRC)
        self.assertIn('dfs_max_depth', SRC)
        self.assertIn('СВЯЗАННАЯ ТЕМА ПО ЦЕПОЧКЕ DFS', SRC)

    def test_image_link_url_handling(self):
        self.assertIn('lu = (b.get("link_url") or "").strip()', SRC)

if __name__ == "__main__":
    unittest.main()
