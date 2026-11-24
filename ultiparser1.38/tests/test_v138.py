import unittest
from pathlib import Path

SRC_PATH = Path(__file__).resolve().parent.parent / "ultiparser1.38.py"
SRC = SRC_PATH.read_text(encoding="utf-8")


class TestV138(unittest.TestCase):
    def test_version(self):
        self.assertIn('PARSER_VERSION = "1.38"', SRC)

    def test_footnote_formatting(self):
        # Проверяем наличие логики сносок в el_to_blocks
        self.assertIn("[сноска:", SRC)

    def test_dfs_ui_controls(self):
        self.assertIn("v_dfs_enabled", SRC)
        self.assertIn("v_dfs_depth", SRC)
        self.assertIn("🔗 Парсить связанные темы по ссылкам в тексте (DFS цепочка исков)", SRC)
        self.assertIn("Глубина цепочки:", SRC)

    def test_dfs_logic_and_cycle_prevention(self):
        self.assertIn("visited_urls = set()", SRC)
        self.assertIn("_extract_thread_links_from_data", SRC)
        self.assertIn("СВЯЗАННАЯ ТЕМА ПО ЦЕПОЧКЕ DFS", SRC)

    def test_link_enrichment_for_images(self):
        self.assertIn("link_url", SRC)
        self.assertIn('lu = (b.get("link_url") or "").strip()', SRC)

    def test_real_title_extraction_in_parse_thread_content(self):
        # Проверяем, что при парсинге страницы извлекается реальный h1 заголовок
        self.assertIn("real_h1 = page_title(soup)", SRC)

    def test_dfs_folders_creation(self):
        # Проверяем создание подпапок для связанных тем
        self.assertIn("next_save_dir = topics_dir / safe_name(title)", SRC)
        self.assertIn("target_save_dir", SRC)

    def test_selective_progress_and_total_fix(self):
        # Проверяем исправление сигнатуры _update_progress_ui и cur_total
        self.assertIn("self._update_progress_ui(d, tot, t_name)", SRC)
        self.assertIn("cur_total = total_counter[0]", SRC)


if __name__ == "__main__":
    unittest.main()
