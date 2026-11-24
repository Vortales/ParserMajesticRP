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
        self.assertIn("def _extract_thread_links_from_data", SRC)
        self.assertIn("СВЯЗАННАЯ ТЕМА ПО ЦЕПОЧКЕ DFS", SRC)

    def test_embedded_hyperlink_extraction_in_dfs(self):
        # Проверяем, что в DFS парсятся как открытые URL, так и вшитые в текст гиперссылки
        self.assertIn("raw_thread_links = []", SRC)
        self.assertIn("footnotes = re.findall", SRC)
        self.assertIn('b.get("link_url")', SRC)

    def test_inline_hyperlink_preservation_in_blocks(self):
        # Проверяем корректное встраивание гиперссылок внутрь текста фраз без разрыва строк
        self.assertIn("walk_inline", SRC)
        self.assertIn("extract_single_image_block", SRC)
        self.assertIn("[сноска: {href}]", SRC)

    def test_dfs_folders_creation(self):
        # Проверяем создание подпапок для связанных тем
        self.assertIn("next_save_dir = topics_dir / safe_name(title)", SRC)
        self.assertIn("target_save_dir", SRC)

    def test_selective_progress_and_total_fix(self):
        # Проверяем исправление сигнатуры _update_progress_ui и cur_total
        self.assertIn("self._update_progress_ui(d, tot, t_name)", SRC)
        self.assertIn("cur_total = total_counter[0]", SRC)

    def test_selective_always_direct_parse(self):
        # Проверяем, что выборочный парсер всегда парсит напрямую с форума
        self.assertIn("Прямая загрузка с форума темы:", SRC)
        # Убеждаемся, что в SelectiveParserWindow нет вызова load_thread_cache
        pos_spw = SRC.find("class SelectiveParserWindow")
        pos_end_spw = SRC.find("class TopicFilterWindow", pos_spw)
        spw_body = SRC[pos_spw:pos_end_spw]
        self.assertNotIn("load_thread_cache", spw_body)


if __name__ == "__main__":
    unittest.main()
