"""Извлекает тестируемые фрагменты из ultiparser1.24.py и исполняет их с подменами.

Тесты лежат в репозитории (tests/), чтобы переживать сброс песочницы.
Запуск:  cd ultiparser1.24/tests && python3 test_transport.py && python3 test_features.py
"""
import re
from pathlib import Path

HERE = Path(__file__).resolve().parent
TARGET = HERE.parent / "ultiparser1.24.py"
SRC = TARGET.read_bytes().decode("utf-8").replace("\r\n", "\n")


def grab(start_marker, end_marker):
    i = SRC.index(start_marker)
    j = SRC.index(end_marker, i + len(start_marker))
    return SRC[i:j]


def const_line(name):
    m = re.search(rf"^{name}\s+= .*$", SRC, re.M)
    assert m, name
    return m.group(0) + "\n"


def def_block(fname):
    """Тело top-level def fname до следующей строки с колонки 0."""
    i = SRC.index(f"\ndef {fname}(")
    m = re.search(r"\n(?=[^\s])", SRC[i + 1:])
    j = i + 1 + m.start() if m else len(SRC)
    return SRC[i + 1:j] + "\n\n"


def transport_code():
    parts = []
    for c in ("PAGE_DELAY_MIN", "PAGE_DELAY_MAX", "FAST_DELAY_MIN", "FAST_DELAY_MAX",
              "FAST_TIMEOUT", "FAST_MAX_ATTEMPTS", "FAST_BLOCK_THRESHOLD",
              "FAST_BROWSER_ONLY_BASE", "FAST_BROWSER_ONLY_MAX",
              "FAST_MIN_INTERVAL", "FAST_BACKOFF_MAX_MULT",
              "BREAK_STREAK_TRIGGER", "BREAK_RESET_SUCCESSES", "BREAK_WAIT_CHUNK",
              "BROWSER_FALLBACK_THRESHOLD", "OCR_IMAGE_MAX_BYTES",
              "FORUM_BASE_URL", "FORUM_HOST", "FAST_PATH_ENABLED_DEFAULT"):
        parts.append(const_line(c))
    for tup in ("THREAD_POST_MARKERS", "THREAD_ERROR_MARKERS", "BLOCK_MARKERS",
              "LOGGED_IN_MARKERS", "STRONG_LOGIN_MARKERS", "BREAK_LEVELS"):
        m = re.search(rf"{tup} = \(.*?\)\n", SRC, re.S)
        assert m, tup
        parts.append(m.group(0) + "\n")
    parts.append(def_block("html_ok"))
    parts.append(def_block("_login_markers_found"))
    parts.append(def_block("_strong_login_markers_found"))
    parts.append(def_block("_page_has_expected_content"))
    parts.append(def_block("analyze_login_refusal"))
    parts.append(def_block("is_logged_in_html"))
    i2 = SRC.index("def build_real_ua", SRC.index("def build_real_ua") + 1)
    j2 = SRC.index("class ProcessEmitter", i2)
    parts.append(SRC[i2:j2])
    parts.append(grab("class _RateLimiter:", "def base_of("))
    parts.append(def_block("_valid_image_bytes"))
    return "\n".join(parts)


def exclude_code():
    return def_block("normalize_title_for_match") + "\n" + def_block("is_excluded_by_words")


def match_code():
    return def_block("normalize_title_for_match") + "\n" + def_block("match_keyword")


def plural_code():
    return def_block("plural")


def history_code():
    parts = []
    for c in ("HISTORY_FILE", "HISTORY_MAX_EVENTS", "HISTORY_MAX_DIFF_LINES",
              "HISTORY_MAX_LINE_LEN"):
        parts.append(const_line(c))
    parts.append(grab("HISTORY_TS_FMT =", "class Scanner:"))
    parts.append(def_block("norm_thread_url"))
    return "\n".join(parts)


def diff_code():
    parts = [
        def_block("normalize_text"),
        grab("GARBAGE_PATTERNS = [", "def normalize_text"),
        def_block("is_garbage_line"),
        def_block("blocks_to_lines"),
        grab("def _diff_meta_fallback", "def diff_thread_blocks"),
        grab("def diff_thread_blocks", "HISTORY_TS_FMT ="),
        def_block("compute_content_checksum"),
    ]
    return "\n".join(parts)


def ocr_gate_code():
    parts = [const_line("OCR_MIN_BYTES"), const_line("OCR_MIN_DIM")]
    parts.append(def_block("_ocr_worth_running"))
    parts.append(def_block("_ocr_content_key"))
    parts.append(def_block("_ocr_lookup_by_content"))
    parts.append(def_block("_ocr_store_by_content"))
    return "\n".join(parts)


def image_block_code():
    parts = []
    for c in ("IMAGES_SUBDIR", "OCR_MAX_IMAGES_PER_POST", "IMAGE_FALLBACK_TIMEOUT",
              "FORUM_BASE_URL", "FORUM_HOST"):
        parts.append(const_line(c))
    parts.append(def_block("_detect_image_ext"))
    parts.append(ocr_gate_code())
    parts.append(def_block("_process_image_block"))
    parts.append(def_block("_process_blocks_with_images"))
    parts.append(def_block("enrich_thread_data_with_images"))
    return "\n".join(parts)


def emitter_code():
    return grab("class ProcessEmitter:", "class Browser:")


def dates_code():
    parts = [
        def_block("thread_dates_path"),
        def_block("load_thread_dates"),
        grab("_thread_dates_lock = threading.Lock()", "def load_thread_dates_cached"),
        grab("def load_thread_dates_cached", "def evaluate_thread_keyword_match"),
    ]
    return "\n".join(parts)


def date_fn_code():
    i = SRC.index("_TIME_DT_RE = re.compile(")
    j = SRC.index("def thread_dates_path")
    return SRC[i:j]


def docx_code():
    parts = []
    for c in ("DOC_FONT", "DOC_SIZE", "DOC_LS", "DOC_MARGIN"):
        parts.append(const_line(c))
    parts.append(def_block("normalize_text"))
    parts.append(grab("def _remove_rsid_attrs", "def build_server_docx"))
    parts.append(grab("def build_server_docx", "class Project"))
    return "\n".join(parts)
