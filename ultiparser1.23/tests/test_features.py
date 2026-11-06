"""Тесты фич 1.23: plurals, кламп прогресса, история, diff-диагностика,
картинки (дедуп/нумерация/OCR-гейты), docx-паритет, статические стражи."""
import json
import queue
import re
import sys
import tempfile
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from extract23 import (SRC, TARGET, def_block, plural_code, history_code, diff_code,
                       image_block_code, ocr_gate_code, emitter_code, docx_code,
                       dates_code, date_fn_code, exclude_code, match_code)

# ---------- P1: plurals ----------
ns = {}
exec(plural_code(), ns)
plural = ns["plural"]
cases = [(1, "пост", "поста", "постов", "пост"), (2, "пост", "поста", "постов", "поста"),
         (5, "пост", "поста", "постов", "постов"), (11, "пост", "поста", "постов", "постов"),
         (21, "пост", "поста", "постов", "пост"), (22, "пост", "поста", "постов", "поста"),
         (111, "пост", "поста", "постов", "постов"), (0, "файл", "файла", "файлов", "файлов"),
         (1, "тема", "темы", "тем", "тема"), (3, "тема", "темы", "тем", "темы"),
         (14, "раздел", "раздела", "разделов", "разделов"), (101, "узел", "узла", "узлов", "узел")]
for n, a, b, c, want in cases:
    assert plural(n, a, b, c) == want, (n, want)
print("P1 ok: plurals")

# ---------- P2: кламп прогресса ----------
ens = {"queue": queue, "uuid": uuid, "time": __import__("time")}
exec(emitter_code(), ens)
em = ens["ProcessEmitter"]()
pid = em.process_started("X", 4)
em.step_progress(pid, 120, 100)
em.step_progress(pid, 50, 100)
msgs = [em.q.get_nowait() for _ in range(3)]
vals = [m["value"] for m in msgs if m.get("key") == "progress"]
assert vals[0] == "[▓▓▓▓▓▓▓▓▓▓] 100%", vals
assert vals[1] == "[▓▓▓▓▓░░░░░] 50%", vals
print("P2 ok: clamp 120/100 → 100%")

# ---------- H1-H4: история ----------
hns = {"re": re, "json": json, "Path": Path,
       "datetime": __import__("datetime").datetime}
exec(history_code(), hns)
append = hns["history_append"]
U = "https://forum.majestic-rp.ru/threads/ugolovnyj-kodeks.5/"
d = {"topics": {}}
assert append(d, U, "new", title="УК", ts="2026-09-30 10:00:00") is True
assert append(d, U, "new", title="УК", ts="2026-09-30 10:01:00") is False  # дубль
assert append(d, U, "changed", title="УК", checksum="a" * 8,
              diff=["[Пост #1] изменения:", "    БЫЛО:  x", "    СТАЛО: y"],
              ts="2026-10-01 12:00:00") is True
assert append(d, U, "changed", title="УК", checksum="a" * 8,
              diff=["..."], ts="2026-10-01 12:01:00") is False  # тот же checksum
assert append(d, U, "changed", title="УК", checksum="b" * 8, diff=["[Пост #2] изменения:"],
              ts="2026-10-01 13:00:00") is True
t = d["topics"][hns["norm_thread_url"](U)]
assert [e["kind"] for e in t["events"]] == ["new", "changed", "changed"]
assert t["updated"] == "2026-10-01 13:00:00" and t["url"] == U and t["title"] == "УК"
print("H1 ok: append + dedup")

# H2: caps
big = ["ln%d" % i for i in range(250)]
assert append(d, U, "changed", title="УК", checksum="c" * 8, diff=big,
              ts="2026-10-01 14:00:00") is True
stored = d["topics"][hns["norm_thread_url"](U)]["events"][-1]["diff"]
assert len(stored) == 201 and stored[-1].startswith("… обрезано"), len(stored)
assert append(d, U, "changed", title="УК", checksum="d" * 8, diff=["z" * 600],
              ts="2026-10-01 15:00:00") is True
assert d["topics"][hns["norm_thread_url"](U)]["events"][-1]["diff"] == ["z" * 500 + "…"]
for i in range(110):
    append(d, U, "found", title="УК", checksum=f"s{i}",
           ts=f"2026-10-02 00:{i % 60:02d}:00")
evts = d["topics"][hns["norm_thread_url"](U)]["events"]
assert len(evts) == 100 and evts[0]["checksum"] == "s10", (len(evts), evts[0])
print("H2 ok: caps (diff 200, строка 500, событий 100)")

# H3: round-trip
tmp = Path(tempfile.mkdtemp())
assert hns["save_history"](tmp, d) is True
d2 = hns["load_history"](tmp)
assert d2["topics"].keys() == d["topics"].keys()
assert hns["load_history"](tmp / "nope") == {"topics": {}}
(tmp / "history.json").write_text("not json{", encoding="utf-8")
assert hns["load_history"](tmp) == {"topics": {}}
print("H3 ok: save/load + толерантность")

# H4: сегменты + порядок
ev = {"ts": "2026-10-01 12:00:00", "kind": "changed", "section": "Законы",
      "diff": ["[Пост #1, автор: A] изменения:", "    БЫЛО:  было", "    СТАЛО: стало",
               "    ДОБАВЛЕНО: новь", "    УДАЛЕНО: хлам", "прочий текст"]}
segs = hns["history_event_segments"](ev)
tags = [t for _, t in segs]
assert tags == ["h_date", "ctx", "post", "was", "became", "add", "dlt", "ctx"], tags
assert segs[0][0].startswith("♻️ 2026-10-01 12:00:00 — Изменилась"), segs[0]
assert hns["history_event_segments"]({"ts": "t", "kind": "gone"})[0][1] == "h_date"
topic = {"events": [{"ts": "2026-09-30"}, {"ts": "2026-10-01"}]}
assert [e["ts"] for e in hns["history_sorted_events"](topic, True)] == ["2026-10-01", "2026-09-30"]
assert [e["ts"] for e in hns["history_sorted_events"](topic, False)] == ["2026-09-30", "2026-10-01"]
print("H4 ok: сегменты/теги/порядок")

# H5: scopes слоя (general/server) в событии и в теме
d5 = {"topics": {}}
assert append(d5, U, "changed", title="УК", checksum="e" * 8,
              scopes=["general", "server:Memphis"]) is True
t5 = d5["topics"][hns["norm_thread_url"](U)]
assert t5["scopes"] == ["general", "server:Memphis"]
assert t5["events"][-1]["scopes"] == ["general", "server:Memphis"]
# дубль по checksum событие не пишет, но scopes мёржит
assert append(d5, U, "changed", title="УК", checksum="e" * 8,
              scopes=["server:Miami"]) is False
assert t5["scopes"] == ["general", "server:Memphis", "server:Miami"]
segs5 = hns["history_event_segments"](t5["events"][-1])
assert ("    Слой: Общие правила; Сервер «Memphis»", "ctx") in segs5, segs5
assert hns["history_event_segments"]({"ts": "t", "kind": "new", "scopes": []})[0][1] == "h_date"
print("H5 ok: scopes + строка слоя")

# H6: слои истории (общие + сервер, чужие скрыты)
lns = {}
exec(def_block("history_layers"), lns)
hl = lns["history_layers"]
assert hl([], "Miami") == {"general"}
assert hl(None, "Miami") == {"general"}
assert hl(["general"], "Miami") == {"general"}
assert hl(["server:Miami"], "Miami") == {"server"}
assert hl(["server:Miami"], "Leo") == set()
assert hl(["server:Miami"], "") == set()
assert hl(["general", "server:Miami"], "Miami") == {"general", "server"}
assert hl(["general", "server:Miami"], "Leo") == {"general"}
print("H6 ok: history_layers")

# E1: слова-исключения (тема + раздел)
xns = {"re": re}
exec(exclude_code(), xns)
fex = xns["is_excluded_by_words"]
assert fex("Закон о полиции", "Раздел", ["отмен"]) is False
assert fex("Закон о полиции", "Отмененные законопроекты", ["отмен"]) is True
assert fex("Обращение о даче разъяснений", "", ["обращение"]) is True
assert fex("Обращение о даче разъяснений", "", ["ОБРАЩЕНИЕ"]) is True
assert fex("Уголовный кодекс", "", ["обращение"]) is False
assert fex("Что-то", "Где-то", []) is False
assert fex("Что-то", "", None) is False
assert fex("Ёжик в тумане", "", ["ежик"]) is True
assert fex("Тема", "Раздел", [{"word": "раздел"}]) is True
assert fex("Конституционный закон", "Отмененные законопроекты", ["Отмененные законопроекты"]) is True
assert fex("Конституционный закон", "Отмененные законопроекты", ["отмененные"]) is True
assert fex("Конституционный закон", "Действующие законы", ["Отмененные законопроекты"]) is False
print("E1 ok: исключения")

# E2: матчер ключевых слов (целое слово + позиция)
mns = {"re": re}
exec(match_code(), mns)
mk = mns["match_keyword"]
# строки — как раньше (подстрока)
assert mk("Законодательная база", "закон") is True
assert mk("Уголовный кодекс", "закон") is False
# целое слово
assert mk("Закон О полиции", {"keyword": "закон", "whole": True}) is True
assert mk("Законодательная база", {"keyword": "закон", "whole": True}) is False
assert mk("ЗАКОН", {"keyword": "закон", "whole": True}) is True
assert mk("Тут закон есть", {"keyword": "закон", "whole": True}) is True
assert mk("Тут законы есть", {"keyword": "закон", "whole": True}) is False
# позиция
assert mk("Закон О полиции", {"keyword": "закон", "pos": "start"}) is True
assert mk("Тут закон", {"keyword": "закон", "pos": "start"}) is False
assert mk("Тут закон", {"keyword": "закон", "pos": "end"}) is True
assert mk("Закон тут", {"keyword": "закон", "pos": "end"}) is False
assert mk("Закон", {"keyword": "закон", "pos": "end"}) is True
# целое + начало
assert mk("Закон О полиции", {"keyword": "закон", "whole": True, "pos": "start"}) is True
assert mk("Законодательство", {"keyword": "закон", "whole": True, "pos": "start"}) is False
assert mk("Новый закон", {"keyword": "закон", "whole": True, "pos": "start"}) is False
# фраза целиком
assert mk("Закон о полиции", {"keyword": "закон о", "whole": True}) is True
assert mk("Законопроект о полиции", {"keyword": "закон о", "whole": True}) is False
# регистр/ё/пустые/мусор
assert mk("Ёжик", {"keyword": "ежик", "whole": True}) is True
assert mk("", "закон") is False and mk("Закон", "") is False
assert mk("Закон", {}) is False and mk("Закон", {"keyword": "закон", "pos": "wtf"}) is True
print("E2 ok: match_keyword")

# ---------- D1-D2: diff-диагностика ----------
dns = {"re": re, "difflib": __import__("difflib"), "json": json,
       "hashlib": __import__("hashlib")}
exec(diff_code(), dns)
diff_fn = dns["diff_thread_blocks"]


def post(author="A", date="01.01.2026", blocks=None):
    return {"author": author, "date": date,
            "blocks": blocks if blocks is not None else [{"type": "paragraph", "text": "норма"}]}


old = {"posts": [post()]}
# только автор
new = {"posts": [post(author="B")]}
dl = diff_fn(old, new)
assert dl and any("автор:" in s for s in dl), dl
# только пробелы
new = {"posts": [post(blocks=[{"type": "paragraph", "text": "  норма   "}])]}
dl = diff_fn(old, new)
assert dl and any("техническое" in s for s in dl), dl
# только src картинки
b1 = [{"type": "image", "src": "https://f/a.png", "ocr_text": ""}]
b2 = [{"type": "image", "src": "https://f/b.png", "ocr_text": ""}]
dl = diff_fn({"posts": [post(blocks=b1)]}, {"posts": [post(blocks=b2)]})
assert dl and any("картинки" in s for s in dl), dl
# идентично → пусто
assert diff_fn(old, {"posts": [post()]}) == []
# настоящий текст → БЫЛО/СТАЛО (регрессия обычного пути)
new = {"posts": [post(blocks=[{"type": "paragraph", "text": "ИЗМЕНЕНО"}])]}
dl = diff_fn(old, new)
assert any("БЫЛО:" in s for s in dl) and any("СТАЛО:" in s for s in dl), dl
# checksum отличается ⟺ diff не пуст
assert dns["compute_content_checksum"](old) != dns["compute_content_checksum"](new)
print("D1-D2 ok: МЕТА-диагностика, 0-изм невозможен")

# ---------- I1-I3: картинки ----------
ons = {"re": re, "hashlib": __import__("hashlib"), "time": __import__("time"),
       "os": __import__("os"), "Path": Path, "urlparse": __import__("urllib.parse", fromlist=["urlparse"]).urlparse}


class FakeEm:
    def __init__(self):
        self.labels = []
        self.states = []

    def step_added(self, pid, label, icon=""):
        self.labels.append(label)
        return f"s{len(self.labels)}"

    def step_state(self, sid, state, detail=""):
        self.states.append((sid, state, detail))

    def step_metric(self, sid, key, val):
        self.states.append((sid, key, val))


calls = {"fetch": 0, "ocr": 0, "dl": 0}
FAKE_IMG = b"\x89PNG\r\n\x1a\n" + b"q" * 2000


def fake_fetch(url, proxy_url=""):
    calls["fetch"] += 1
    return FAKE_IMG + url.encode(), "stub"  # разные байты на URL (иначе сработает контент-кэш)


img_ns = dict(ons)
img_ns["_download_image_bytes"] = lambda *a, **k: (calls.__setitem__("dl", calls["dl"] + 1), None)[1]
img_ns["_ocr_image_bytes"] = lambda raw: (calls.__setitem__("ocr", calls["ocr"] + 1), "РАСПОЗНАНО")[1]
exec(image_block_code(), img_ns)

IDIR = Path(tempfile.mkdtemp())
em = FakeEm()
data = {"url": "https://forum.majestic-rp.ru/threads/t.1/", "posts": [
    {"blocks": [{"type": "image", "src": "https://cdn/x.png", "alt": ""},
                {"type": "image", "src": "https://cdn/y.png", "alt": ""}]},
    {"blocks": [{"type": "image", "src": "https://cdn/x.png", "alt": ""}]},  # дубль
]}
out = img_ns["enrich_thread_data_with_images"](data, project_dir=IDIR, fetcher=fake_fetch,
                                               emitter=em, parent_id="p")
assert calls["fetch"] == 2, calls  # x скачан 1 раз, y 1 раз
assert calls["ocr"] == 2, calls
img_labels = [l for l in em.labels if l.startswith("🖼 Изображение")]
assert img_labels == ["🖼 Изображение #1", "🖼 Изображение #2", "🖼 Изображение #3"], em.labels
assert any(s[2] == "дубль URL — из памяти" for s in em.states), em.states
assert out["posts"][1]["blocks"][0]["ocr_text"] == "РАСПОЗНАНО"
print("I1 ok: дедуп URL + сквозная нумерация")

# I2: негативный кэш
calls.update(fetch=0, ocr=0, dl=0)
img_ns2 = dict(ons)
img_ns2["_download_image_bytes"] = lambda *a, **k: (calls.__setitem__("dl", calls["dl"] + 1), None)[1]
img_ns2["_ocr_image_bytes"] = lambda raw: (calls.__setitem__("ocr", calls["ocr"] + 1), "T")[1]
exec(image_block_code(), img_ns2)
IDIR2 = Path(tempfile.mkdtemp())
em2 = FakeEm()
data2 = {"url": "https://forum.majestic-rp.ru/threads/t.2/", "posts": [
    {"blocks": [{"type": "image", "src": "https://cdn/dead.png", "alt": ""}]},
    {"blocks": [{"type": "image", "src": "https://cdn/dead.png", "alt": ""}]},
]}
img_ns2["enrich_thread_data_with_images"](
    data2, project_dir=IDIR2,
    fetcher=lambda u, proxy_url="": (calls.__setitem__("fetch", calls["fetch"] + 1), (None, ""))[1],
    emitter=em2, parent_id="p")
assert calls["fetch"] == 1 and calls["dl"] == 1, calls  # второй — из негативного кэша
print("I2 ok: негативный кэш")

# I3: OCR-гейты
gns = dict(ons)
exec(ocr_gate_code(), gns)
assert gns["_ocr_worth_running"](b"x" * 100) is False
assert gns["_ocr_worth_running"](b"") is False
from PIL import Image
tiny = IDIR / "tiny.png"
Image.new("RGB", (10, 10)).save(tiny)
assert gns["_ocr_worth_running"](tiny.read_bytes()) is False
big = IDIR / "big.png"
import random as _rnd
_px = bytes(_rnd.getrandbits(8) for _ in range(200 * 100 * 3))
Image.frombytes("RGB", (200, 100), _px).save(big)
raw_big = big.read_bytes()
assert len(raw_big) >= 1500, len(raw_big)
assert gns["_ocr_worth_running"](raw_big) is True
assert gns["_ocr_lookup_by_content"](raw_big, IDIR) is None
gns["_ocr_store_by_content"](raw_big, IDIR, "HELLO")
assert gns["_ocr_lookup_by_content"](raw_big, IDIR) == "HELLO"
print("I3 ok: гейты размера + кэш по содержимому")

# ---------- X1: docx-паритет ----------
dxx = {"os": __import__("os"), "re": re, "json": json, "time": __import__("time"),
       "threading": __import__("threading"),
       "Document": __import__("docx").Document,
       "Pt": __import__("docx.shared", fromlist=["Pt"]).Pt,
       "WD_ALIGN_PARAGRAPH": __import__("docx.enum.text", fromlist=["WD_ALIGN_PARAGRAPH"]).WD_ALIGN_PARAGRAPH,
       "WD_LINE_SPACING": __import__("docx.enum.text", fromlist=["WD_LINE_SPACING"]).WD_LINE_SPACING,
       "OxmlElement": __import__("docx.oxml", fromlist=["OxmlElement"]).OxmlElement,
       "qn": __import__("docx.oxml.ns", fromlist=["qn"]).qn,
       "RGBColor": __import__("docx.shared", fromlist=["RGBColor"]).RGBColor,
       "BeautifulSoup": __import__("bs4").BeautifulSoup,
       "logging": __import__("logging"), "Path": Path,
       "urlparse": __import__("urllib.parse", fromlist=["urlparse"]).urlparse,
       "urljoin": __import__("urllib.parse", fromlist=["urljoin"]).urljoin,
       "html": __import__("html")}
exec(docx_code(), dxx)
PROMPT = "Строка1\nСтрока2"
TAIL = "Действую согласно промта в начале файла и в конце ответа всегда отвечай ===КОНЕЦ ОТВЕТА===."
srv = dxx["build_server_docx"](PROMPT, ["П1_тест"], "Srv", ["С1 a=b"])
fp = str(tmp / "srv.docx")
srv.save(fp)
paras = [p.text for p in __import__("docx").Document(fp).paragraphs]
ne = [x for x in paras if x.strip()]
# v1.19: ОДИН абзац, промт первый 1→2→3, хвост-приказ, чистка = и _
assert len(ne) == 1, ne
flat = ne[0]
assert flat.startswith("Строка1 Строка2 "), flat[:60]
assert flat.endswith(TAIL), flat[-140:]
i_prompt = flat.index("Строка1")
i_gen = flat.index("ОБЩИЕ ПРАВИЛА")
i_p1 = flat.index("П1тест")  # "_" вычищен
i_srv = flat.index("ЗАКОНОДАТЕЛЬНАЯ БАЗА: Srv")
i_c1 = flat.index("С1 ab")  # "=" вычищен
i_tail = flat.index("Действую согласно")
assert i_prompt < i_gen < i_p1 < i_srv < i_c1 < i_tail, flat[:200]
assert "  " not in flat and "\n" not in flat and "\t" not in flat
assert flat.count("===КОНЕЦ ОТВЕТА===") == 1  # только хвост, разделители убиты
clean = flat.replace("===КОНЕЦ ОТВЕТА===", "")
assert "=" not in clean and "_" not in clean
print("X1 ok: слои 1→2→3, хвост, чистка =/_")

# X2: общий документ — промт первый, тоже один абзац + хвост + чистка
cmn = dxx["build_common_docx"](PROMPT, ["П1_x=1"])
fp2 = str(tmp / "common.docx")
cmn.save(fp2)
ne2 = [x for x in [p.text for p in __import__("docx").Document(fp2).paragraphs] if x.strip()]
assert len(ne2) == 1 and ne2[0].startswith("Строка1 Строка2 "), ne2
assert ne2[0].index("Строка1") < ne2[0].index("ОБЩИЕ ПРАВИЛА") < ne2[0].index("П1x1")
assert ne2[0].endswith(TAIL)
assert ne2[0].count("===КОНЕЦ ОТВЕТА===") == 1
clean2 = ne2[0].replace("===КОНЕЦ ОТВЕТА===", "")
assert "=" not in clean2 and "_" not in clean2 and "  " not in ne2[0]
# пустой промт — один абзац, начинается с ОБЩИЕ ПРАВИЛА, хвост на месте
cmn0 = dxx["build_common_docx"]("", ["П1"])
fp0 = str(tmp / "common0.docx")
cmn0.save(fp0)
ne0 = [x for x in [p.text for p in __import__("docx").Document(fp0).paragraphs] if x.strip()]
assert len(ne0) == 1 and ne0[0].startswith("ОБЩИЕ ПРАВИЛА П1"), ne0
assert ne0[0].endswith(TAIL) and ne0[0].count("===КОНЕЦ ОТВЕТА===") == 1
print("X2 ok: common-docx, пустой промт")

# ---------- G1-G2: статические стражи ----------
assert SRC.count('new_threads_found.append((node.title, title, t["url"]))') == 1
assert SRC.count('disappeared_selected_threads.append((node.title, title, t["url"]))') == 1
for dead in ("_show_topics_report", "_show_threads_diff_report", "_show_state_report",
             "FAST_BROWSER_ONLY_COOLDOWN", "_scope_visible", "_on_hmode", "v_hmode"):
    assert dead not in SRC, dead
assert SRC.count("btn_history") >= 3
assert re.search(r"^IMAGE_DOWNLOAD_DEADLINE = 12$", SRC, re.M)
assert re.search(r"^IMAGE_FALLBACK_TIMEOUT = 6$", SRC, re.M)
assert "time.monotonic() - t0 < 2" in SRC and "time.monotonic() - t1 < 4" in SRC
assert 'PARSER_VERSION = "1.23"' in SRC
assert "def match_keyword" in SRC and '"whole"' in SRC and '"pos"' in SRC
assert "в начале" in SRC and "в конце" in SRC and 'text="целое"' in SRC
assert "STRONG_LOGIN_MARKERS" in SRC and "def analyze_login_refusal" in SRC
assert "def _page_has_expected_content" in SRC and "fast_trusted" in SRC
assert "доверяю маркерам входа" in SRC and "маркеры входа: " in SRC
assert 'END_ANSWER_MARKER = "===КОНЕЦ ОТВЕТА==="' in SRC
assert "DOC_TAIL_TEXT" in SRC and SRC.count("g.add(DOC_TAIL_TEXT)") == 2
assert SRC.count("g.flatten()") == 2
assert "def verify_fast_session" in SRC and "def _cookie_diag" in SRC
assert "cookies_persister" in SRC and "_persist_harvested_cookies" in SRC
assert "LOGGED_IN_MARKERS" in SRC and "def _login_markers_found" in SRC
assert "def is_excluded_by_words" in SRC and "auto_check_thread_exclude" in SRC
assert SRC.count("is_excluded_by_words(") >= 5
assert "HistoryExcludeWindow" in SRC and "history_exclude_words" in SRC
assert '"workers"' in SRC and "btn_check_cookies" in SRC
assert '"history_view"' in SRC and "def history_layers" in SRC
assert "───── 🖥 Сервер" in SRC and "def _on_hserver" in SRC
assert "def _apply_exclusions_to_checks" in SRC and "def _checked_thread_urls" in SRC
assert "Активные исключения" in SRC and "только отмеченные темы" in SRC
assert "def _apply_struct_filter" in SRC and "v_tsearch" in SRC
assert "исключения бьют и отметку раздела" in SRC
assert "BREAK_STREAK_TRIGGER" in SRC and "BREAK_LEVELS" in SRC
assert "def _note_block_signal" in SRC and "def _wait_breaker" in SRC
assert "def _sleep_abortable" in SRC and "abort_event" in SRC
assert "stealth_mult" in SRC and "pace_mult" in SRC
assert "v_stealth" in SRC and "Тихий режим" in SRC
assert re.search(r"^SAFE_WORKERS_FAST = 2\b", SRC, re.M)
assert "breaker_left" in SRC and "форум отдыхает" in SRC
assert "unchecked_by_exclude" in SRC and "Снято отметок (исключения)" in SRC
print("G1-G2 ok: статические стражи")

print("ALL FEATURE TESTS PASSED")
