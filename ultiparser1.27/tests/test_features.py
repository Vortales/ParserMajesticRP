"""Тесты фич 1.27: plurals, кламп прогресса, история, diff-диагностика,
картинки (дедуп/нумерация/OCR-гейты), docx-паритет, статические стражи."""
import json
import queue
import re
import sys
import tempfile
import uuid
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from extract27 import (SRC, TARGET, def_block, plural_code, history_code, diff_code,
                       image_block_code, ocr_gate_code, emitter_code, docx_code,
                       dates_code, date_fn_code, exclude_code, match_code,
                       topic_json_code)

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
# v1.27: свежие слои ЗАМЕНЯЮТ накопленные (раньше мёржились навсегда)
assert append(d5, U, "changed", title="УК", checksum="f" * 8,
              scopes=["server:Miami"]) is True
assert t5["scopes"] == ["server:Miami"]
# дубль по checksum событие не пишет, но слои обновляет
assert append(d5, U, "changed", title="УК", checksum="f" * 8,
              scopes=["server:Leo"]) is False
assert t5["scopes"] == ["server:Leo"]
# пустые scopes записанные не затирают
assert append(d5, U, "changed", title="УК", checksum="9" * 8) is True
assert t5["scopes"] == ["server:Leo"]
segs5 = hns["history_event_segments"](t5["events"][0])
assert ("    Слой: Общие правила; Сервер «Memphis»", "ctx") in segs5, segs5
assert hns["history_event_segments"]({"ts": "t", "kind": "new", "scopes": []})[0][1] == "h_date"
print("H5 ok: scopes-replace + строка слоя")

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

# H8: таблица diff БЫЛО/СТАЛО (пары + границы)
tns = {}
exec(def_block("_strip_diff_prefix"), tns)
exec(def_block("pair_diff_lines"), tns)
exec(def_block("render_diff_table"), tns)
assert tns["_strip_diff_prefix"]("БЫЛО:  текст") == "текст"
assert tns["_strip_diff_prefix"]("СТАЛО:x") == "x"
assert tns["_strip_diff_prefix"]("обычный текст") == "обычный текст"
pr = tns["pair_diff_lines"]([("БЫЛО: a", "was"), ("СТАЛО: b", "became"), ("СТАЛО: c", "became")])
assert pr == [("a", "b"), ("", "c")], pr
assert tns["pair_diff_lines"]([("УДАЛЕНО: x", "dlt")]) == [("x", "")]
chunks = tns["render_diff_table"]([("старое длинное", "новое")])
txt = "".join(s for s, _ in chunks)
tbl = txt.split("\n")
assert tbl[-1] == ""
tbl = tbl[:-1]
assert len({len(l) for l in tbl}) == 1, [len(l) for l in tbl]
assert tbl[0].startswith("┌") and tbl[0].endswith("┐")
assert "БЫЛО" in tbl[1] and "СТАЛО" in tbl[1]
assert tbl[-1].startswith("└") and tbl[-1].endswith("┘")
assert any("старое длинное" in l and "новое" in l for l in tbl)
ch2 = tns["render_diff_table"]([("w" * 100, "коротко")], max_col=20)
t2 = "".join(s for s, _ in ch2).split("\n")[:-1]
assert len({len(l) for l in t2}) == 1
assert sum("w" in l for l in t2) >= 2
print("H8 ok: diff-таблица")

# H9: приоритет сервера над общими (кросс-листинг)
pns = {}
exec(def_block("drop_general_when_server"), pns)
dg = pns["drop_general_when_server"]
assert dg(["general", "server:Memphis"]) == ["server:Memphis"]
assert dg(["general"]) == ["general"]
assert dg(["server:Memphis"]) == ["server:Memphis"]
assert dg(["general", "server:A", "server:B"]) == ["server:A", "server:B"]
assert dg([]) == [] and dg(None) == []
assert dg([" general ", "server:X"]) == ["server:X"]
print("H9 ok: drop_general_when_server")

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

# E3: финальное решение по галочке (исключение бьёт всё)
fns = {}
exec(def_block("final_thread_check_value"), fns)
fv = fns["final_thread_check_value"]
assert fv(excluded=True, kw_eval="check", was_checked=True, kw_match=True) is False
assert fv(excluded=True, kw_eval=None, was_checked=True, kw_match=False) is False
assert fv(excluded=True, kw_eval="uncheck", was_checked=False, kw_match=False) is False
assert fv(excluded=False, kw_eval="uncheck", was_checked=True, kw_match=True) is False
assert fv(excluded=False, kw_eval="check", was_checked=True, kw_match=False) is True
assert fv(excluded=False, kw_eval=None, was_checked=False, kw_match=True) is True
assert fv(excluded=False, kw_eval="check", was_checked=False, kw_match=True) is True
assert fv(excluded=False, kw_eval=None, was_checked=False, kw_match=False) is False
print("E3 ok: final_thread_check_value")

# E4: скрытие строк структуры (поиск + исключения)
sns = {"re": re}
exec(exclude_code(), sns)
exec(def_block("struct_row_hidden"), sns)
sh = sns["struct_row_hidden"]
assert sh("Исковое заявление №1", "Раздел", "", ["исковое"]) is True
assert sh("Исковое заявление №1", "Раздел", "", []) is False
assert sh("Закон о полиции", "Раздел", "закон", []) is False
assert sh("Закон о полиции", "Раздел", "кодекс", []) is True
assert sh("Закон о полиции", "Раздел", "", ["полиции"]) is True
assert sh("Закон о полиции", "Отмененные дела", "", ["отмененные"]) is True
assert sh("Закон о полиции", "Раздел", "закон", ["полиции"]) is True
assert sh("Закон о полиции", "Раздел", "", ["полиции"], True) is False
assert sh("Закон о полиции", "Раздел", "", ["полиции"], False) is True
assert sh("Закон о полиции", "Раздел", "кодекс", [], True) is True
assert sh("Закон о полиции", "Раздел", "закон", ["полиции"], True) is False
print("E4 ok: struct_row_hidden + checked-bypass")

# E5: автоотметка тем — только слова тем (слова разделов не действуют)
e5ns = {"re": re}
exec(match_code(), e5ns)
exec(def_block("set_autocheck_keywords"), e5ns)
exec(def_block("should_autocheck_new_thread"), e5ns)
exec(def_block("thread_autocheck_match"), e5ns)
e5ns["set_autocheck_keywords"](["Правила"], ["Уголовный кодекс"], [])
tam = e5ns["thread_autocheck_match"]
assert tam(None, "Правила поведения") is False
assert tam("check", "Правила поведения") is True
assert tam(None, "Уголовный кодекс штата") is True
assert tam("uncheck", "Уголовный кодекс штата") is True
print("E5 ok: thread_autocheck_match")

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

# ---------- J1: сегментация строк на разделы ----------
jns = {"re": re, "json": json}
exec(topic_json_code(), jns)
seg = jns["segment_sections"]
secs = seg(["1. Общее положение", "1.1 Текст правила", "1.2 Ещё правило", "2. Другое"])
assert [s["title"] for s in secs] == ["1. Общее положение", "1.1 Текст правила", "1.2 Ещё правило", "2. Другое"]
assert [s["level"] for s in secs] == [1, 2, 2, 1]
assert [s["line_start"] for s in secs] == [0, 1, 2, 3]
assert secs[1]["text"] == "1.1 Текст правила"
secs = seg(["Преамбула темы", "1. Начало", "строка содержания"])
assert [(s["title"], s["level"]) for s in secs] == [("", 0), ("1. Начало", 1)]
assert secs[0]["line_start"] == 0 and secs[0]["line_end"] == 0
assert "строка содержания" in secs[1]["text"]
secs = seg(["Статья 5. Нарушения", "содержание статьи"])
assert len(secs) == 1 and secs[0]["level"] == 1 and secs[0]["title"].startswith("Статья 5")
secs = seg(["14.07.2023 в 23:39", "1.1 Правило"])
assert [s["title"] for s in secs] == ["", "1.1 Правило"]  # дата — не заголовок
assert seg([]) == [] and seg(["", "   "]) == []
print("J1 ok: segment_sections")

# ---------- J2: карточка темы ultiparser-topic/1 ----------
btj = jns["build_topic_json"]
thr = {"title": "УК", "url": "https://forum.majestic-rp.ru/threads/uk.5/"}
dat = {"title": "УК", "url": thr["url"], "created_date": "01.01.2020",
       "posts": [{"author": "a", "date": "d",
                  "blocks": [{"type": "heading", "text": "Раздел 1"},
                             {"type": "paragraph", "text": "1.1 Текст"}]}]}
card = btj(thr, dat, ["БЫЛО: x"], False, checksum="c" * 8,
           layer="server:Memphis", parsed_at="2026-10-11 12:00:00")
assert card["format"] == "ultiparser-topic/1"
assert card["title"] == "УК" and card["url"] == thr["url"]
assert card["norm_url"] == jns["norm_thread_url"](thr["url"])
assert card["status"] == "changed" and card["changes"] == ["БЫЛО: x"]
assert card["checksum"] == "c" * 8 and card["layer"] == "server:Memphis"
assert card["parsed_at"] == "2026-10-11 12:00:00" and card["created_date"] == "01.01.2020"
assert len(card["posts"]) == 1 and card["posts"][0]["n"] == 1
assert card["posts"][0]["lines"] == ["Раздел 1", "1.1 Текст"]
assert card["posts"][0]["text"] == "Раздел 1\n1.1 Текст"
assert [s["title"] for s in card["sections"]] == ["Раздел 1", "1.1 Текст"]
assert btj(thr, dat, None, True)["status"] == "new"
assert btj(thr, dat, None, False)["status"] == "unchanged"
assert btj(thr, dat, None, False)["changes"] is None
json.dumps(card, ensure_ascii=False)  # сериализуема
print("J2 ok: build_topic_json")

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

# ---------- X3: clean-копия без форматирования ----------
dcl = dict(dxx)
cl = dcl["build_common_docx"](PROMPT, ["П1_тест"], flat=False)
cfp = str(tmp / "clean_common.docx")
cl.save(cfp, clean=True)
cparas = [p.text for p in __import__("docx").Document(cfp).paragraphs]
cne = [x for x in cparas if x.strip()]
assert len(cne) > 1, cne  # абзацы НЕ схлопнуты
ctext = "\n".join(cne)
assert ctext.startswith("Строка1"), ctext[:40]  # промт первый
assert "П1_тест" in ctext and "=" * 60 in ctext  # без чистки =/_
assert "ОБЩИЕ ПРАВИЛА" in ctext
assert ctext.index("Строка1") < ctext.index("ОБЩИЕ ПРАВИЛА") < ctext.index(TAIL[:20])
srv_cl = dcl["build_server_docx"](PROMPT, ["П1"], "Srv", ["С1"], flat=False)
sfp = str(tmp / "clean_srv.docx")
srv_cl.save(sfp, clean=True)
sne = [x.text for x in __import__("docx").Document(sfp).paragraphs if x.text.strip()]
assert len(sne) > 1 and any("ЗАКОНОДАТЕЛЬНАЯ БАЗА: Srv" in x for x in sne)
print("X3 ok: clean-docx без форматирования")

# ---------- R1: контракт снапшота скрытия ----------
def _sim_restore(children, hidden):
    kids = [c for c in children if c not in {h[0] for h in hidden}]
    for iid, par, idx in sorted(hidden, key=lambda e: (str(e[1]), e[2])):
        kids.insert(idx, iid)
    return kids
_kids = ["a", "b", "c", "d", "e"]
# снапшот ДО detach (b->1, d->3): порядок восстанавливается
assert _sim_restore(_kids, [("b", "P", 1), ("d", "P", 3)]) == _kids
# баг 1.25: индексы ПОСЛЕ detach (b->1, d->2) порядок ломают — так нельзя
assert _sim_restore(_kids, [("b", "P", 1), ("d", "P", 2)]) != _kids
print("R1 ok: снапшот-then-detach восстанавливает порядок")

# ---------- G1-G2: статические стражи ----------
assert SRC.count('new_threads_found.append((node.title, title, t["url"]))') == 1
assert SRC.count('disappeared_selected_threads.append((node.title, title, t["url"]))') == 1
for dead in ("_show_topics_report", "_show_threads_diff_report", "_show_state_report",
             "FAST_BROWSER_ONLY_COOLDOWN", "_scope_visible", "_on_hmode", "v_hmode",
             "split_diff_columns", "txt_left", "txt_right", "_combined_text",
             "_save_topic_docx", "btn_scan_deep", "btn_generate", "btn_stage1",
             "_apply_keywords_silent", "node_match"):
    assert dead not in SRC, dead
assert SRC.count("btn_history") >= 3
assert re.search(r"^IMAGE_DOWNLOAD_DEADLINE = 12$", SRC, re.M)
assert re.search(r"^IMAGE_FALLBACK_TIMEOUT = 6$", SRC, re.M)
assert "time.monotonic() - t0 < 2" in SRC and "time.monotonic() - t1 < 4" in SRC
assert 'PARSER_VERSION = "1.27"' in SRC
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
assert "def final_thread_check_value" in SRC and "blocked_by_exclude" in SRC
assert "исключения, финал" in SRC and "рестор тоже уважает исключения" in SRC
assert "def struct_row_hidden" in SRC and "structure_exclude_words" in SRC
assert "StructureExcludeWindow" in SRC and "Слов-исключений «Структуры»" in SRC
assert "def render_diff_table" in SRC and "def pair_diff_lines" in SRC
assert "def _thread_scopes_by_url" in SRC and "БЫЛО" in SRC
assert "def drop_general_when_server" in SRC and "Приоритет сервера" in SRC
assert "def segment_sections" in SRC and "def build_topic_json" in SRC
assert "ultiparser-topic/1" in SRC and "def _save_topic_json" in SRC
assert "def _do_run_all" in SRC and "def _chain_next_after_stage2" in SRC
assert "self._chain" in SRC and "Выполнить всё (2→3)" in SRC
assert "def _global_copy" in SRC and "bind_all" in SRC
assert "btn_stage2" in SRC and "btn_stage3" in SRC and "btn_run_all" in SRC
assert "btn_keywords" in SRC and "Ключевые слова" in SRC
assert "def thread_autocheck_match" in SRC
assert "v_struct_excl_on" in SRC and "def _toggle_struct_excl" in SRC
assert "def _guarded_after_build" in SRC and "Ошибка перестроения дерева" in SRC
assert "clean_Общие правила" in SRC and "flat=False" in SRC
assert "def _prune_unchecked_topics" in SRC and "_layer_by_url" in SRC
assert "def snap(iid)" in SRC and "checked=_checked" in SRC
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
