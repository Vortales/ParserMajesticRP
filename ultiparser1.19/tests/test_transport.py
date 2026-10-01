"""Тесты транспорта 1.19: fast-hit, фолбэки, канонический jar, диагностика,
harvest-retry, адаптивное окно browser-only."""
import sys
import tempfile
import threading
import time as real_time
import types
from pathlib import Path
from urllib.parse import urlparse

sys.path.insert(0, str(Path(__file__).resolve().parent))
from extract19 import transport_code

TMP = Path(tempfile.mkdtemp())

# быстрый сон (запись пауз)
sleeps = []
_orig_sleep = real_time.sleep
real_time.sleep = lambda s: sleeps.append(s)  # noqa: E731


class FakeResponse:
    def __init__(self, status=200, text="", url=""):
        self.status_code = status
        self._text = text
        self.url = url
        self.headers = {"Content-Type": "text/html; charset=utf-8"}
        self.encoding = "utf-8"
        self.ok = status == 200
        self.content = text.encode("utf-8", "ignore")

    @property
    def text(self):
        return self._text

    @text.setter
    def text(self, v):
        self._text = v


class FakeJar(dict):
    def set(self, name, value, domain="", path="/"):
        self[name] = value


HANDLER = {"fn": None}


class FakeSession:
    instances = []

    def __init__(self):
        self.headers = {}
        self.cookies = FakeJar()
        FakeSession.instances.append(self)

    def get(self, url, headers=None, timeout=None, allow_redirects=True):
        fn = HANDLER["fn"]
        assert fn is not None, "handler not set"
        return fn(url, headers, timeout)


fake_requests = types.ModuleType("requests")
fake_requests.Session = FakeSession
fake_requests.cookies = types.SimpleNamespace(RequestsCookieJar=FakeJar)
fake_requests.exceptions = types.SimpleNamespace(Timeout=Exception, ConnectionError=Exception)


class FakeLog:
    def __init__(self):
        self.msgs = []

    def debug(self, m, *a): self.msgs.append(("d", m))
    def info(self, m, *a): self.msgs.append(("i", m))
    def warning(self, m, *a): self.msgs.append(("w", m))
    def error(self, m, *a): self.msgs.append(("e", m))
    def exception(self, m, *a): self.msgs.append(("x", m))


FAKE_LOG = FakeLog()

ns = {
    "requests": fake_requests,
    "get_logger": lambda: FAKE_LOG,
    "find_chrome": lambda: "C:\\chrome.exe",
    "chrome_ver": lambda b: "132.0.6834.83",
    "UA": "UA-DEFAULT",
    "SERVERS_ROOT": TMP,
    "time": real_time,
    "threading": threading,
    "random": __import__("random"),
    "re": __import__("re"),
    "urlparse": urlparse,
}
exec(transport_code(), ns)
HttpClient = ns["HttpClient"]

FORUM = "https://forum.majestic-rp.ru"
GOOD = ("<html><body>" + "x" * 600 +
        '<article class="message message--post"><div class="bbWrapper">hi</div></article>'
        '<a href="/logout/">o</a></body></html>')
LOGIN_PAGE = ("<html><body>" + "w" * 600 +
              '<form action="/login/login"><input name="login"><input name="password"></form></body></html>')
COOKIES = [
    {"name": "xf_user", "value": "123_abc", "domain": "forum.majestic-rp.ru", "path": "/"},
    {"name": "xf_csrf", "value": "tok", "domain": "forum.majestic-rp.ru", "path": "/"},
]

browser_calls = []


def stub_browser(url):
    browser_calls.append(url)
    return GOOD


def fresh_client():
    c = HttpClient()
    c.set_cookies(COOKIES)
    c.browser_fallback = stub_browser
    return c


def reset():
    browser_calls.clear()
    sleeps.clear()
    FakeSession.instances.clear()
    FAKE_LOG.msgs.clear()
    HANDLER["fn"] = None


# T1: fast-hit + канонический jar (отравление ответом не липнет)
reset()
HANDLER["fn"] = lambda u, h, t: FakeResponse(200, GOOD, u)
c = fresh_client()
assert c.get(FORUM + "/threads/a.1/") == GOOD and not browser_calls
sess = FakeSession.instances[-1]
assert sess.cookies.get("xf_user") == "123_abc"
sess.cookies.set("xf_user", "POISONED", domain="forum.majestic-rp.ru", path="/")
assert c.get(FORUM + "/threads/b.2/") == GOOD and not browser_calls
assert c._session().cookies.get("xf_user") == "123_abc", "jar должен пересобраться"
assert c.get_stats()["fast"] == 2
print("T1 ok: fast-hit, jar канонический (poison не липнет)")

# T2: 403 → браузер + причина в статистике
reset()
HANDLER["fn"] = lambda u, h, t: FakeResponse(403, "x", u)
c = fresh_client()
assert c.get(FORUM + "/threads/a.1/") == GOOD and len(browser_calls) == 1
st = c.get_stats()
assert st["refusals"].get("http_403") == 1, st
assert st["last_refusal"]["cause"] == "http_403" and st["last_refusal"]["status"] == 403
print("T2 ok: 403 → browser, refusal учтён:", st["last_refusal"])

# T3: сплит «требуется вход» + форензика на диск
reset()
HANDLER["fn"] = lambda u, h, t: FakeResponse(200, LOGIN_PAGE, FORUM + "/login/")
c = fresh_client()
assert c.get(FORUM + "/threads/a.1/") == GOOD
assert c.get_stats()["refusals"].get("redirect_login") == 1
reset()
HANDLER["fn"] = lambda u, h, t: FakeResponse(200, LOGIN_PAGE, u)
c = fresh_client()
assert c.get(FORUM + "/threads/a.1/") == GOOD
st = c.get_stats()
assert st["refusals"].get("login_form") == 1, st
dumps = list((TMP / "logs" / "fast_refused").glob("*.html"))
assert dumps, "нет дампа форензики"
txt = dumps[0].read_text(encoding="utf-8", errors="replace")
assert "login/login" in txt and "<!-- url=" in txt
print("T3 ok: redirect_login/login_form разделены, дамп:", dumps[0].name)

# T4: harvest-retry — свежие cookies спасают без браузера
reset()
calls = {"n": 0}


def flaky(u, h, t):
    calls["n"] += 1
    if calls["n"] == 1:
        return FakeResponse(200, LOGIN_PAGE, u)
    return FakeResponse(200, GOOD, u)


HANDLER["fn"] = flaky
c = fresh_client()
c.cookie_harvester = lambda: [
    {"name": "xf_user", "value": "999", "domain": "forum.majestic-rp.ru", "path": "/"}]
assert c.get(FORUM + "/threads/a.1/") == GOOD
assert calls["n"] == 2 and not browser_calls, (calls, browser_calls)
assert [x["value"] for x in c.get_cookies() if x["name"] == "xf_user"] == ["999"]
assert any("повторяю быстрым" in m for _, m in FAKE_LOG.msgs)
print("T4 ok: harvest→retry fast без браузера")

# T5: адаптивное окно browser-only 300→600→1200→1800(cap), сброс на успехе
reset()
HANDLER["fn"] = lambda u, h, t: FakeResponse(429, "slow", u)
c = fresh_client()
windows = []
for _ in range(4):
    for i in range(3):
        c.get(f"{FORUM}/threads/w{i}.1/")
    windows.append(round(c.get_stats()["browser_only_left"], -1))
    c._fast_browser_only_until = 0.0  # окно «истекло»
assert windows == [300.0, 600.0, 1200.0, 1800.0], windows
HANDLER["fn"] = lambda u, h, t: FakeResponse(200, GOOD, u)
c.get(FORUM + "/threads/ok.1/")
for i in range(3):
    HANDLER["fn"] = lambda u, h, t: FakeResponse(429, "slow", u)
    c.get(f"{FORUM}/threads/z{i}.1/")
assert round(c.get_stats()["browser_only_left"], -1) == 300.0
print("T5 ok: окна", windows, "+ сброс после успеха")

# T6: страница ошибки принимается, подозрительная → браузер
reset()
ERR = "<html><body>" + "z" * 600 + "<div>oops! we ran into some problems. no permission</div></body></html>"
HANDLER["fn"] = lambda u, h, t: FakeResponse(200, ERR, u)
c = fresh_client()
assert c.get(FORUM + "/threads/a.1/") == ERR and not browser_calls
reset()
EMPTY = "<html><body>" + "y" * 600 + "<div>empty</div></body></html>"
HANDLER["fn"] = lambda u, h, t: FakeResponse(200, EMPTY, u)
c = fresh_client()
assert c.get(FORUM + "/threads/a.1/") == GOOD and len(browser_calls) == 1
assert c.get_stats()["refusals"].get("suspicious") == 1
print("T6 ok: error-page принята, suspicious → browser")

# T7: download_bytes + канонический jar
reset()
img = b"\x89PNG\r\n\x1a\n" + b"q" * 200


def img_fn(u, h, t):
    r = FakeResponse(200, "", u)
    r.content = img
    return r


HANDLER["fn"] = img_fn
c = fresh_client()
assert c.download_bytes(FORUM + "/p.png") == img
c._session().cookies.set("xf_user", "POISON", domain="x", path="/")
assert c.download_bytes(FORUM + "/p2.png") == img
assert c._session().cookies.get("xf_user") == "123_abc"
assert c.download_bytes("https://other/p.png") is None
print("T7 ok: download_bytes + jar")

# T8: verify_fast_session — проба сессии без браузера
reset()
HANDLER["fn"] = lambda u, h, t: FakeResponse(200, GOOD, u)
c = fresh_client()
res = c.verify_fast_session()
assert res["ok"] is True and res["cause"] == "ok", res
assert "logout" in res["marker"], res
assert res["cookies_n"] == 2 and res["xf_user"] is True, res
assert not browser_calls
HANDLER["fn"] = lambda u, h, t: FakeResponse(200, LOGIN_PAGE, u)
res2 = c.verify_fast_session()
assert res2["ok"] is False and res2["cause"] == "login_form", res2
assert res2["xf_user"] is True  # куки есть, но форум их не признал
c2 = HttpClient()
res3 = c2.verify_fast_session()
assert res3["cause"] == "no_cookies" and res3["ok"] is False, res3
print("T8 ok: verify_fast_session")

# T9: merge_cookies зовёт persister только при изменении
reset()
saved = []
c = fresh_client()
c.cookies_persister = lambda cookies: saved.append(list(cookies))
assert c.merge_cookies(COOKIES) is False and not saved
assert c.merge_cookies([{"name": "xf_user", "value": "999", "domain": "forum.majestic-rp.ru"}]) is True
assert len(saved) == 1 and [x["value"] for x in saved[0] if x["name"] == "xf_user"] == ["999"]
print("T9 ok: persister свежих cookies")

# T10: отказ login_form несёт диагностику куки
reset()
HANDLER["fn"] = lambda u, h, t: FakeResponse(200, LOGIN_PAGE, u)
c = fresh_client()
assert c.get(FORUM + "/threads/a.1/") == GOOD and len(browser_calls) == 1
lr = c.get_stats()["last_refusal"]
assert lr["cause"] == "login_form" and lr["cookies"] == 2 and lr["xf_user"] is True, lr
print("T10 ok: last_refusal + cookie diag")

real_time.sleep = _orig_sleep
print("ALL TRANSPORT TESTS PASSED")
