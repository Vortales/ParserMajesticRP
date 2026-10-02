import io
import os
import base64
import re
import sys
import json
import time
import shutil
import webbrowser
import hashlib
import logging
import zipfile
import difflib
import platform
import subprocess
import threading
import queue
import uuid
import random
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext, filedialog
from pathlib import Path
from urllib.parse import urljoin, urlparse
from concurrent.futures import ThreadPoolExecutor
from importlib import import_module

PARSER_NAME = "ultiparser"
PARSER_VERSION = "1.25"

def _pip(pkg: str) -> bool:
    try:
        subprocess.check_call([sys.executable, "-m", "pip", "install", "--upgrade", "-q", pkg])
        return True
    except Exception:
        return False

def _ensure(module: str, pkg: str, required=True) -> bool:
    module = module.strip()
    pkg = pkg.strip()

    try:
        import_module(module)
        return True
    except ImportError:
        print(f"  📦 Устанавливаю {pkg} ...")
        if _pip(pkg):
            try:
                import_module(module)
                return True
            except ImportError:
                pass
        if required:
            print(f"  ❌ Не удалось установить {pkg}")
            return False


print("=" * 60)
print(f"  🚀 ФОРУМ-МЕНЕДЖЕР v{PARSER_VERSION}")
print("=" * 60)
print("\n🔧 Проверка зависимостей...")

HAS_REQUESTS = _ensure("requests", "requests")
HAS_BS4 = _ensure("bs4", "beautifulsoup4")
HAS_DOCX = _ensure("docx", "python-docx")
HAS_LXML = _ensure("lxml", "lxml", required=False)
HAS_SELENIUM = _ensure("selenium", "selenium", required=False)
HAS_PIL = _ensure("PIL", "pillow", required=False)
if not (HAS_REQUESTS and HAS_BS4 and HAS_DOCX):
    print("\n❌ pip install requests beautifulsoup4 python-docx lxml selenium")
    sys.exit(1)

import requests
from bs4 import BeautifulSoup, NavigableString, Tag
from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_LINE_SPACING

BS_PARSER = "lxml" if HAS_LXML else "html.parser"
print(f"✅ Готово (парсер: {BS_PARSER})\n")


# v1.16: папка парсера = папка данных.
# Скрипт лежит в ultiparser1.16/ultiparser1.16.py, все рабочие файлы
# (structure.json, config.json, checks.json, Результаты, _cache, images, logs)
# создаются рядом со скриптом, внутри ultiparser1.16/.
SERVERS_ROOT = Path(__file__).resolve().parent
STRUCTURE_FILE = "structure.json"
CONFIG_FILE = "config.json"
COOKIES_FILE = "cookies.json"
OUTPUT_SUBDIR = "Результаты"
COMMON_OUTPUT_SUBDIR = "Общие правила"
SERVERS_OUTPUT_SUBDIR = "Серверы"
TOPICS_SUBDIR = "Темы"
CACHE_SUBDIR = "_cache"
CHECKS_FILE = "checks.json"
IMAGES_SUBDIR = "images"

OCR_IMAGE_MAX_BYTES = 20 * 1024 * 1024
OCR_MAX_IMAGES_PER_POST = 50
IMAGE_DOWNLOAD_DEADLINE = 12
IMAGE_LOCK_WAIT = 3
IMAGE_FALLBACK_TIMEOUT = 6
OCR_MIN_BYTES = 1500
OCR_MIN_DIM = 30
HISTORY_FILE = "history.json"
HISTORY_MAX_EVENTS = 100
HISTORY_MAX_DIFF_LINES = 200
HISTORY_MAX_LINE_LEN = 500
MAX_DEPTH = 8

PAGE_DELAY_MIN = 3.5
PAGE_DELAY_MAX = 7.0
DDOS_WAIT = 120
MAX_WORKERS = 1
DEFAULT_FIRST_POST_ONLY = True
DEBUG_LINKS = True

# ── v1.16: гибридный транспорт (requests+cookies → браузер-фолбэк) ──
# Быстрый путь идёт ПЕРВЫМ: обычные HTTP-запросы с сохранёнными cookies,
# keep-alive сессиями и браузерными заголовками. Браузер поднимается только
# при блокировках — лениво, по требованию. Сайт чувствителен к темпу:
# глобальный рейт-лимитер держит паузу между ЛЮБЫМИ запросами независимо
# от числа потоков, при сериях блокировок включается режим «только браузер».
FAST_PATH_ENABLED_DEFAULT = True
FAST_DELAY_MIN = 1.2          # пауза перед быстрым запросом (+ джиттер лимитера)
FAST_DELAY_MAX = 2.8
FAST_TIMEOUT = 25             # таймаут одного быстрого запроса, сек
FAST_MAX_ATTEMPTS = 2         # попыток быстрым путём перед уходом в браузер
FAST_BLOCK_THRESHOLD = 3      # блокировок подряд → временный «только браузер»
FAST_BROWSER_ONLY_BASE = 300.0    # стартовое окно «только браузер», сек
FAST_BROWSER_ONLY_MAX = 1800.0    # потолок адаптивного окна, сек
FAST_MIN_INTERVAL = 1.2       # базовый минимальный интервал между запросами, сек
FAST_BACKOFF_MAX_MULT = 4.0   # потолок адаптивного замедления (x к интервалу)
BREAK_STREAK_TRIGGER = 2      # v1.21: блок-сигналов подряд → общая пауза всех потоков
BREAK_LEVELS = (60.0, 120.0, 300.0, 600.0)  # v1.21: эскалация общих пауз, сек
BREAK_RESET_SUCCESSES = 15    # v1.21: чистых хитов подряд → сброс эскалации
BREAK_WAIT_CHUNK = 2.0        # v1.21: квант ожидания (проверка кнопки «Стоп»)
SAFE_WORKERS_FAST = 2         # v1.21: мягкий предел «Потоков» (4 потока палились)

# Маркеры «тема загрузилась нормально» (для проверки быстрых ответов)
THREAD_POST_MARKERS = (
    'article class="message', "message--post", "js-post", "bbwrapper", "message-body",
)
# Маркеры «страница ошибки/нет прав» — такой ответ принимаем как есть, без ретрая
THREAD_ERROR_MARKERS = (
    "недостаточно прав", "no permission", "you do not have permission",
    "у вас нет прав", "oops! we ran into", "oops, we ran into",
    "запрошенная страница не найдена", "страница не найдена",
    "thread not found", "тема не найдена", "an error occurred",
)

DOC_FONT = "Times New Roman"
DOC_SIZE = 4
DOC_LS = 1.0
DOC_MARGIN = 0.4

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
      "AppleWebKit/537.36 (KHTML, like Gecko) "
      "Chrome/131.0.0.0 Safari/537.36")  # используется ТОЛЬКО для requests-сессии (HttpClient)

def build_real_ua(chrome_version: str) -> str:
    """
    Строит UA под реально установленную версию Chrome,
    чтобы избежать рассинхрона fingerprint ↔ заявленный UA,
    который триггерит бесконечную проверку DDoS-Guard/Cloudflare.
    """
    if not chrome_version:
        return UA
    return (f"Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            f"AppleWebKit/537.36 (KHTML, like Gecko) "
            f"Chrome/{chrome_version} Safari/537.36")


FORUM_BASE_URL  = "https://forum.majestic-rp.ru/"
FORUM_LOGIN_URL = "https://forum.majestic-rp.ru/login/"
FORUM_HOST      = "forum.majestic-rp.ru"

BAD_URL_SUFFIXES = (
    "/mark-read", "/mark-unread", "/unread", "/latest", "/latest-activity",
    "/whats-new", "/post-thread", "/create-thread", "/preview", "/watch",
    "/watchers", "/participants", "/prefixes", "/apply-filter", "/reset-filter",
    "/edit", "/delete", "/report", "/react", "/rss", "/index.rss",
    "/subscribe", "/unsubscribe", "/hide", "/move",
)

BAD_URL_CONTAINS = (
    "?direction=", "?order=", "&order=", "?prefix_id=", "&prefix_id=",
    "?starter_id=", "&starter_id=", "?last_days=", "&last_days=",
    "?_xfResponseType=", "&_xfResponseType=", "/page-", "/goto/",
)


def is_bad_url(url: str):
    if not url:
        return True, "пустой URL"
    low = url.lower()
    for suf in BAD_URL_SUFFIXES:
        if low.endswith(suf) or low.endswith(suf + "/"):
            return True, f"суффикс '{suf}'"
        if suf + "/" in low or suf + "?" in low:
            return True, f"содержит '{suf}'"
    for c in BAD_URL_CONTAINS:
        if c in low:
            return True, f"содержит '{c}'"
    return False, ""


_logger = None
LOG_PATH = None


def get_logger() -> logging.Logger:
    global _logger, LOG_PATH

    if _logger:
        return _logger

    SERVERS_ROOT.mkdir(parents=True, exist_ok=True)

    ld = SERVERS_ROOT / "logs"
    ld.mkdir(exist_ok=True)

    LOG_PATH = ld / f"run_{time.strftime('%Y%m%d_%H%M%S')}.log"

    _logger = logging.getLogger("fm")
    _logger.setLevel(logging.DEBUG)

    fmt = logging.Formatter("%(asctime)s | %(levelname)-7s | %(message)s", "%H:%M:%S")

    fh = logging.FileHandler(LOG_PATH, encoding="utf-8")
    fh.setLevel(logging.DEBUG)
    fh.setFormatter(fmt)
    _logger.addHandler(fh)

    ch = logging.StreamHandler(sys.stdout)
    ch.setLevel(logging.INFO)
    ch.setFormatter(fmt)
    _logger.addHandler(ch)

    return _logger


BLOCK_MARKERS = (
    "please turn javascript", "checking your browser", "just a moment",
    "ddos-guard", "cf-browser-verification", "enable javascript and cookies",
    "page could not be loaded", "attention required", "verify you are human",
)

def html_ok(html: str) -> bool:
    if not html or len(html) < 500:
        return False
    low = html[:120000].lower()
    if "page could not be loaded" in low:
        return False
    return not any(m in low for m in BLOCK_MARKERS)


LOGGED_IN_MARKERS = (
    'data-visitor-menu', 'href="/logout', "href='/logout",
    'account-menu', 'p-navgroup--member', 'class="p-navgroup p-account',
    'p-navgroup--user', 'data-logged-in="true"', 'js-logIn"',
    'accountmenu', 'p-navgroup-link--user',
)


def _login_markers_found(html: str) -> list:
    """Какие маркеры входа есть в HTML (v1.19, для диагностики)."""
    if not html:
        return []
    low = html.lower()
    return [m for m in LOGGED_IN_MARKERS if m in low]


STRONG_LOGIN_MARKERS = (
    'data-visitor-menu', 'href="/logout', "href='/logout",
    'p-navgroup--member', 'class="p-navgroup p-account',
)


def _strong_login_markers_found(html: str) -> list:
    """Сильные маркеры входа — почти невозможны у гостя (v1.20)."""
    if not html:
        return []
    low = html.lower()
    return [m for m in STRONG_LOGIN_MARKERS if m in low]


def _page_has_expected_content(url: str, html: str) -> bool:
    """Есть ли на странице ожидаемый контент (v1.20, чистая функция)."""
    low = (html or "")[:200000].lower()
    u = (url or "").lower()
    if "/threads/" in u:
        return any(m in low for m in THREAD_POST_MARKERS)
    if "/forums/" in u or "/categories/" in u:
        return ("structitem" in low) or any(m in low for m in THREAD_POST_MARKERS)
    return True


def analyze_login_refusal(html: str, url: str = "") -> dict:
    """Разбор страницы с формой входа для диагностики (v1.20, чистая функция)."""
    low = (html or "")[:200000].lower()
    has_action = ('action="/login/login"' in low or 'action="login/login"' in low)
    has_inputs = ('name="login"' in low and 'name="password"' in low)
    markers = _login_markers_found(html)
    strong = [m for m in markers if m in STRONG_LOGIN_MARKERS]
    title = ""
    try:
        m = re.search(r"<title[^>]*>(.*?)</title>", (html or "")[:200000], re.I | re.S)
        if m:
            title = re.sub(r"\s+", " ", m.group(1)).strip()[:120]
    except Exception:
        title = ""
    content_ok = _page_has_expected_content(url, html)
    if not markers:
        mstate = "нет"
    elif not strong:
        mstate = "только слабые"
    elif not content_ok:
        mstate = "есть, но нет контента"
    else:
        mstate = "есть"
    if has_action:
        form_bit = "форма входа: action"
    elif has_inputs:
        form_bit = "форма входа: inputs"
    else:
        form_bit = "форма входа: ?"
    hint = f"{form_bit}; маркеры входа: {mstate}"
    if title:
        hint += f"; title: «{title}»"
    return {"form_action": has_action, "form_inputs": has_inputs, "markers": markers,
            "strong": strong, "title": title, "content_ok": content_ok,
            "mstate": mstate, "hint": hint}


def is_logged_in_html(html: str) -> bool:
    if not html:
        return False
    low = html.lower()
    # Расширенный список — разные темы XenForo называют блоки по-разному
    if _login_markers_found(html):
        if 'name="login"' in low and 'name="password"' in low and 'action="/login/login"' in low:
            return False
        return True
    return False


def has_xf_session_cookie(cookies: list) -> bool:
    """
    Основной куки авторизации XenForo 2 — 'xf_user'.
    Он появляется ТОЛЬКО после реального входа и не зависит
    от того, как названа тема/разметка форума.
    Это более надёжный признак входа, чем парсинг HTML.
    """
    for c in cookies or []:
        name = (c.get("name") or "").lower()
        val = c.get("value") or ""
        if name == "xf_user" and val and val not in ("", "0", "deleted"):
            return True
    return False


def find_chrome():
    cands = []
    if sys.platform == "win32":
        pf = os.environ.get("PROGRAMFILES", r"C:\Program Files")
        pf86 = os.environ.get("PROGRAMFILES(X86)", r"C:\Program Files (x86)")
        loc = os.environ.get("LOCALAPPDATA", "")
        cands = [
            rf"{pf}\Google\Chrome\Application\chrome.exe",
            rf"{pf86}\Google\Chrome\Application\chrome.exe",
            rf"{loc}\Google\Chrome\Application\chrome.exe",
            rf"{pf}\BraveSoftware\Brave-Browser\Application\brave.exe",
            rf"{pf86}\Microsoft\Edge\Application\msedge.exe",
        ]
    elif sys.platform == "darwin":
        cands = [
            "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
            "/Applications/Chromium.app/Contents/MacOS/Chromium",
        ]
    else:
        for n in ("google-chrome", "google-chrome-stable", "chromium", "chromium-browser"):
            p = shutil.which(n)
            if p:
                cands.append(p)
        cands += ["/usr/bin/google-chrome", "/usr/bin/chromium", "/snap/bin/chromium"]
    for c in cands:
        if c and Path(c).exists():
            return c
    return None

def kill_stray_automation_processes():
    """
    Убивает зависшие процессы chromedriver и chrome, запущенные
    с нашим автоматизированным профилем. НЕ трогает обычный Chrome
    пользователя (фильтруем по --user-data-dir=...fm_chrome_profile).
    """
    try:
        if sys.platform == "win32":
            # chromedriver.exe безопасно гасить всегда — это не браузер пользователя
            subprocess.run(
                ["taskkill", "/F", "/IM", "chromedriver.exe"],
                capture_output=True, timeout=10,
            )
            # chrome.exe гасим точечно — только процессы с нашим профилем
            try:
                out = subprocess.check_output(
                    ["wmic", "process", "where",
                     "name='chrome.exe'", "get", "ProcessId,CommandLine"],
                    timeout=10, stderr=subprocess.DEVNULL,
                ).decode(errors="ignore")
                marker = str(PERSISTENT_PROFILE).lower()
                for line in out.splitlines():
                    if marker in line.lower():
                        m = re.search(r"(\d+)\s*$", line.strip())
                        if m:
                            subprocess.run(
                                ["taskkill", "/F", "/PID", m.group(1)],
                                capture_output=True, timeout=5,
                            )
            except Exception:
                pass
        else:
            subprocess.run(["pkill", "-f", "chromedriver"], capture_output=True, timeout=10)
            subprocess.run(["pkill", "-f", str(PERSISTENT_PROFILE)], capture_output=True, timeout=10)
    except Exception:
        pass

def chrome_ver(binary: str):
    if sys.platform == "win32":
        try:
            import winreg
            for hive in (winreg.HKEY_CURRENT_USER, winreg.HKEY_LOCAL_MACHINE):
                try:
                    k = winreg.OpenKey(hive, r"SOFTWARE\Google\Chrome\BLBeacon")
                    v, _ = winreg.QueryValueEx(k, "version")
                    return v
                except OSError:
                    continue
        except Exception:
            pass
        try:
            d = Path(binary).parent
            vs = [x.name for x in d.iterdir() if x.is_dir() and re.fullmatch(r"\d+\.\d+\.\d+\.\d+", x.name)]
            if vs:
                return sorted(vs, key=lambda s: [int(x) for x in s.split(".")])[-1]
        except Exception:
            pass
        return None
    try:
        out = subprocess.check_output([binary, "--version"], timeout=10).decode()
        m = re.search(r"(\d+\.\d+\.\d+\.\d+)", out)
        return m.group(1) if m else None
    except Exception:
        return None


def drv_platform() -> str:
    if sys.platform == "win32":
        return "win64" if platform.machine().endswith("64") else "win32"
    if sys.platform == "darwin":
        return "mac-arm64" if platform.machine() == "arm64" else "mac-x64"
    return "linux64"


DRV_CACHE = Path.home() / ".fm_drivers"


def get_chromedriver(version):
    DRV_CACHE.mkdir(exist_ok=True)
    plat = drv_platform()
    exe = "chromedriver.exe" if sys.platform == "win32" else "chromedriver"
    major = version.split(".")[0] if version else None
    if major:
        cached = DRV_CACHE / f"{major}_{plat}" / exe
        if cached.exists():
            return str(cached)
    urls = []
    if version:
        urls.append(f"https://storage.googleapis.com/chrome-for-testing-public/{version}/{plat}/chromedriver-{plat}.zip")
    if major:
        for ep in (
            f"https://googlechromelabs.github.io/chrome-for-testing/LATEST_RELEASE_{major}",
            f"https://chromedriver.storage.googleapis.com/LATEST_RELEASE_{major}",
        ):
            try:
                r = requests.get(ep, timeout=12)
                if r.ok and re.fullmatch(r"[\d.]+", r.text.strip()):
                    v = r.text.strip()
                    urls.append(f"https://storage.googleapis.com/chrome-for-testing-public/{v}/{plat}/chromedriver-{plat}.zip")
            except Exception:
                pass
    try:
        r = requests.get("https://googlechromelabs.github.io/chrome-for-testing/last-known-good-versions.json", timeout=12)
        if r.ok:
            v = r.json()["channels"]["Stable"]["version"]
            urls.append(f"https://storage.googleapis.com/chrome-for-testing-public/{v}/{plat}/chromedriver-{plat}.zip")
    except Exception:
        pass
    seen, uniq = set(), []
    for u in urls:
        if u not in seen:
            seen.add(u)
            uniq.append(u)
    for u in uniq:
        try:
            r = requests.get(u, timeout=90)
            if not r.ok:
                continue
            target = DRV_CACHE / f"{major or 'latest'}_{plat}"
            target.mkdir(parents=True, exist_ok=True)
            with zipfile.ZipFile(io.BytesIO(r.content)) as z:
                for m in z.namelist():
                    if m.endswith(exe):
                        with z.open(m) as src, open(target / exe, "wb") as dst:
                            shutil.copyfileobj(src, dst)
                        break
            p = target / exe
            if p.exists():
                if sys.platform != "win32":
                    os.chmod(p, 0o755)
                return str(p)
        except Exception:
            pass
    return None

PERSISTENT_PROFILE = Path.home() / ".fm_chrome_profile"

def make_options(cls, chrome_bin, chrome_version=None, use_profile=True):
    o = cls()

    if use_profile:
        PERSISTENT_PROFILE.mkdir(parents=True, exist_ok=True)
        o.add_argument(f"--user-data-dir={PERSISTENT_PROFILE}")
    else:
        import tempfile
        tmp_dir = tempfile.mkdtemp(prefix="fm_chrome_tmp_")
        o.add_argument(f"--user-data-dir={tmp_dir}")

    args = [
        "--no-sandbox",
        "--disable-dev-shm-usage",
        "--window-size=1400,900",
        "--lang=ru-RU,ru",
        "--no-first-run",
        "--no-default-browser-check",
        "--disable-extensions",
        "--disable-blink-features=AutomationControlled",
        "--disable-popup-blocking",
        "--remote-allow-origins=*",
        "--disable-infobars",
        "--disable-features=IsolateOrigins,site-per-process,TranslateUI,AutomationControllerForContentScripts,MediaRouter,OptimizationHints,InterestFeedContentSuggestions",
        "--disable-site-isolation-trials",
        "--disable-hang-monitor",
        "--disable-background-networking",
        "--disable-sync",
        "--metrics-recording-only",
        "--enable-features=NetworkService,NetworkServiceInProcess",
        "--disable-backgrounding-occluded-windows",
        "--disable-renderer-backgrounding",
        "--mute-audio",
        "--hide-scrollbars",
        "--disable-component-update",
        "--disable-domain-reliability",
        "--disable-client-side-phishing-detection",
    ]

    if chrome_version:
        args.append(f"--user-agent={build_real_ua(chrome_version)}")

    for a in args:
        o.add_argument(a)

    try:
        o.add_experimental_option("excludeSwitches", ["enable-automation"])
        o.add_experimental_option("useAutomationExtension", False)
        o.add_experimental_option("prefs", {
            "credentials_enable_service": False,
            "profile.password_manager_enabled": False,
            "download.prompt_for_download": True,
            "safebrowsing.enabled": True,
        })
    except Exception:
        pass

    if chrome_bin:
        o.binary_location = chrome_bin

    return o

def build_real_ua(chrome_version: str) -> str:
    if not chrome_version:
        return UA
    return (f"Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            f"AppleWebKit/537.36 (KHTML, like Gecko) "
            f"Chrome/{chrome_version} Safari/537.36")

class ProcessEmitter:
    def __init__(self):
        self.q = queue.Queue()
        self._process_id = None

    def _id(self):
        return uuid.uuid4().hex[:12]

    def process_started(self, label, total=0):
        pid = self._id()
        self._process_id = pid
        self.q.put({"type": "process_started", "id": pid, "label": label, "total": total})
        return pid

    def step_added(self, parent_id, label, icon=""):
        sid = self._id()
        self.q.put({"type": "step_added", "id": sid, "parent": parent_id, "label": label, "icon": icon})
        return sid

    def step_state(self, step_id, state, detail=""):
        self.q.put({"type": "step_state", "id": step_id, "state": state, "detail": detail})

    def step_metric(self, step_id, key, value):
        self.q.put({"type": "step_metric", "id": step_id, "key": key, "value": value})

    def step_progress(self, step_id, current, total):
        if total > 0:
            # v1.17: кламп — проценты и полоса никогда не вылезают за 100%.
            ratio = min(1.0, max(0.0, current / total))
            pct = int(ratio * 100)
            bar_len = 10
            filled = int(bar_len * ratio)
            bar = "▓" * filled + "░" * (bar_len - filled)
            self.q.put({"type": "step_metric", "id": step_id, "key": "progress", "value": f"[{bar}] {pct}%"})
        else:
            self.q.put({"type": "step_metric", "id": step_id, "key": "progress", "value": f"{current}/{total}"})

    def process_finished(self, state="done", elapsed=0):
        self.q.put({"type": "process_finished", "id": self._process_id, "state": state, "elapsed": elapsed})

class Browser:
    def __init__(self):
        self.driver = None

    def start(self, log_cb=None, emitter=None, parent_id=None) -> bool:
        if not HAS_SELENIUM:
            if log_cb:
                log_cb("❌ selenium не установлен")
            return False
        from selenium import webdriver
        from selenium.webdriver.chrome.options import Options
        from selenium.webdriver.chrome.service import Service
        cb = log_cb or print
        em = emitter
        pid = parent_id

        s_find = None
        if em and pid:
            s_find = em.step_added(pid, "Поиск Chrome", "🔍")
            em.step_state(s_find, "running")
        bin_ = find_chrome()
        if not bin_:
            cb("❌ Google Chrome не найден")
            if em and s_find:
                em.step_state(s_find, "error", "Chrome не найден")
            return False
        if em and s_find:
            em.step_state(s_find, "done", bin_[-40:])

        s_ver = None
        if em and pid:
            s_ver = em.step_added(pid, "Определение версии", "🔢")
            em.step_state(s_ver, "running")
        ver = chrome_ver(bin_)
        cb(f"🌐 Chrome: {bin_}  ({ver or '?'})")
        if em and s_ver:
            em.step_state(s_ver, "done", ver or "?")

        s_kill = None
        if em and pid:
            s_kill = em.step_added(pid, "Убийство зависших процессов", "🧹")
            em.step_state(s_kill, "running")
        kill_stray_automation_processes()
        time.sleep(0.5)
        if em and s_kill:
            em.step_state(s_kill, "done")

        s_drv = None
        if em and pid:
            s_drv = em.step_added(pid, "Получение chromedriver", "⬇")
            em.step_state(s_drv, "running")
        STEALTH_JS = """
        Object.defineProperty(navigator, 'webdriver', {get: () => undefined});
        Object.defineProperty(navigator, 'languages', {get: () => ['ru-RU', 'ru', 'en-US', 'en']});
        Object.defineProperty(navigator, 'plugins', {get: () => [1, 2, 3, 4, 5]});
        Object.defineProperty(navigator, 'hardwareConcurrency', {get: () => 8});
        Object.defineProperty(navigator, 'deviceMemory', {get: () => 8});
        Object.defineProperty(navigator, 'maxTouchPoints', {get: () => 0});
        Object.defineProperty(navigator, 'platform', {get: () => 'Win32'});
        Object.defineProperty(navigator, 'vendor', {get: () => 'Google Inc.'});
        Object.defineProperty(navigator, 'doNotTrack', {get: () => null});
        window.chrome = {runtime: {}, loadTimes: function(){}, csi: function(){}, app: {}};
        const originalQuery = window.navigator.permissions.query;
        window.navigator.permissions.query = (parameters) => (
            parameters.name === 'notifications' ?
                Promise.resolve({state: Notification.permission}) :
                originalQuery(parameters)
        );
        const getParameter = WebGLRenderingContext.prototype.getParameter;
        WebGLRenderingContext.prototype.getParameter = function(parameter) {
            if (parameter === 37445) return 'Intel Inc.';
            if (parameter === 37446) return 'Intel Iris OpenGL Engine';
            return getParameter.call(this, parameter);
        };
        """

        def try_svc(label, make, use_profile=True, headless=False, crash_log_path=None):
            s_try = None
            if em and pid:
                s_try = em.step_added(pid, f"Попытка: {label}", "🔌")
                em.step_state(s_try, "running")
            try:
                svc = make()
                if svc is None:
                    if em and s_try:
                        em.step_state(s_try, "skipped")
                    return False
                cb(f"  🔌 {label}")
                opts = make_options(Options, bin_, chrome_version=ver, use_profile=use_profile)
                if headless:
                    opts.add_argument("--headless=new")
                if crash_log_path:
                    opts.add_argument("--enable-logging")
                    opts.add_argument("--v=1")
                    opts.add_argument(f"--log-file={crash_log_path}")
                self.driver = webdriver.Chrome(service=svc, options=opts)
                self.driver.set_page_load_timeout(60)
                try:
                    self.driver.execute_cdp_cmd(
                        "Page.addScriptToEvaluateOnNewDocument",
                        {"source": STEALTH_JS},
                    )
                except Exception:
                    pass
                cb(f"  ✅ Браузер запущен ({label})")
                if em and s_try:
                    em.step_state(s_try, "done")
                return True
            except Exception as e:
                err = str(e).strip()
                cb(f"  ✖ {label}:")
                for line in err.splitlines()[:8]:
                    cb(f"      {line}")
                get_logger().error(f"{label} full error:\n{err}")
                if crash_log_path and Path(crash_log_path).exists():
                    try:
                        tail = Path(crash_log_path).read_text(errors="ignore").splitlines()[-25:]
                        cb("  📋 Хвост chrome_debug.log:")
                        for line in tail:
                            cb(f"      {line}")
                    except Exception:
                        pass
                if em and s_try:
                    em.step_state(s_try, "error", err[:80])
                return False

        crash_log = str(SERVERS_ROOT / "_logs" / "chrome_debug.log")
        Path(crash_log).parent.mkdir(parents=True, exist_ok=True)
        p = get_chromedriver(ver)
        if em and s_drv:
            if p:
                em.step_state(s_drv, "done", "из кэша")
            else:
                em.step_state(s_drv, "warn", "не найден")
        if p and try_svc("Авто-драйвер (профиль)",
                         lambda: Service(executable_path=p),
                         use_profile=True, crash_log_path=crash_log):
            return True
        if try_svc("Selenium Manager (профиль)",
                   lambda: Service(),
                   use_profile=True, crash_log_path=crash_log):
            return True
        cb("  ℹ Пробую с временным чистым профилем...")
        if try_svc("Selenium Manager (временный профиль)",
                   lambda: Service(),
                   use_profile=False, crash_log_path=crash_log):
            cb("  ⚠ Постоянный профиль несовместим — сессия НЕ будет сохраняться "
               "между запусками, пока не пересоздан персистентный профиль.")
            return True
        cb("  ℹ Пробую headless-режим для диагностики...")
        if try_svc("Selenium Manager (headless, диагностика)",
                   lambda: Service(),
                   use_profile=False, headless=True, crash_log_path=crash_log):
            cb("  ⚠ Headless запустился, а обычный режим — нет. "
               "Похоже, проблема в GPU/видеодрайвере или антивирусе, "
               "блокирующем отрисовку окна Chrome.")
            return True
        wp = shutil.which("chromedriver")
        if wp and try_svc("PATH chromedriver", lambda: Service(executable_path=wp),
                          use_profile=False, crash_log_path=crash_log):
            return True
        cb("❌ Не удалось запустить Chrome ни одним из способов.")
        cb(f"   Полный лог диагностики Chrome: {crash_log}")
        cb("   Проверьте антивирус/EDR — часто именно он блокирует автоматизацию.")
        return False

    def get(self, url: str, emitter=None, parent_id=None):
        if not self.driver:
            return None
        em = emitter
        s_get = None
        if em and parent_id:
            s_get = em.step_added(parent_id, f"GET → {url[:60]}", "📡")
            em.step_state(s_get, "running")
        try:
            self.driver.get(url)
        except Exception as e:
            get_logger().warning(f"browser.get: {e}")
            if em and s_get:
                em.step_state(s_get, "error", str(e)[:60])
            return None
        try:
            time.sleep(random.uniform(0.5, 1.5))
            self.driver.execute_script(
                "window.scrollBy({top: Math.floor(Math.random() * 300) + 100, behavior: 'smooth'});"
            )
            time.sleep(random.uniform(0.3, 1.0))
        except Exception:
            pass
        result = self._wait(url_for_retry=url, emitter=em, parent_id=s_get)
        if em and s_get:
            if result:
                size_kb = len(result) // 1024
                em.step_state(s_get, "done", f"{size_kb} КБ")
            else:
                em.step_state(s_get, "error", "не загружено")
        return result

    def _wait(self, url_for_retry: str = None, emitter=None, parent_id=None):
        if not self.driver:
            return None

        em = emitter
        s_ddos = None
        if em and parent_id:
            s_ddos = em.step_added(parent_id, "Ожидание DDoS-проверки", "⏳")
            em.step_state(s_ddos, "running")

        limit = max(DDOS_WAIT, 180)
        t0 = time.monotonic()
        html = ""

        while time.monotonic() - t0 < limit:
            try:
                html = self.driver.page_source
            except Exception:
                html = ""

            if html_ok(html):
                time.sleep(random.uniform(0.8, 2.0))
                try:
                    result = self.driver.page_source
                except Exception:
                    result = html

                if result and html_ok(result):
                    if em and s_ddos:
                        em.step_state(s_ddos, "done", f"{time.monotonic() - t0:.1f}с")
                    return result

            low = (html or "").lower()
            if "page could not be loaded" in low or any(m in low for m in BLOCK_MARKERS):
                time.sleep(random.uniform(3.0, 7.0))
            else:
                time.sleep(random.uniform(1.0, 2.5))

        if url_for_retry:
            if em and s_ddos:
                em.step_state(s_ddos, "warn", "не прошла, пробую reload")

            s_reload = None
            if em and parent_id:
                s_reload = em.step_added(parent_id, "Мягкий reload + повтор", "🔄")
                em.step_state(s_reload, "running")

            try:
                try:
                    self.driver.execute_script("window.location.reload();")
                except Exception:
                    self.driver.get(url_for_retry)

                time.sleep(3)

                t1 = time.monotonic()
                while time.monotonic() - t1 < limit:
                    html = self.driver.page_source
                    if html_ok(html):
                        if em and s_reload:
                            em.step_state(s_reload, "done", f"{time.monotonic() - t1:.1f}с")
                        if em and s_ddos:
                            em.step_state(s_ddos, "done", f"{time.monotonic() - t0:.1f}с")
                        return html
                    time.sleep(random.uniform(2.0, 5.0))

            except Exception as e:
                get_logger().warning(f"retry reload failed: {e}")
                if em and s_reload:
                    em.step_state(s_reload, "error", str(e)[:60])

            if em and s_reload:
                em.step_state(s_reload, "error", "не прошла")

        if em and s_ddos:
            em.step_state(s_ddos, "error", f"не прошла за {limit}с")

        return html or None

    def export_cookies(self):
        if not self.driver:
            return []
        try:
            return self.driver.get_cookies()
        except Exception as e:
            get_logger().warning(f"export_cookies: {e}")
            return []

    def check_logged_in(self, base_url: str, log_cb=None) -> bool:
        cb = log_cb or (lambda s: None)
        if not self.driver:
            return False
        try:
            cookies = self.export_cookies()
            names = sorted(set(c.get("name", "") for c in cookies))
            get_logger().debug(f"Экспортировано cookies: {len(cookies)} — {names}")

            # 1) Самый надёжный признак — cookie xf_user не зависит от темы форума
            if has_xf_session_cookie(cookies):
                cb("  🍪 Найден cookie xf_user — вход подтверждён по cookie")
                return True

            # 2) Фолбэк — парсим HTML, даём странице больше времени прогрузиться
            self.driver.get(base_url)
            time.sleep(3)
            html = self.driver.page_source
            ok = is_logged_in_html(html)
            if not ok:
                cb("  ⚠ Не найден ни cookie xf_user, ни маркеры входа в HTML")
            return ok
        except Exception as e:
            get_logger().warning(f"check_logged_in: {e}")
            return False

    def close(self):
        try:
            if self.driver:
                self.driver.quit()
        except Exception:
            pass
        finally:
            self.driver = None

BROWSER_FALLBACK_THRESHOLD = 3   # после стольких блокировок подряд — сразу браузер, без повторных попыток requests

class _RateLimiter:
    """Глобальный ограничитель темпа запросов (v1.16).

    Гарантирует минимальный интервал между ЛЮБЫМИ двумя запросами быстрым путём,
    независимо от числа потоков. При блокировках временно увеличивает интервал
    (backoff), при успехах — плавно возвращается к базовому.
    """

    def __init__(self, base_interval: float):
        self._base = max(0.2, float(base_interval))
        self._mult = 1.0
        self.pace_mult = 1.0  # v1.21: «Тихий режим» ставит 3.0
        self._next_allowed = 0.0
        self._lock = threading.Lock()

    def wait(self):
        with self._lock:
            now = time.monotonic()
            delay = self._next_allowed - now
            sleep_for = delay if delay > 0 else 0.0
            # Резервируем слот СРАЗУ (под локом), чтобы параллельные потоки
            # выстраивались в очередь, а не шли пачкой после одного ожидания.
            interval = self._base * self._mult * self.pace_mult + random.uniform(0.0, 0.6)
            self._next_allowed = max(now, self._next_allowed) + interval
        if sleep_for > 0:
            time.sleep(sleep_for)

    def report(self, ok: bool):
        with self._lock:
            if ok:
                self._mult = max(1.0, self._mult * 0.9)
            else:
                self._mult = min(FAST_BACKOFF_MAX_MULT, self._mult * 1.5 + 0.5)


def _is_login_url(url: str) -> bool:
    try:
        path = urlparse(url or "").path.lower()
    except Exception:
        return False
    return "/login" in path


def _looks_like_login_page(html: str) -> bool:
    if not html:
        return False
    low = html[:200000].lower()
    has_form = ('action="/login/login"' in low or 'action="login/login"' in low
                or ('name="login"' in low and 'name="password"' in low))
    if not has_form:
        return False
    return not is_logged_in_html(html)


def _suspicious_thread_html(url: str, html: str) -> bool:
    """Тема без постов и без признаков ошибки — возможно, битый/урезанный ответ."""
    if "/threads/" not in (url or "").lower():
        return False
    if not html or not html_ok(html):
        return False
    low = html[:200000].lower()
    if any(m in low for m in THREAD_POST_MARKERS):
        return False
    if any(m in low for m in THREAD_ERROR_MARKERS):
        return False
    return True


class HttpClient:
    """Гибридный транспорт v1.16: быстрый путь requests+cookies, браузер — фолбэк.

    Порядок попытки для каждого URL форума:
      1) быстрый путь — requests-сессия с сохранёнными cookies (keep-alive,
         браузерные заголовки, глобальный рейт-лимитер);
      2) при блокировке/подозрительном ответе — существующий путь через браузер
         (логика 1.15 без изменений);
      3) после серии блокировок — временный режим «только браузер».
    Публичный API get()/set_cookies()/get_cookies()/browser_fallback сохранён.
    """

    def __init__(self):
        self._lock = threading.Lock()
        self._cookies = []
        self._cookie_rev = 0
        self.browser_fallback = None
        self.browser_ensurer = None      # callable() -> bool: ленивый старт браузера
        self.cookie_harvester = None     # callable() -> list|None: свежие cookies из браузера
        self.cookies_persister = None    # v1.19: callable(list) — сохранить свежие cookies в файл
        self._fallback_lock = threading.Lock()
        self.no_browser_errors = 0
        # v1.21: общий рубильник — при блок-сигналах ВСЕ потоки и ОБА канала ждут.
        self._break_lock = threading.Lock()
        self._break_until = 0.0
        self._break_level = 0
        self._break_streak = 0
        self._break_ok_run = 0
        self.abort_event = None      # threading.Event «Стоп» от приложения
        self.stealth_mult = 1.0      # «Тихий режим» ставит 3.0
        self.fast_enabled = FAST_PATH_ENABLED_DEFAULT
        self._tls = threading.local()
        self._rate = _RateLimiter(FAST_MIN_INTERVAL)
        self._block_count = 0
        self._block_lock = threading.Lock()
        self._cooldown_until = 0.0
        self._fast_block_count = 0
        self._fast_lock = threading.Lock()
        self._fast_browser_only_until = 0.0
        self._fast_only_cycles = 0
        self._refusals = {}
        self._last_refusal = None
        self._forensics_dumped = 0
        self._ua_override = ""
        self._ua_detected = False
        self._last_ensure_attempt = 0.0
        self._auth_warn_at = 0.0
        self._stats_lock = threading.Lock()
        self.fast_hits = 0
        self.fast_trusted = 0    # v1.20: принято по маркерам входа (форма проигнорирована)
        self.fast_blocks = 0
        self.browser_hits = 0

    # ── cookies ──────────────────────────────────────────────
    def set_cookies(self, cookies):
        with self._lock:
            self._cookies = list(cookies or [])
            self._cookie_rev += 1
            self._fast_only_cycles = 0

    def get_cookies(self):
        with self._lock:
            return list(self._cookies)

    def has_cookies(self) -> bool:
        with self._lock:
            return bool(self._cookies)

    def merge_cookies(self, cookies) -> bool:
        """Вливает свежие cookies (из браузера) в общий набор. True, если что-то изменилось."""
        if not cookies:
            return False
        changed = False
        with self._lock:
            cur = {c.get("name"): c for c in self._cookies if (c or {}).get("name")}
            for c in cookies:
                name = (c or {}).get("name")
                if not name:
                    continue
                old = cur.get(name)
                if not old or old.get("value") != c.get("value"):
                    changed = True
                cur[name] = dict(c)
            if changed:
                self._cookies = list(cur.values())
                self._cookie_rev += 1
                self._fast_only_cycles = 0
                to_persist = list(cur.values())
            else:
                to_persist = None
        if to_persist and self.cookies_persister:
            try:
                self.cookies_persister(to_persist)
            except Exception:
                pass
        return changed

    # ── UA ───────────────────────────────────────────────────
    def set_ua(self, ua: str):
        if ua:
            with self._lock:
                self._ua_override = ua

    def ensure_real_ua(self):
        """Однажды определяет версию Chrome и фиксирует UA для сессий (best-effort)."""
        with self._lock:
            if self._ua_detected:
                return
            self._ua_detected = True
        try:
            binary = find_chrome()
            ver = chrome_ver(binary) if binary else None
            if ver:
                self.set_ua(build_real_ua(ver))
        except Exception:
            pass

    # ── v1.21: общий рубильник антиблока ──────────────────────
    def _note_block_signal(self, weight: int = 1) -> float:
        """Блок-сигнал с любого канала. Серия → общая пауза. Возвращает паузу."""
        open_for = 0.0
        with self._break_lock:
            self._break_streak += weight
            self._break_ok_run = 0
            if self._break_streak >= BREAK_STREAK_TRIGGER:
                self._break_streak = 0
                lvl = min(self._break_level, len(BREAK_LEVELS) - 1)
                open_for = BREAK_LEVELS[lvl]
                self._break_level = min(self._break_level + 1, len(BREAK_LEVELS) - 1)
                self._break_until = time.monotonic() + open_for
        if open_for > 0:
            get_logger().warning(
                f"breaker: форум ограничил доступ — общая пауза {open_for:.0f}с "
                f"(все потоки и оба канала ждут, ретраи остановлены)")
        return open_for

    def _note_clean_hit(self):
        """Чистый хит: гасит серию сигналов, серией — сбрасывает эскалацию."""
        with self._break_lock:
            self._break_streak = 0
            self._break_ok_run += 1
            if self._break_ok_run >= BREAK_RESET_SUCCESSES:
                self._break_ok_run = 0
                self._break_level = 0

    def _sleep_abortable(self, seconds: float) -> bool:
        """Сон квантами с проверкой abort (v1.21). False — прервано «Стоп»."""
        left = max(0.0, seconds)
        while left > 0:
            time.sleep(min(BREAK_WAIT_CHUNK, left))
            left -= BREAK_WAIT_CHUNK
            try:
                if self.abort_event is not None and self.abort_event.is_set():
                    return False
            except Exception:
                pass
        return True

    def _wait_breaker(self, emitter=None, parent_id=None) -> bool:
        """Ждёт конца общей паузы (v1.21). False — прервано через abort."""
        with self._break_lock:
            left = self._break_until - time.monotonic()
        if left <= 0:
            return True
        if emitter and parent_id:
            try:
                s_br = emitter.step_added(parent_id, f"⏸ Форум отдыхает {left:.0f}с", "⏸")
                emitter.step_state(s_br, "paused")
            except Exception:
                pass
        # Снимок паузы ждём один раз; если другой поток продлил — дождём
        # остаток при следующем запросе (все запросы идут через эту точку).
        return self._sleep_abortable(left)

    # ── статистика ───────────────────────────────────────────
    def get_stats(self) -> dict:
        with self._stats_lock:
            with self._fast_lock:
                left = max(0.0, self._fast_browser_only_until - time.monotonic())
            with self._break_lock:
                bleft = max(0.0, self._break_until - time.monotonic())
            return {
                "fast": self.fast_hits,
                "fast_trusted": self.fast_trusted,
                "fast_blocks": self.fast_blocks,
                "browser": self.browser_hits,
                "enabled": self.fast_enabled,
                "browser_only_left": left,
                "breaker_left": bleft,
                "refusals": dict(self._refusals),
                "last_refusal": dict(self._last_refusal) if self._last_refusal else None,
            }

    def _bump(self, name: str):
        with self._stats_lock:
            if name == "fast":
                self.fast_hits += 1
            elif name == "fast_trusted":
                self.fast_trusted += 1
            elif name == "fast_block":
                self.fast_blocks += 1
            elif name == "browser":
                self.browser_hits += 1

    def _cookie_diag(self) -> str:
        """Короткая диагностика cookie-jar для логов (v1.19)."""
        try:
            with self._lock:
                cookies = list(self._cookies)
        except Exception:
            cookies = []
        has_user = False
        for c in cookies or []:
            try:
                if (c.get("name") or "") == "xf_user" and (c.get("value") or "") not in ("", "0", "deleted"):
                    has_user = True
                    break
            except Exception:
                continue
        return f"cookies: {len(cookies)}, xf_user: {'есть' if has_user else 'НЕТ'}"

    # ── v1.17: диагностика отказов быстрого пути ─────────────────
    def _record_refusal(self, cause: str, url: str, status: int,
                        final_url: str, html: str):
        """Причина отказа fast + образец HTML для разбора (best-effort)."""
        try:
            with self._stats_lock:
                self._refusals[cause] = self._refusals.get(cause, 0) + 1
                try:
                    with self._lock:
                        _ck = list(self._cookies)
                except Exception:
                    _ck = []
                _has_user = any(((c or {}).get("name") == "xf_user"
                                 and (c or {}).get("value") not in ("", "0", "deleted", None))
                                for c in _ck)
                self._last_refusal = {
                    "ts": time.strftime("%H:%M:%S"),
                    "cause": cause,
                    "url": url or "",
                    "status": int(status or 0),
                    "final": final_url or "",
                    "kb": len(html or "") // 1024,
                    "cookies": len(_ck),
                    "xf_user": bool(_has_user),
                }
            if html and self._forensics_dumped < 3:
                self._forensics_dumped += 1
                d = SERVERS_ROOT / "logs" / "fast_refused"
                d.mkdir(parents=True, exist_ok=True)
                stamp = time.strftime("%Y%m%d_%H%M%S")
                head = (html or "")[:8192] if cause != "login_form" else (html or "")[:524288]
                diag = ""
                if cause == "login_form":
                    try:
                        _a = analyze_login_refusal(html, url)
                        diag = (f" markers={len(_a['markers'])}/{len(_a['strong'])}"
                                f" content={_a['content_ok']} title={_a['title']!r}")
                    except Exception:
                        diag = ""
                (d / f"{stamp}_{cause}_{self._forensics_dumped}.html").write_text(
                    f"<!-- url={url} status={status} final={final_url}{diag} -->\n{head}",
                    encoding="utf-8", errors="replace",
                )
        except Exception:
            pass

    def _try_harvest_merge(self) -> bool:
        """Пытается подлить свежие cookies из уже запущенного браузера."""
        h = self.cookie_harvester
        if not h:
            return False
        try:
            fresh = h()
        except Exception:
            return False
        try:
            return bool(self.merge_cookies(fresh or []))
        except Exception:
            return False

    def _fresh_jar(self):
        """Канонический jar из текущих cookies.

        v1.17: requests.Session сохраняет cookies из ответов, и один ответ
        с протухшей сессией навсегда «отравлял» jar потока (все дальнейшие
        быстрые запросы шли гостем → «требуется вход»). Поэтому перед каждым
        запросом jar пересобирается из эталона self._cookies.
        """
        jar = requests.cookies.RequestsCookieJar()
        with self._lock:
            cookies = list(self._cookies)
        for c in cookies:
            try:
                name = (c or {}).get("name")
                if not name:
                    continue
                jar.set(
                    name, (c or {}).get("value", ""),
                    domain=((c or {}).get("domain") or FORUM_HOST),
                    path=(c or {}).get("path") or "/",
                )
            except Exception:
                continue
        return jar

    # ── сессии (по одной на поток, keep-alive) ───────────────
    def _session(self):
        sess = getattr(self._tls, "sess", None)
        rev = getattr(self._tls, "rev", -1)
        with self._lock:
            cur_rev = self._cookie_rev
            ua = self._ua_override or UA
        if sess is None or rev != cur_rev:
            sess = requests.Session()
            sess.headers.update({
                "User-Agent": ua,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,"
                          "image/avif,image/webp,*/*;q=0.8",
                "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
                "Referer": FORUM_BASE_URL,
                "Upgrade-Insecure-Requests": "1",
                "Sec-Fetch-Dest": "document",
                "Sec-Fetch-Mode": "navigate",
                "Sec-Fetch-Site": "same-origin",
                "Sec-Fetch-User": "?1",
            })
            sess.cookies = self._fresh_jar()
            self._tls.sess = sess
            self._tls.rev = cur_rev
        return sess

    # ── главная точка входа (сигнатура как в 1.15) ───────────
    def get(self, url: str, retries: int = 5, emitter=None, parent_id=None):
        try:
            host = (urlparse(url).netloc or "").lower()
        except Exception:
            host = ""
        is_forum = bool(host) and host.endswith(FORUM_HOST.lower())

        # v1.21: одна точка входа — во время общей паузы ждут ОБА канала.
        if is_forum and not self._wait_breaker(emitter=emitter, parent_id=parent_id):
            return None

        if self.fast_enabled and is_forum and self.has_cookies():
            with self._fast_lock:
                browser_only = time.monotonic() < self._fast_browser_only_until
            if not browser_only:
                html = self._get_fast(url, emitter=emitter, parent_id=parent_id)
                if html is not None:
                    return html
                # Быстрый путь не справился — идём в браузер (ниже).

        return self._get_via_browser(url, retries=retries, emitter=emitter, parent_id=parent_id)

    # ── быстрый путь ─────────────────────────────────────────
    def _get_fast(self, url: str, emitter=None, parent_id=None):
        em = emitter
        s_fast = None
        if em and parent_id:
            s_fast = em.step_added(parent_id, "⚡ Быстрый запрос (cookies)", "⚡")
            em.step_state(s_fast, "running")

        def _finish(state, detail=""):
            if em and s_fast:
                em.step_state(s_fast, state, detail)

        last_soft = ""
        max_attempts = FAST_MAX_ATTEMPTS
        recovered = False
        for attempt in range(FAST_MAX_ATTEMPTS + 1):
            if attempt >= max_attempts:
                break
            self._rate.wait()
            time.sleep(random.uniform(FAST_DELAY_MIN, FAST_DELAY_MAX) * self.stealth_mult)
            t0 = time.monotonic()
            try:
                sess = self._session()
                sess.cookies = self._fresh_jar()
                r = sess.get(url, timeout=FAST_TIMEOUT, allow_redirects=True)
            except Exception as e:
                last_soft = type(e).__name__
                get_logger().debug(f"fast: {last_soft} ({attempt + 1}/{max_attempts}): {url}")
                self._record_refusal("error", url, 0, url, "")
                self._rate.report(False)
                self._note_block_signal(weight=1)
                time.sleep(2.0 + random.uniform(0.0, 2.0))
                continue

            elapsed = time.monotonic() - t0
            status = getattr(r, "status_code", 0) or 0
            final_url = getattr(r, "url", url) or url

            if status != 200:
                self._record_refusal(f"http_{status}", url, status, final_url, "")
                self._note_fast_block(f"HTTP {status}")
                if status in (403, 429):
                    self._note_block_signal(weight=2 if status == 429 else 1)
                _finish("error", f"HTTP {status} → браузер")
                return None

            try:
                ctype = ((getattr(r, "headers", {}) or {}).get("Content-Type") or "").lower()
                enc = (getattr(r, "encoding", "") or "").lower().replace("_", "-")
                if (not enc or enc in ("iso-8859-1", "latin-1")) and "charset=" not in ctype:
                    r.encoding = "utf-8"
                html = r.text
            except Exception:
                html = ""

            if _is_login_url(final_url) or _looks_like_login_page(html):
                if _is_login_url(final_url):
                    cause = "redirect_login"
                    detail = "редирект на вход → браузер"
                    analysis = None
                else:
                    cause = "login_form"
                    analysis = analyze_login_refusal(html, url)
                    # v1.20: страница с формой входа, НО с сильными маркерами входа
                    # и ожидаемым контентом — форум отдал СВОЮ страницу, а форма лишь
                    # встроена в тему оформления. Доверяем маркерам, принимаем.
                    if analysis["strong"] and analysis["content_ok"]:
                        self._rate.report(True)
                        with self._fast_lock:
                            self._fast_block_count = 0
                            self._fast_only_cycles = 0
                        self._bump("fast")
                        self._bump("fast_trusted")
                        self._note_clean_hit()
                        _finish("done", f"{len(html) // 1024} КБ (маркеры входа есть, форма проигнорирована)")
                        get_logger().info(
                            f"fast: доверяю маркерам входа ({analysis['strong'][0]}), "
                            f"форма входа проигнорирована: {url}")
                        return html
                    detail = (f"форма входа в HTML ({len(html) // 1024} КБ; "
                              f"{self._cookie_diag()}; маркеры входа: {analysis['mstate']}) → браузер")
                self._record_refusal(cause, url, status, final_url, html)
                if not recovered and self._try_harvest_merge():
                    recovered = True
                    max_attempts += 1
                    _finish("warn", "cookies обновлены, повтор быстрым путём")
                    get_logger().info("fast: cookies обновлены из браузера — повторяю быстрым путём")
                    continue
                self._note_fast_block("требуется вход")
                self._warn_auth_lost(analysis["hint"] if analysis else "")
                _finish("error", detail)
                return None

            if not html_ok(html):
                self._record_refusal("antibot", url, status, final_url, html)
                self._note_fast_block("антибот-маркер")
                self._note_block_signal(weight=1)
                _finish("error", "антибот → браузер")
                return None

            if _suspicious_thread_html(url, html):
                self._record_refusal("suspicious", url, status, final_url, html)
                self._note_fast_block("подозрительный ответ")
                _finish("warn", "проверка через браузер")
                return None

            self._rate.report(True)
            with self._fast_lock:
                self._fast_block_count = 0
                self._fast_only_cycles = 0
            self._bump("fast")
            self._note_clean_hit()
            _finish("done", f"{len(html) // 1024} КБ за {elapsed:.1f}с")
            return html

        self._note_fast_block(last_soft or "недоступно")
        _finish("error", "недоступно → браузер")
        return None

    def _note_fast_block(self, reason: str):
        self._bump("fast_block")
        self._rate.report(False)
        window = 0.0
        with self._fast_lock:
            self._fast_block_count += 1
            n = self._fast_block_count
            if n >= FAST_BLOCK_THRESHOLD:
                self._fast_block_count = 0
                # v1.17: адаптивное окно — каждый бесплодный цикл удваивает
                # паузу (до потолка), успех/свежие cookies сбрасывают.
                window = min(FAST_BROWSER_ONLY_BASE * (2 ** self._fast_only_cycles),
                             FAST_BROWSER_ONLY_MAX)
                self._fast_only_cycles += 1
                self._fast_browser_only_until = time.monotonic() + window
        if window > 0:
            get_logger().warning(
                f"fast: серия блокировок ({reason}) — "
                f"только браузер на {window:.0f}с")

    def _warn_auth_lost(self, hint: str = ""):
        now = time.monotonic()
        with self._fast_lock:
            if now - self._auth_warn_at < 300:
                return
            self._auth_warn_at = now
        extra = f"Разбор: {hint}. " if hint else ""
        get_logger().warning(
            f"fast: форум требует вход ({self._cookie_diag()}) — сессия мертва? "
            f"{extra}"
            "Нажмите «🔑 Войти» заново (с галочкой «Запомнить меня») или "
            "«🍪 Проверить cookies» для диагностики.")

    def verify_fast_session(self, probe_url: str = "") -> dict:
        """Проверяет, узнаёт ли форум сохранённые cookies (v1.19).

        Одноразовый GET главной через свежий requests.Session (TLS-кэш не трогаем).
        Возвращает dict: ok / cause / kb / marker / cookies_n / xf_user / status / final.
        Причины: ok | no_cookies | redirect_login | login_form | antibot |
                 no_markers | http_NNN | error:<имя исключения>.
        """
        url = (probe_url or FORUM_BASE_URL).strip() or FORUM_BASE_URL
        try:
            with self._lock:
                cookies = list(self._cookies)
                ua = self._ua_override or UA
        except Exception:
            cookies = []
            ua = "Mozilla/5.0"
        res = {"ok": False, "cause": "", "kb": 0, "marker": "",
               "cookies_n": len(cookies),
               "xf_user": any(((c or {}).get("name") == "xf_user"
                               and (c or {}).get("value") not in ("", "0", "deleted", None))
                              for c in cookies)}
        if not cookies:
            res["cause"] = "no_cookies"
            return res
        try:
            self._rate.wait()
        except Exception:
            pass
        try:
            sess = requests.Session()
            sess.headers.update({
                "User-Agent": ua,
                "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
                "Accept-Language": "ru-RU,ru;q=0.9,en;q=0.8",
                "Referer": FORUM_BASE_URL,
                "Upgrade-Insecure-Requests": "1",
            })
            sess.cookies = self._fresh_jar()
            r = sess.get(url, timeout=FAST_TIMEOUT, allow_redirects=True)
            status = getattr(r, "status_code", 0) or 0
            final_url = getattr(r, "url", url) or url
            try:
                html = r.text or ""
            except Exception:
                html = ""
        except Exception as e:
            res["cause"] = f"error:{type(e).__name__}"
            return res
        res["kb"] = len(html) // 1024
        res["final"] = final_url
        res["status"] = status
        if status != 200:
            res["cause"] = f"http_{status}"
            return res
        if _is_login_url(final_url):
            res["cause"] = "redirect_login"
            return res
        if _looks_like_login_page(html):
            _a = analyze_login_refusal(html, url)
            res["analysis"] = _a
            if _a["strong"] and _a["content_ok"]:
                res["marker"] = _a["markers"][0] if _a["markers"] else ""
                res["ok"] = True
                res["cause"] = "ok"
                res["note"] = "форма входа проигнорирована: есть маркеры входа"
                return res
            res["cause"] = "login_form"
            return res
        if not html_ok(html):
            res["cause"] = "antibot"
            return res
        found = _login_markers_found(html)
        res["marker"] = found[0] if found else ""
        if found:
            res["ok"] = True
            res["cause"] = "ok"
        else:
            res["cause"] = "no_markers"
        return res

    # ── ленивый старт браузера ───────────────────────────────
    def _ensure_browser(self) -> bool:
        with self._fallback_lock:
            if self.browser_fallback:
                return True
        ensurer = self.browser_ensurer
        if not ensurer:
            return False
        now = time.monotonic()
        with self._fallback_lock:
            if now - self._last_ensure_attempt < 60.0:
                return self.browser_fallback is not None
            self._last_ensure_attempt = now
        try:
            ok = bool(ensurer())
        except Exception as e:
            get_logger().warning(f"browser_ensurer: {e}")
            ok = False
        with self._fallback_lock:
            return bool(self.browser_fallback) or ok

    def _maybe_harvest_cookies(self):
        harvester = self.cookie_harvester
        if not harvester:
            return
        with self._fast_lock:
            need = (self._fast_block_count >= FAST_BLOCK_THRESHOLD
                    or time.monotonic() < self._fast_browser_only_until)
        if not need:
            return
        try:
            fresh = harvester()
        except Exception:
            return
        if self.merge_cookies(fresh or []):
            get_logger().info("cookies обновлены из браузера — быстрый путь восстановлен")
            with self._fast_lock:
                self._fast_block_count = 0
                now = time.monotonic()
                if self._fast_browser_only_until > now + 60.0:
                    self._fast_browser_only_until = now + 60.0

    # ── путь через браузер (логика 1.15 без изменений) ───────
    def _get_via_browser(self, url: str, retries: int = 5, emitter=None, parent_id=None):
        if not self._ensure_browser():
            get_logger().error(f"Нет активного браузера — запрос отклонён: {url}")
            with self._fallback_lock:
                self.no_browser_errors += 1
            return None

        with self._fallback_lock:
            fb = self.browser_fallback

        if not fb:
            get_logger().error(f"Нет активного браузера — запрос отклонён: {url}")
            with self._fallback_lock:
                self.no_browser_errors += 1
            return None

        em = emitter

        while True:
            wait_cooldown = self._cooldown_until - time.monotonic()
            if wait_cooldown <= 0:
                break

            if em and parent_id:
                s_wait = em.step_added(parent_id, f"Пауза антибота {wait_cooldown:.0f}с", "⏸")
                em.step_state(s_wait, "paused")

            time.sleep(min(wait_cooldown, 30.0))

        time.sleep(random.uniform(PAGE_DELAY_MIN, PAGE_DELAY_MAX) * self.stealth_mult)

        for attempt in range(retries):
            try:
                html = fb(url)

                if html and html_ok(html):
                    with self._block_lock:
                        self._block_count = 0
                    with self._fallback_lock:
                        self.no_browser_errors = 0
                    self._bump("browser")
                    self._note_clean_hit()
                    self._maybe_harvest_cookies()
                    return html

                with self._block_lock:
                    self._block_count += 1
                    block_count = self._block_count

                if block_count >= BROWSER_FALLBACK_THRESHOLD:
                    cooldown = 180.0 + random.uniform(0.0, 120.0)
                    # v1.21: серия в браузере останавливает и быстрый канал тоже.
                    self._note_block_signal(weight=2)
                    with self._block_lock:
                        self._block_count = 0
                        self._cooldown_until = time.monotonic() + cooldown

                    get_logger().warning(
                        f"browser: серия блокировок, пауза {cooldown:.1f}с: {url}")

                    if em and parent_id:
                        s_pause = em.step_added(parent_id, f"Пауза антибота {cooldown:.0f}с", "⏸")
                        em.step_state(s_pause, "paused")

                    if not self._sleep_abortable(cooldown):
                        return None
                    continue

                if html and "page could not be loaded" in html.lower():
                    wait = 8.0 * (attempt + 1) + random.uniform(3.0, 10.0)
                    self._note_block_signal(weight=1)
                    get_logger().warning(
                        f"browser: 'Page could not be loaded' ({attempt + 1}/{retries}), ждём {wait:.1f}с: {url}")

                    if em and parent_id:
                        s_retry = em.step_added(
                            parent_id,
                            f"Антибот-ожидание {attempt + 1}/{retries} ({wait:.0f}с)",
                            "⏸"
                        )
                        em.step_state(s_retry, "paused")

                    if not self._sleep_abortable(wait):
                        return None
                    continue

                get_logger().warning(f"browser: пустой/заблокированный ответ ({attempt + 1}/{retries}): {url}")

                if em and parent_id:
                    s_retry = em.step_added(
                        parent_id,
                        f"Повтор {attempt + 1}/{retries} через {5 * (attempt + 1)}с",
                        "⏸"
                    )
                    em.step_state(s_retry, "paused")

                if not self._sleep_abortable(5.0 * (attempt + 1) + random.uniform(2.0, 8.0)):
                    return None

            except Exception as e:
                with self._block_lock:
                    self._block_count += 1

                get_logger().warning(f"browser_fallback.get {url}: {e}")

                if em and parent_id:
                    s_retry = em.step_added(
                        parent_id,
                        f"Ошибка браузера, повтор через {5 * (attempt + 1)}с",
                        "⏸"
                    )
                    em.step_state(s_retry, "paused")

                if not self._sleep_abortable(5.0 * (attempt + 1) + random.uniform(2.0, 8.0)):
                    return None

        return None

    # ── скачивание файлов общей сессией (картинки форума) ────
    def download_bytes(self, url: str, timeout: int = 15, referer: str = "") -> bytes | None:
        """Скачивание файла keep-alive сессией с cookies. Только forum-хост."""
        try:
            host = (urlparse(url).netloc or "").lower()
        except Exception:
            return None
        if not host.endswith(FORUM_HOST.lower()):
            return None
        if not self.has_cookies():
            return None
        if not self._wait_breaker():
            return None
        try:
            self._rate.wait()
            sess = self._session()
            sess.cookies = self._fresh_jar()
            r = sess.get(url, headers={
                "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
                "Referer": referer or FORUM_BASE_URL,
                "Sec-Fetch-Dest": "image",
                "Sec-Fetch-Mode": "no-cors",
                "Sec-Fetch-Site": "same-origin",
            }, timeout=timeout)
            if not r.ok:
                return None
            if not _valid_image_bytes(r.content):
                return None
            return r.content
        except Exception:
            return None


def base_of(url: str) -> str:
    p = urlparse(url)
    return f"{p.scheme}://{p.netloc}"

def abs_url(href: str, base: str) -> str:
    return urljoin(base, href).split("#")[0].rstrip("/")

def plural(n, one: str, few: str, many: str) -> str:
    """Русская plural-форма: 1 пост, 3 поста, 5 постов (v1.17)."""
    try:
        n = abs(int(n))
    except Exception:
        return many
    if 11 <= (n % 100) <= 14:
        return many
    d = n % 10
    if d == 1:
        return one
    if 2 <= d <= 4:
        return few
    return many


def norm_url(url: str) -> str:
    """Ключ для дедупликации: lower, без слеша, без query/fragment."""
    return url.rstrip("/").lower().split("?")[0].split("#")[0]

def norm_thread_url(url: str) -> str:
    """
    Нормализует URL темы для дедупликации.
    Убирает /post-NNN, /unread, query, fragment, trailing slash.
    Пример:
      /threads/foo.123/post-456  → /threads/foo.123
      /threads/foo.123/unread    → /threads/foo.123
    """
    u = url.split("#")[0].split("?")[0]
    # Убираем /post-NNN в конце
    u = re.sub(r'/post-\d+$', '', u)
    # Убираем /unread, /latest и подобные суффиксы
    u = re.sub(r'/(unread|latest|page-\d+)$', '', u)
    return u.rstrip("/").lower()

def is_forum_href(href: str) -> bool:
    """
    Определяет является ли ссылка разделом/категорией форума.
    Исключает якорные ссылки /#...
    """
    if not href or href.startswith("#") or CATEGORY_ANCHOR_RE.match(href):
        return False
    p = urlparse(href).path if href.startswith("http") else href
    return (
        "/forums/"     in p or
        "/categories/" in p or
        "/nodes/"      in p
    )


def in_nav(el: Tag) -> bool:
    NAV = (
        "breadcrumb", "p-breadcrumbs", "p-navEl", "pageNav", "p-nav",
        "headerMobile", "menu", "navigation", "p-footer", "footer", "sidebar",
        "node-controls", "node-extra", "node-stats",
        "structItem-cell--latest", "structItem-cell--icon",
    )
    cur, d = el.parent, 0
    while cur and d < 8:
        if isinstance(cur, Tag):
            cls = " ".join(cur.get("class", []))
            tag = (cur.name or "").lower()
            if any(n in cls for n in NAV) or tag in ("nav", "header", "footer"):
                return True
        cur = cur.parent
        d += 1
    return False


def page_title(soup: BeautifulSoup) -> str:
    for sel in ("h1.p-title-value", "h1.p-title", "h1", ".p-title-value"):
        t = soup.select_one(sel)
        if t:
            txt = t.get_text(strip=True)
            if txt:
                return txt
    if soup.title:
        return soup.title.get_text(strip=True)
    return ""


def title_from_url(url: str) -> str:
    slug = url.rstrip("/").split("/")[-1]
    slug = re.sub(r"\.\d+$", "", slug)
    return slug.replace("-", " ").replace("_", " ").strip().title() or url


CATEGORY_ANCHOR_RE = re.compile(r'^/#', re.IGNORECASE)

SKIP_TITLES = frozenset((
    "отметить прочитанным", "отметить прочтённым", "mark forum read",
    "mark read", "watch forum", "unwatch forum",
    "смотреть форум", "не смотреть форум", "новые сообщения",
    "whats new", "what's new",
))


def _parse_uix_category_title(anchor_tag: Tag) -> str:
    """
    Извлекает ID категории из href="/#name.97" → возвращает "pravila-proyekta.97"
    или пустую строку если не подходит.
    """
    href = anchor_tag.get("href", "")
    m = re.match(r'^/#(.+)$', href)
    return m.group(1) if m else ""


def find_subforums(soup: BeautifulSoup, base: str, debug_cb=None):
    """
    Парсит страницу XenForo с UIX-темой.

    Структура HTML на главной странице:
        uix_categoryTitle  (href="/#category-slug.ID")  →  РАЗДЕЛ верхнего уровня
            .node-main .node-title a (href="/forums/..." или "/categories/...")
                                                         →  ПОДРАЗДЕЛ этого раздела

    Возвращает:
        Если это ГЛАВНАЯ страница (есть uix_categoryTitle):
            список Node-объектов с заполненными children
        Если это страница РАЗДЕЛА (нет uix_categoryTitle):
            список (url, title) дочерних подразделов
    """

    def make_url(href: str) -> str | None:
        """Абсолютный URL или None если ссылка служебная/якорная."""
        if not href or CATEGORY_ANCHOR_RE.match(href):
            return None
        if not is_forum_href(href):
            return None
        u = abs_url(href, base)
        bad, reason = is_bad_url(u)
        if bad:
            if debug_cb and DEBUG_LINKS:
                debug_cb(f"    ⏭ СЛУЖЕБНАЯ ({reason}): {u}")
            return None
        return u

    def clean_title(tag: Tag) -> str:
        return (tag.get_text(strip=True) or "").strip()

    # ════════════════════════════════════════════════════════════
    #  ВАРИАНТ А: Главная страница с uix_categoryTitle
    # ════════════════════════════════════════════════════════════
    category_anchors = soup.find_all(
        "a",
        class_=lambda c: c and "uix_categoryTitle" in c,
    )

    if category_anchors:
        if debug_cb:
            debug_cb(f"  🗂 Режим: Главная страница ({len(category_anchors)} категорий)")

        result = []  # список Node

        for cat_a in category_anchors:
            cat_title = clean_title(cat_a)
            cat_href  = cat_a.get("href", "")

            if not cat_title:
                continue

            if debug_cb and DEBUG_LINKS:
                debug_cb(f"  📌 Категория: '{cat_title}' ({cat_href})")

            # Создаём узел-категорию (без реального URL — это якорь)
            cat_node = Node(cat_href, cat_title)

            # Ищем все .node-title a которые идут ПОСЛЕ этого якоря
            # и ДО следующего uix_categoryTitle
            # Используем позицию в DOM: берём родительский контейнер категории
            # В UIX структура такая:
            #   <div class="uix_categoryBlock"> или похожий враппер
            #     <a class="uix_categoryTitle" href="/#...">Название</a>
            #     <div class="node ...">
            #       <div class="node-main">
            #         <h3 class="node-title"><a href="/forums/...">...</a></h3>
            #       </div>
            #     </div>
            #     ...следующие .node...
            #   </div>

            # Стратегия: ищем родительский блок категории
            cat_block = _find_category_block(cat_a)

            if cat_block:
                for node_title_a in cat_block.select(".node-title a[href]"):
                    href  = node_title_a.get("href", "")
                    title = clean_title(node_title_a)
                    u     = make_url(href)
                    if u and title and title.lower() not in SKIP_TITLES:
                        child = Node(u, title)
                        cat_node.children.append(child)
                        if debug_cb and DEBUG_LINKS:
                            debug_cb(f"    ✅ подраздел '{title}' → {u}")
            else:
                # Fallback: ищем .node после этого якоря до следующего якоря
                for sub_a, sub_title in _collect_siblings_until_next_cat(cat_a, debug_cb):
                    u = make_url(sub_a.get("href", ""))
                    if u and sub_title and sub_title.lower() not in SKIP_TITLES:
                        child = Node(u, sub_title)
                        cat_node.children.append(child)
                        if debug_cb and DEBUG_LINKS:
                            debug_cb(f"    ✅ [sibling] '{sub_title}' → {u}")

            if debug_cb:
                debug_cb(f"    → в категории: {len(cat_node.children)} {plural(len(cat_node.children), 'подраздел', 'подраздела', 'подразделов')}")

            result.append(cat_node)

        return result  # список Node (не tuple!)

    # ════════════════════════════════════════════════════════════
    #  ВАРИАНТ Б: Страница раздела — ищем дочерние подразделы
    # ════════════════════════════════════════════════════════════
    if debug_cb:
        debug_cb(f"  📂 Режим: Страница раздела")

    result, seen = [], set()

    def add(href: str, title: str, source: str = ""):
        u = make_url(href)
        if not u:
            return
        t = title.strip()
        if not t or t.lower() in SKIP_TITLES:
            return
        key = norm_url(u)
        if key in seen:
            return
        seen.add(key)
        result.append((u, t))
        if debug_cb and DEBUG_LINKS:
            debug_cb(f"    ✅ [{source}] '{t}' → {u}")

    # XenForo 2: node-title
    for a in soup.select(".node-title a[href]"):
        add(a.get("href", ""), clean_title(a), "node-title")

    # XenForo 2: structItem
    for a in soup.select(".structItem-title a[href]"):
        add(a.get("href", ""), clean_title(a), "structItem-title")

    # Legacy
    if not result:
        for item in soup.select(".nodeList .node, li.forum, .forum-item"):
            a = item.select_one(".nodeTitle a, h3 a")
            if a:
                add(a.get("href", ""), clean_title(a), "legacy")

    # Fallback
    if not result:
        root = soup.select_one(".p-body-main, .p-body-content") or soup
        for a in root.find_all("a", href=True):
            if in_nav(a):
                continue
            add(a.get("href", ""), clean_title(a), "fallback")

    return result  # список tuple (url, title)


def _find_category_block(cat_anchor: Tag):
    """
    Поднимается вверх по DOM от uix_categoryTitle,
    ищет контейнер который содержит и сам якорь и .node-title ссылки.
    """
    cur = cat_anchor.parent
    depth = 0
    while cur and depth < 6:
        if isinstance(cur, Tag):
            # Проверяем есть ли внутри .node-title ссылки
            if cur.select(".node-title a[href]"):
                return cur
        cur = cur.parent
        depth += 1
    return None


def _collect_siblings_until_next_cat(cat_anchor: Tag, debug_cb=None):
    """
    Собирает .node-title ссылки из сиблингов после cat_anchor
    до следующего uix_categoryTitle.
    Используется как fallback если _find_category_block не нашёл контейнер.
    """
    results = []
    # Идём по родителю и его сиблингам
    parent = cat_anchor.parent
    if not parent:
        return results

    found_self = False
    for sibling in parent.parent.children if parent.parent else []:
        if not isinstance(sibling, Tag):
            continue

        # Нашли себя — начинаем собирать следующих
        if sibling is parent:
            found_self = True
            continue

        if not found_self:
            continue

        # Следующая категория — стоп
        if sibling.find("a", class_=lambda c: c and "uix_categoryTitle" in c):
            break

        # Собираем .node-title из этого сиблинга
        for a in sibling.select(".node-title a[href]"):
            title = a.get_text(strip=True)
            if title:
                results.append((a, title))

    return results


def crawl_all_forum_roots(http: HttpClient, base_url: str, debug_cb=None) -> list:
    """
    Загружает главную страницу форума и собирает разделы ВЕРХНЕГО уровня.
    Возвращает список (url, title) — только прямые потомки главной.
    Глубокий обход делает Scanner рекурсивно.
    """
    html = http.get(base_url)
    if not html:
        return []
    soup = BeautifulSoup(html, BS_PARSER)
    base = base_of(base_url)
    return find_subforums(soup, base, debug_cb)

def parse_thread_item(item: Tag, base: str):
    for sel in (
        ".structItem-title a[data-tp-primary]",
        ".structItem-title a[href*='/threads/']",
        "h3 a[href*='/threads/']",
        "a[href*='/threads/']",
    ):
        a = item.select_one(sel)
        if a and "/threads/" in a.get("href", ""):
            t = a.get_text(strip=True)
            if t:
                return {
                    "title": t,
                    "url": urljoin(base, a["href"]).split("?")[0],
                    "sticky": bool(item.select_one(".structItem-status--sticky")),
                }
    return None

def fallback_threads(soup: BeautifulSoup, base: str):
    seen, res = set(), []
    for a in soup.find_all("a", href=True):
        h = a["href"]
        if "/threads/" not in h or "/latest" in h or "/unread" in h:
            continue
        u = urljoin(base, h).split("?")[0].split("#")[0]
        if u in seen:
            continue
        seen.add(u)
        t = a.get_text(strip=True)
        if len(t) > 3:
            res.append({"title": t, "url": u, "sticky": False})
    return res

def extract_threads(soup: BeautifulSoup, base: str):
    items = [it for it in soup.select(".structItem--thread")
             if "structItem--forum" not in " ".join(it.get("class", []))]
    if items:
        res = [r for r in (parse_thread_item(i, base) for i in items) if r]
        if res:
            return res
    return fallback_threads(soup, base)


def next_page_url(soup: BeautifulSoup, base: str):
    for sel in ("a.pageNav-jump--next", 'a[rel="next"]'):
        t = soup.select_one(sel)
        if t and t.get("href"):
            return urljoin(base, t["href"])
    return None


def collect_threads(http: HttpClient, forum_url: str, max_pages: int = 50):
    base = base_of(forum_url)
    threads, url, page = [], forum_url, 1
    while url and page <= max_pages:
        html = http.get(url)
        if not html:
            break
        soup = BeautifulSoup(html, BS_PARSER)
        threads += extract_threads(soup, base)
        url = next_page_url(soup, base)
        page += 1
    seen, uniq = set(), []
    for t in threads:
        k = t["url"].rstrip("/")
        if k not in seen:
            seen.add(k)
            uniq.append(t)
    return uniq


def extract_posts(soup: BeautifulSoup, base: str = ""):
    out = []
    for a in soup.select("article.message, .js-post, .message--post"):
        author = "—"
        for sel in (".message-name .username", ".username", "[itemprop='name']"):
            t = a.select_one(sel)
            if t:
                author = t.get_text(strip=True)
                break

        date = ""
        t = a.select_one("time")
        if t:
            date = t.get("title") or t.get("datetime") or t.get_text(strip=True)

        body = (a.select_one(".message-body .bbWrapper")
                or a.select_one(".bbWrapper")
                or a.select_one(".message-body"))
        if not body:
            continue

        blocks = body_to_blocks(body, base=base)
        if blocks:
            out.append({"author": author, "date": date, "blocks": blocks})

    return out

def body_to_blocks(body: Tag, base: str = ""):
    out = []
    for el in body.children:
        r = el_to_blocks(el, base=base)
        if r:
            out += r if isinstance(r, list) else [r]
    return out

def _extract_real_image_url(src: str, data_url: str, data_src: str, data_original: str, base: str) -> str:
    """
    Извлекает реальный URL изображения, учитывая XenForo-прокси и data-атрибуты.
    Приоритет источников:
      1. data-url        — самый надёжный (прямая ссылка)
      2. data-original   — lazy-load оригинал
      3. data-src        — lazy-load
      4. src с proxy.php — извлекаем параметр image=
      5. обычный src
    """
    from urllib.parse import parse_qs

    candidates = []

    # 1. data-url — прямая ссылка (XenForo всегда её ставит)
    if data_url:
        candidates.append(data_url)

    # 2. data-original / data-src
    if data_original:
        candidates.append(data_original)
    if data_src:
        candidates.append(data_src)

    # 3. src с proxy.php?image=...
    if src and "proxy.php" in src:
        try:
            parsed = urlparse(src)
            qs = parse_qs(parsed.query)
            if "image" in qs and qs["image"]:
                candidates.append(qs["image"][0])
        except Exception:
            pass

    # 4. Обычный src (если не прокси)
    if src and "proxy.php" not in src:
        candidates.append(src)

    # Возвращаем первый валидный абсолютный URL
    for c in candidates:
        if not c:
            continue
        c = c.strip()
        if not c:
            continue
        if base:
            c = urljoin(base, c)
        if c.startswith("http://") or c.startswith("https://") or c.startswith("//"):
            return c
        # Если base не помог — пробуем добавить базовый URL форума
        if c.startswith("/"):
            return urljoin(FORUM_BASE_URL, c)
    return ""


def el_to_blocks(el, base: str = ""):
    if isinstance(el, NavigableString):
        t = str(el).strip()
        return {"type": "paragraph", "text": t} if t else None

    if not isinstance(el, Tag):
        return None

    tag = el.name.lower()
    cls = el.get("class", [])
    cls_txt = " ".join(cls).lower()

    if tag in ("script", "style", "noscript", "br"):
        return None

    if tag in ("h1", "h2", "h3", "h4", "h5", "h6"):
        t = el.get_text(strip=True)
        return {"type": "heading", "text": t} if t else None

    if tag == "hr":
        return {"type": "divider"}

    # ═══════════════════════════════════════════════════════
    #  ИЗОБРАЖЕНИЯ — полностью переработано под XenForo
    # ═══════════════════════════════════════════════════════
    if tag == "img":
        # Смайлики/эмодзи пропускаем — их не нужно распознавать
        if any(x in cls_txt for x in (
            "smilie", "emoji", "bbcode-smiley", "smiley", "mce-emoticon"
        )):
            return None

        # Совсем маленькие картинки — иконки/смайлики
        try:
            w = int(el.get("width") or 0)
            h = int(el.get("height") or 0)
            if 0 < w <= 40 and 0 < h <= 40:
                return None
        except Exception:
            pass

        src           = (el.get("src")           or "").strip()
        data_url      = (el.get("data-url")      or "").strip()
        data_src      = (el.get("data-src")      or "").strip()
        data_original = (el.get("data-original") or "").strip()
        alt           = (el.get("alt")           or "").strip()

        real_url = _extract_real_image_url(
            src, data_url, data_src, data_original, base
        )

        if real_url:
            # Нормализуем: убираем #fragment
            if real_url.startswith("//"):
                real_url = "https:" + real_url
            return {
                "type": "image",
                "src": real_url.split("#")[0],
                "alt": alt,
                "original_src": src,
            }

        # Если URL извлечь не удалось — возвращаем alt как текст
        if alt and 1 < len(alt) < 200:
            return {"type": "paragraph", "text": f"[ALT: {alt}]"}

        return None

    if tag in ("code", "pre"):
        return {"type": "code", "text": el.get_text()}

    if tag in ("ul", "ol"):
        return [
            {"type": "list_item", "text": li.get_text(strip=True)}
            for li in el.select("li") if li.get_text(strip=True)
        ]

    if tag == "table":
        rows = []
        for tr in el.select("tr"):
            cells = [td.get_text(strip=True) for td in tr.select("td, th")]
            if any(cells):
                rows.append(cells)
        return {"type": "table", "rows": rows} if rows else None

    if any("spoiler" in c.lower() for c in cls):
        ct = el.select_one(".bbCodeBlock-title, .bbCodeSpoiler-button")
        cc = el.select_one(".bbCodeBlock-content, .bbCodeSpoiler-content")
        return {
            "type": "spoiler",
            "title": ct.get_text(strip=True) if ct else "Спойлер",
            "blocks": body_to_blocks(cc, base=base) if cc else [],
        }

    if any("quote" in c.lower() for c in cls):
        ct = el.select_one(".bbCodeBlock-title")
        cc = el.select_one(".bbCodeBlock-content")
        return {
            "type": "quote",
            "author": ct.get_text(strip=True) if ct else "",
            "blocks": body_to_blocks(cc, base=base) if cc else [],
        }

    t = el.get_text(strip=True)
    # Проверяем, есть ли внутри изображения — они не дают текста через get_text(),
    # но должны быть обработаны рекурсивно. Иначе <p><img src="..."></p>
    # вернёт None и картинка будет потеряна.
    has_images = bool(el.find("img"))

    if not t and not has_images:
        return None

    res = []
    for ch in el.children:
        r = el_to_blocks(ch, base=base)
        if r:
            res += r if isinstance(r, list) else [r]

    if res:
        return res
    return {"type": "paragraph", "text": t} if t else None


def parse_thread_content(http: HttpClient, thread: dict, first_post_only: bool, max_pages: int = 100,
                         emitter=None, parent_id=None, initial_html=None):
    base = base_of(thread["url"])
    posts, url, page = [], thread["url"], 1
    em = emitter
    pid = parent_id
    created_date = ""
    fail_count = 0

    while url and page <= max_pages:
        s_page = None
        if em and pid:
            s_page = em.step_added(pid, f"Загрузка страницы {page}", "🌐")
            em.step_state(s_page, "running")

        if initial_html and page == 1:
            html = initial_html
            initial_html = None

            if not html or not html_ok(html):
                fail_count += 1

                if em and s_page:
                    em.step_state(s_page, "error", "недоступно")

                if fail_count >= 3:
                    return {
                        "title": thread["title"],
                        "url": thread["url"],
                        "posts": posts,
                        "ok": False,
                        "created_date": created_date
                    }

                time.sleep(8.0 * fail_count + random.uniform(3.0, 10.0))
                continue

            fail_count = 0

            if em and s_page:
                size_kb = len(html) // 1024
                em.step_state(s_page, "done", f"{size_kb} КБ")
        else:
            html = http.get(url, retries=2, emitter=em, parent_id=s_page)

            if not html or not html_ok(html):
                fail_count += 1

                if em and s_page:
                    em.step_state(s_page, "error", "недоступно")

                if fail_count >= 3:
                    return {
                        "title": thread["title"],
                        "url": thread["url"],
                        "posts": posts,
                        "ok": False,
                        "created_date": created_date
                    }

                time.sleep(8.0 * fail_count + random.uniform(3.0, 10.0))
                continue

            fail_count = 0

            if em and s_page:
                size_kb = len(html) // 1024
                em.step_state(s_page, "done", f"{size_kb} КБ")

        if page == 1 and not created_date:
            created_date = extract_thread_creation_date(html)

        soup = BeautifulSoup(html, BS_PARSER)
        page_posts = extract_posts(soup, base=url)
        posts += page_posts

        if em and s_page:
            em.step_metric(s_page, "detail", f"{len(page_posts)} {plural(len(page_posts), 'пост', 'поста', 'постов')}")

        if first_post_only and posts:
            break

        url = next_page_url(soup, base)
        page += 1

    return {
        "title": thread["title"],
        "url": thread["url"],
        "posts": posts,
        "ok": True,
        "created_date": created_date
    }


class Node:
    def __init__(self, url: str, title: str):
        self.url = url
        self.title = title
        self.children = []
        self.thread_count = 0
        self.threads = []
        self.mark = None
        self.thread_checks = {}   # {thread_url: bool} — выбор отдельных тем

    def to_dict(self):
        return {
            "url": self.url,
            "title": self.title,
            "thread_count": self.thread_count,
            "threads": self.threads,
            "mark": self.mark,
            "thread_checks": self.thread_checks,
            "children": [c.to_dict() for c in self.children],
        }

    @staticmethod
    def from_dict(d):
        n = Node(d["url"], d["title"])
        n.thread_count = d.get("thread_count", 0)
        n.threads = d.get("threads", [])
        n.mark = d.get("mark")
        n.thread_checks = d.get("thread_checks", {})
        n.children = [Node.from_dict(c) for c in d.get("children", [])]
        return n


def flatten_nodes(roots):
    out = []

    def walk(lst):
        for n in lst:
            out.append(n)
            walk(n.children)

    walk(roots)
    return out


def find_marked_nodes(roots, mark):
    found = []

    def walk(n):
        if n.mark == mark:
            found.append(n)
        for c in n.children:
            walk(c)

    for r in roots or []:
        walk(r)
    return found


def gather_subtree_threads(node: Node, is_checked=None):
    """
    Собирает темы из поддерева узла (включая сам узел).
    Если is_checked передан — темы узла добавляются только если is_checked(n) == True,
    но спуск в дочерние узлы происходит всегда (у вложенного раздела может быть своя галочка).
    """
    result = []
    seen = set()

    def walk(n):
        if is_checked is None or is_checked(n):
            for t in n.threads:
                k = t["url"].rstrip("/")
                if k not in seen:
                    seen.add(k)
                    result.append((n, t))
        for c in n.children:
            walk(c)

    walk(node)
    return result

def _merge_key(u: str) -> str:
    """Ключ слияния: якоря /#... — точным URL, остальное — norm_url.

    v1.16: раньше все якорные категории схлопывались norm_url() в один ключ "/"
    и при слиянии получали пометки одного произвольного якоря.
    """

    u = (u or "").strip()
    if u.startswith("#") or u.startswith("/#"):
        return u.lower().rstrip("/")
    return norm_url(u)


def _index_nodes_by_url(nodes):
    """Плоский индекс url → Node по всему дереву (для быстрого поиска при merge)."""
    idx = {}
    def walk(n):
        idx[_merge_key(n.url)] = n
        for c in n.children:
            walk(c)
    for n in nodes or []:
        walk(n)
    return idx


def merge_preserve_state(new_nodes: list, old_nodes: list):
    """
    Переносит mark / thread_checks / threads из старого дерева в новое (по URL).
    Если у нового узла ещё нет детей (Этап 1 их не парсит), а в старом дереве
    у этого узла уже было поддерево (после Этапа 2) — переносим поддерево
    целиком, ничего не теряя.
    """
    old_idx = _index_nodes_by_url(old_nodes)

    def walk(n: Node):
        old = old_idx.get(_merge_key(n.url))
        if old:
            n.mark = old.mark
            n.thread_checks = dict(old.thread_checks)
            if not n.threads and old.threads:
                n.threads = old.threads
                n.thread_count = old.thread_count
            if not n.children and old.children:
                n.children = old.children
                return  # поддерево уже старое и корректное — глубже не лезем
        for c in n.children:
            walk(c)

    for n in new_nodes or []:
        walk(n)


def compute_content_checksum(data: dict) -> str:
    """
    Контрольная сумма по фактическому содержимому темы (посты/блоки),
    а не по сырому HTML. Не зависит от счётчиков просмотров, рекламы,
    времени визита и прочей динамики — только реальный текст постов.
    """
    try:
        payload = json.dumps(data.get("posts", []), ensure_ascii=False, sort_keys=True)
    except Exception:
        payload = str(data.get("posts", []))
    return hashlib.sha256(payload.encode("utf-8")).hexdigest()


def _diff_meta_fallback(old_posts: list, new_posts: list) -> list:
    """Диагностика, когда посты отличаются, а построчных diff нет (v1.17).

    Такое бывает при смене автора/даты, URL картинок, пробелов: checksum
    видит разницу, а нормализованные строки совпадают. Возвращает поясняющие
    строки, чтобы тема не помечалась загадочным «0 изм.».
    """
    out = []
    if len(old_posts) != len(new_posts):
        out.append(f"    МЕТА: постов: {len(old_posts)} → {len(new_posts)} (текст совпал)")
        return out
    for i, (op, np_) in enumerate(zip(old_posts, new_posts)):
        if not isinstance(op, dict) or not isinstance(np_, dict):
            if op != np_:
                out.append(f"[Пост #{i+1}] МЕТА: техническое изменение (текст совпал)")
            continue
        if op == np_:
            continue
        marks = []
        if (op.get("author") or "") != (np_.get("author") or ""):
            marks.append(f"автор: «{op.get('author', '?')}» → «{np_.get('author', '?')}»")
        if (op.get("date") or "") != (np_.get("date") or ""):
            marks.append(f"дата: «{op.get('date', '?')}» → «{np_.get('date', '?')}»")
        ob, nb = op.get("blocks", []), np_.get("blocks", [])
        if isinstance(ob, list) and isinstance(nb, list) and len(ob) != len(nb):
            marks.append(f"блоков: {len(ob)} → {len(nb)}")
        elif isinstance(ob, list) and isinstance(nb, list):
            osrc = sorted(b.get("src", "") for b in ob
                          if isinstance(b, dict) and b.get("type") == "image")
            nsrc = sorted(b.get("src", "") for b in nb
                          if isinstance(b, dict) and b.get("type") == "image")
            if osrc != nsrc:
                marks.append(f"картинки: состав изменился ({len(osrc)} → {len(nsrc)})")
        if marks:
            out.append(f"[Пост #{i+1}] МЕТА: " + "; ".join(marks))
        else:
            out.append(f"[Пост #{i+1}] МЕТА: техническое изменение (текст совпал)")
    if not out:
        out.append("    МЕТА: техническое изменение структуры (текст совпал)")
    return out


def diff_thread_blocks(old_data: dict, new_data: dict) -> list:
    """Построчный diff двух версий темы. Возвращает список строк для отчёта/docx."""
    old_posts = old_data.get("posts", []) if old_data else []
    new_posts = new_data.get("posts", []) if new_data else []
    out = []
    for i in range(max(len(old_posts), len(new_posts))):
        old_p = old_posts[i] if i < len(old_posts) else None
        new_p = new_posts[i] if i < len(new_posts) else None
        if old_p is None:
            out.append(f"[Пост #{i+1}] ДОБАВЛЕН (автор: {new_p.get('author','?')})")
            continue
        if new_p is None:
            out.append(f"[Пост #{i+1}] УДАЛЁН (был автор: {old_p.get('author','?')})")
            continue
        old_lines = blocks_to_lines(old_p.get("blocks", []))
        new_lines = blocks_to_lines(new_p.get("blocks", []))
        if old_lines == new_lines:
            continue
        chunk = [f"[Пост #{i+1}, автор: {new_p.get('author','?')}] изменения:"]
        sm = difflib.SequenceMatcher(None, old_lines, new_lines)
        for tag, i1, i2, j1, j2 in sm.get_opcodes():
            if tag == "equal":
                continue
            if tag == "replace":
                chunk.append(f"    БЫЛО:  {' | '.join(old_lines[i1:i2])}")
                chunk.append(f"    СТАЛО: {' | '.join(new_lines[j1:j2])}")
            elif tag == "delete":
                chunk.append(f"    УДАЛЕНО: {' | '.join(old_lines[i1:i2])}")
            elif tag == "insert":
                chunk.append(f"    ДОБАВЛЕНО: {' | '.join(new_lines[j1:j2])}")
        if len(chunk) > 1:
            out.extend(chunk)
    if not out and (old_posts != new_posts):
        # v1.17: checksum различается, а строковых diff нет — поясняем.
        out.extend(_diff_meta_fallback(old_posts, new_posts))
    return out

# ── v1.17: журнал изменений тем ──────────────────────────────────
HISTORY_TS_FMT = "%Y-%m-%d %H:%M:%S"
HISTORY_KIND_ICON = {"new": "🆕", "changed": "♻️", "found": "🔎", "gone": "⚠"}
HISTORY_KIND_LABEL = {
    "new": "Первый парсинг",
    "changed": "Изменилась",
    "found": "Найдена при Этапе 2",
    "gone": "Исчезла из раздела",
}


def history_now_ts() -> str:
    return datetime.now().strftime(HISTORY_TS_FMT)


def history_path(project_dir) -> Path:
    return Path(project_dir) / HISTORY_FILE


def load_history(project_dir) -> dict:
    """Читает history.json. Отсутствует/бит — пустой журнал."""
    try:
        p = history_path(project_dir)
        if not p.exists():
            return {"topics": {}}
        data = json.loads(p.read_text(encoding="utf-8"))
        if isinstance(data, dict) and isinstance(data.get("topics"), dict):
            return data
        if isinstance(data, dict):
            return {"topics": data}
    except Exception:
        pass
    return {"topics": {}}


def save_history(project_dir, data: dict) -> bool:
    try:
        history_path(project_dir).write_text(
            json.dumps(data or {"topics": {}}, ensure_ascii=False, indent=1),
            encoding="utf-8",
        )
        return True
    except Exception:
        return False


def _history_sig(kind: str, checksum: str) -> str:
    return f"{kind}|{checksum or ''}"


def history_append(data: dict, url: str, kind: str, title: str = "",
                   section: str = "", checksum: str = "", diff=None,
                   ts: str = None, first_url: str = "", scopes=None) -> bool:
    """Добавляет событие в журнал. False — дубль (не добавлено).

    Дедуп: подряд два одинаковых события (вид + checksum) не пишутся.
    Ограничения: событий на тему — HISTORY_MAX_EVENTS (старые вытесняются),
    строк diff — HISTORY_MAX_DIFF_LINES, длина строки — HISTORY_MAX_LINE_LEN.
    """
    try:
        key = norm_thread_url(url or "")
    except Exception:
        key = (url or "").strip().lower()
    if not key:
        return False
    if not isinstance(data, dict):
        return False
    topics = data.setdefault("topics", {})
    t = topics.setdefault(key, {"title": "", "url": "", "updated": "", "events": [], "scopes": []})
    if scopes:
        known = t.setdefault("scopes", [])
        for _s in scopes:
            _s = str(_s or "").strip()
            if _s and _s not in known:
                known.append(_s)
                if len(known) > 20:
                    del known[:len(known) - 20]
    if title:
        t["title"] = title
    if not t.get("url"):
        t["url"] = first_url or url or ""
    sig = _history_sig(kind, checksum)
    evts = t.setdefault("events", [])
    if evts:
        last = evts[-1] or {}
        if _history_sig(last.get("kind", ""), last.get("checksum", "")) == sig:
            return False
    lines = []
    for ln in (diff or [])[:HISTORY_MAX_DIFF_LINES]:
        s = str(ln)
        if len(s) > HISTORY_MAX_LINE_LEN:
            s = s[:HISTORY_MAX_LINE_LEN] + "…"
        lines.append(s)
    if diff and len(diff) > HISTORY_MAX_DIFF_LINES:
        lines.append(f"… обрезано, всего строк: {len(diff)}")
    stamp = ts or history_now_ts()
    evts.append({
        "ts": stamp, "kind": kind, "title": title or t.get("title", ""),
        "section": section or "", "checksum": checksum or "", "diff": lines,
        "scopes": [str(s) for s in (scopes or []) if str(s or "").strip()],
    })
    if len(evts) > HISTORY_MAX_EVENTS:
        del evts[:len(evts) - HISTORY_MAX_EVENTS]
    t["updated"] = stamp
    return True


def history_sorted_events(topic: dict, newest_first: bool = True) -> list:
    evts = list((topic or {}).get("events", []))
    if newest_first:
        evts.reverse()
    return evts


def history_event_segments(ev: dict) -> list:
    """Событие → [(текст, тег)] для вставки в Text (чистая функция)."""
    kind = (ev or {}).get("kind", "")
    ts = (ev or {}).get("ts", "")
    icon = HISTORY_KIND_ICON.get(kind, "•")
    label = HISTORY_KIND_LABEL.get(kind, kind)
    seg = [(f"{icon} {ts} — {label}", "h_date")]
    section = (ev or {}).get("section", "")
    if section:
        seg.append((f"    Раздел: {section}", "ctx"))
    _scopes = [s for s in ((ev or {}).get("scopes") or []) if s]
    if _scopes:
        _names = []
        for _s in _scopes:
            if _s == "general":
                _names.append("Общие правила")
            elif str(_s).startswith("server:"):
                _names.append(f"Сервер «{str(_s)[7:]}»")
            else:
                _names.append(str(_s))
        seg.append((f"    Слой: {'; '.join(_names)}", "ctx"))
    if kind == "new":
        seg.append(("    Первый парсинг темы, содержимое сохранено в кэш.", "ctx"))
    elif kind == "found":
        seg.append(("    Тема впервые найдена при сканировании (Этап 2).", "ctx"))
    elif kind == "gone":
        seg.append(("    Тема была выбрана, но не найдена при повторном сканировании.", "ctx"))
    for ln in (ev or {}).get("diff", []):
        s = str(ln)
        st = s.strip()
        if st.startswith("БЫЛО:"):
            seg.append((s, "was"))
        elif st.startswith("СТАЛО:"):
            seg.append((s, "became"))
        elif st.startswith("ДОБАВЛЕНО:"):
            seg.append((s, "add"))
        elif st.startswith("УДАЛЕНО:"):
            seg.append((s, "dlt"))
        elif st.startswith("[Пост"):
            seg.append((s, "post"))
        else:
            seg.append((s, "ctx"))
    return seg


class Scanner:
    """
    Этап 1: scan_top_level   — только главная страница, плоский список разделов
    Этап 2: scan_deep        — DFS внутрь переданных узлов, только названия тем
    """

    def __init__(self, http: HttpClient, executor: ThreadPoolExecutor,
                 abort_event: threading.Event,
                 log_cb=None, progress_cb=None, debug_cb=None,
                 emitter=None, parent_id=None, project_dir=None):
        self.http = http
        self.executor = executor
        self.abort = abort_event
        self.log = log_cb or (lambda s: None)
        self.progress = progress_cb or (lambda s: None)
        self.debug = debug_cb or (lambda s: None)
        self.emitter = emitter
        self.emitter_parent = parent_id
        self.project_dir = project_dir or SERVERS_ROOT
        self._visited = set()
        self._vlock = threading.Lock()
        self.disappeared_selected_threads = []
        self.new_threads_found = []

    # ──────────────────────────────────────────────
    #  Внутренние утилиты
    # ──────────────────────────────────────────────

    def _try_visit(self, url: str) -> bool:
        key = norm_url(url)
        with self._vlock:
            if key in self._visited:
                return False
            self._visited.add(key)
            return True

    def _fetch_soup(self, url: str, emitter=None, parent_id=None):
        em = emitter or self.emitter
        pid = parent_id or self.emitter_parent
        attempts = 0
        max_attempts = 5

        while True:
            if self.abort.is_set():
                return None

            attempts += 1
            html = self.http.get(url, retries=2, emitter=em, parent_id=pid)

            if html and html_ok(html):
                return BeautifulSoup(html, BS_PARSER)

            if attempts >= max_attempts:
                self.log(f"    ❌ Страница недоступна после {attempts} попыток: {url}")
                return None

            wait = min(10.0 * attempts + random.uniform(5.0, 20.0), 180.0)
            self.log(f"    ⚠ Страница недоступна, повтор через {wait:.0f}с...: {url}")

            if em and pid:
                s_wait = em.step_added(pid, f"Повтор через {wait:.0f}с", "⏸")
                em.step_state(s_wait, "paused")

            time.sleep(wait)

    def _diagnose_page(self, soup: BeautifulSoup, url: str):
        """Выводит диагностику найденных ссылок на странице."""
        self.debug(f"  🔬 Диагностика страницы: {url}")

        # Считаем все ссылки по типам
        all_links = soup.find_all("a", href=True)
        forum_links   = [a for a in all_links if is_forum_href(a.get("href",""))]
        anchor_links  = [a for a in all_links if CATEGORY_ANCHOR_RE.match(a.get("href",""))]
        thread_links  = [a for a in all_links if "/threads/" in a.get("href","")]

        self.debug(f"    Всего ссылок на странице: {len(all_links)}")
        self.debug(f"    Якоря-категории (/#...):  {len(anchor_links)}")
        self.debug(f"    Форумы/категории:          {len(forum_links)}")
        self.debug(f"    Темы (/threads/):          {len(thread_links)}")

        # Показываем якорные категории — это группировщики
        if anchor_links:
            self.debug("  📌 Найденные группы (якоря):")
            for a in anchor_links[:20]:
                self.debug(f"      [{a.get('class','')}] "
                           f"'{a.get_text(strip=True)}' → {a['href']}")

        # Показываем реальные разделы
        if forum_links:
            self.debug("  📁 Найденные разделы:")
            for a in forum_links[:30]:
                self.debug(f"      '{a.get_text(strip=True)}' → {a['href']}")

        # CSS-классы на странице для отладки
        node_classes = set()
        for tag in soup.find_all(class_=True)[:200]:
            for c in tag.get("class", []):
                if any(k in c.lower() for k in ("node", "forum", "categor", "uix")):
                    node_classes.add(c)
        if node_classes:
            self.debug(f"  🎨 Классы узлов: {sorted(node_classes)}")

    def _emergency_fallback(self, soup: BeautifulSoup, base: str) -> list:
        """
        Аварийный сбор разделов если основной парсер не нашёл ничего.
        Берёт ВСЕ ссылки вида /forums/... и /categories/... с страницы.
        """
        result, seen = [], set()
        for a in soup.find_all("a", href=True):
            href = a.get("href", "")
            if CATEGORY_ANCHOR_RE.match(href):
                continue
            if not is_forum_href(href):
                continue
            u = abs_url(href, base)
            bad, _ = is_bad_url(u)
            if bad:
                continue
            t = a.get_text(strip=True)
            if not t or len(t) < 2:
                continue
            key = norm_url(u)
            if key in seen:
                continue
            seen.add(key)
            result.append((u, t))
        return result

    # ──────────────────────────────────────────────
    #  ЭТАП 1: плоский список разделов верхнего уровня
    # ──────────────────────────────────────────────

    def scan_top_level(self) -> list:
        self._visited = set()
        base = base_of(FORUM_BASE_URL)
        main_url = FORUM_BASE_URL.rstrip("/")
        em = self.emitter
        pid = self.emitter_parent
        self.log("=" * 60)
        self.log("🌐 ЭТАП 1: Загрузка главной страницы")
        self.log(f"   {main_url}")
        self.log("=" * 60)
        self.progress("Загрузка главной страницы...")
        self._try_visit(main_url)
        s_load = None
        if em and pid:
            s_load = em.step_added(pid, "Загрузка главной страницы", "🌐")
            em.step_state(s_load, "running")
        soup = self._fetch_soup(main_url, emitter=em, parent_id=s_load)
        if soup is None:
            self.log("❌ Главная страница недоступна")
            if em and s_load:
                em.step_state(s_load, "error")
            return []
        if em and s_load:
            em.step_state(s_load, "done")
        s_diag = None
        if em and pid:
            s_diag = em.step_added(pid, "Диагностика страницы", "🔬")
            em.step_state(s_diag, "running")
        self._diagnose_page(soup, main_url)
        if em and s_diag:
            em.step_state(s_diag, "done")
        s_parse = None
        if em and pid:
            s_parse = em.step_added(pid, "Парсинг категорий", "🗂")
            em.step_state(s_parse, "running")
        raw = find_subforums(soup, base, self.debug)
        if raw and isinstance(raw[0], Node):
            roots = []
            for cat_node in raw:
                if self.abort.is_set():
                    break
                self._try_visit(cat_node.url)
                for child in cat_node.children:
                    self._try_visit(child.url)
                roots.append(cat_node)
                self.log(f"  🌐 {cat_node.title} ({len(cat_node.children)} {plural(len(cat_node.children), 'подраздел', 'подраздела', 'подразделов')})")
                for child in cat_node.children:
                    self.log(f"      📁 {child.title}")
                if em and s_parse:
                    s_cat = em.step_added(s_parse, f"📌 {cat_node.title}", "📌")
                    em.step_state(s_cat, "done", f"{len(cat_node.children)} {plural(len(cat_node.children), 'подраздел', 'подраздела', 'подразделов')}")
                    for child in cat_node.children:
                        s_sub = em.step_added(s_cat, f"📁 {child.title}", "📁")
                        em.step_state(s_sub, "done")
        else:
            roots = []
            for item in raw:
                if self.abort.is_set():
                    break
                if isinstance(item, tuple):
                    url, title = item
                else:
                    continue
                if not self._try_visit(url):
                    continue
                node = Node(url, title)
                roots.append(node)
                self.log(f"  📁 {title}")
                if em and s_parse:
                    s_node = em.step_added(s_parse, f"📁 {title}", "📁")
                    em.step_state(s_node, "done")
        if em and s_parse:
            em.step_state(s_parse, "done", f"{len(roots)} категорий")
        if not roots:
            self.log("  ⚠ Разделов не найдено! Запускаю аварийный fallback...")
            s_fb = None
            if em and pid:
                s_fb = em.step_added(pid, "Аварийный fallback", "🆘")
                em.step_state(s_fb, "running")
            fallback = self._emergency_fallback(soup, base)
            for url, title in fallback:
                if not self._try_visit(url):
                    continue
                roots.append(Node(url, title))
                self.log(f"  🆘 {title}")
            if em and s_fb:
                em.step_state(s_fb, "done", f"{len(fallback)} {plural(len(fallback), 'раздел', 'раздела', 'разделов')}")
        self.log(f"\n✅ Этап 1 завершён.")
        self.log(f"   Категорий верхнего уровня: {len(roots)}")
        _tot_sub = sum(len(n.children) for n in roots)
        self.log(f"   Всего: {_tot_sub} {plural(_tot_sub, 'подраздел', 'подраздела', 'подразделов')}")
        return roots

    # ──────────────────────────────────────────────
    #  ЭТАП 2: DFS вглубь выбранных разделов
    # ──────────────────────────────────────────────

    def scan_deep(self, selected_nodes: list) -> None:
        self.log("")
        self.log("=" * 60)
        self.log(f"🔍 ЭТАП 2: Глубокое сканирование {len(selected_nodes)} {plural(len(selected_nodes), 'раздел', 'раздела', 'разделов')}")
        self.log("=" * 60)
        _ex = [str(w) for w in (_autocheck_thread_exclude or []) if str(w or "").strip()]
        if _ex:
            self.log(f"🚫 Активные исключения ({len(_ex)}): {', '.join(_ex[:8])}{' …' if len(_ex) > 8 else ''}")
        else:
            self.log("⚠ Исключений нет — отмечаются все совпадения ключевых слов")
        total_done = [0]
        # v1.17: total — всё известное поддерево (раньше только корни → 1000%+).
        total_nodes = [sum(len(flatten_nodes([n])) for n in selected_nodes)]
        em = self.emitter
        root_pid = self.emitter_parent

        def dfs(node: Node, depth: int, emitter_pid=None):
            if self.abort.is_set():
                return
            indent = "  " * depth
            icons = ["📁", "📂", "📋", "📄"]
            icon = icons[min(depth, len(icons) - 1)]
            self.log(f"{indent}{icon} {node.title}")
            self.progress(f"Сканирую: {node.title[:50]}")
            s_node = None
            if em and emitter_pid:
                s_node = em.step_added(emitter_pid, f"{icon} {node.title}", "📂")
                em.step_state(s_node, "running")
            soup = self._fetch_soup(node.url, emitter=em, parent_id=s_node)
            if soup is None:
                if em and s_node:
                    em.step_state(s_node, "error", "страница недоступна")
                return
            real = page_title(soup)
            if real:
                node.title = real
            base = base_of(node.url)
            s_subs = None
            if em and s_node:
                s_subs = em.step_added(s_node, "Поиск подразделов", "📁")
                em.step_state(s_subs, "running")
            raw_subs = find_subforums(soup, base, self.debug)
            sub_pairs = []
            for item in raw_subs:
                if isinstance(item, Node):
                    sub_pairs.append((item.url, item.title))
                elif isinstance(item, tuple):
                    sub_pairs.append(item)
            existing = {norm_url(ch.url): ch for ch in node.children}
            new_children = []
            for sf_url, sf_title in sub_pairs:
                if self.abort.is_set():
                    break
                key = norm_url(sf_url)
                if key in existing:
                    existing[key].title = sf_title
                    continue
                if not self._try_visit(sf_url):
                    continue
                child = Node(sf_url, sf_title)
                node.children.append(child)
                existing[key] = child
                new_children.append(child)
                total_nodes[0] += 1
            if em and s_subs:
                em.step_state(s_subs, "done", f"{len(sub_pairs)} найдено")
            s_threads = None
            if em and s_node:
                s_threads = em.step_added(s_node, "Сбор тем", "📄")
                em.step_state(s_threads, "running")
            old_threads = node.threads or []
            old_urls = {norm_thread_url(t["url"]) for t in old_threads}
            old_checked_urls = {
                u for u in old_urls
                if node.thread_checks.get(
                    next((t["url"] for t in old_threads if norm_thread_url(t["url"]) == u), ""),
                    False
                )
            }
            threads = self._collect_thread_titles(soup, base, node.url, emitter=em, parent_id=s_threads)
            new_urls = {norm_thread_url(t["url"]) for t in threads}
            thread_dates = load_thread_dates_cached(self.project_dir)
            cfg_kw = _autocheck_thread_keywords
            for t in threads:
                thr_url = t.get("url", "")
                thr_title = t.get("title", "")
                cached_date = thread_dates.get(norm_thread_url(thr_url), '')
                kw_eval = evaluate_thread_keyword_match(thr_title, cached_date, cfg_kw)
                if kw_eval == 'uncheck':
                    node.thread_checks[thr_url] = False
                elif (kw_eval == 'check' or should_autocheck_new_thread(thr_title)) \
                        and not is_excluded_by_words(thr_title, node.title, _autocheck_thread_exclude):
                    node.thread_checks[thr_url] = True
            disappeared = old_checked_urls - new_urls
            if disappeared:
                for t in old_threads:
                    if norm_thread_url(t["url"]) not in disappeared:
                        continue
                    title = t["title"]
                    self.log(
                        f"{indent}  ⚠ ВНИМАНИЕ: ранее выбранная тема пропала "
                        f"после повторного сканирования: «{title}»"
                    )
                    self.disappeared_selected_threads.append((node.title, title, t["url"]))
            appeared = new_urls - old_urls
            if appeared:
                appeared_threads = [
                    t for t in threads
                    if norm_thread_url(t["url"]) in appeared
                ]
                self.log(f"{indent}  🆕 {len(appeared_threads)} {plural(len(appeared_threads), 'новая тема', 'новые темы', 'новых тем')}")
                autochecked_count = 0
                thread_dates = load_thread_dates_cached(self.project_dir)
                cfg_kw = _autocheck_thread_keywords
                for t in appeared_threads:
                    title = t.get("title", "")
                    cached_date = thread_dates.get(norm_thread_url(t["url"]), '')
                    kw_eval = evaluate_thread_keyword_match(title, cached_date, cfg_kw)
                    if kw_eval == 'uncheck':
                        node.thread_checks[t["url"]] = False
                        self.log(f"{indent}    ⏭ Пропущена (дата до фильтра): {title}")
                    elif kw_eval == 'check' or should_autocheck_new_thread(title):
                        if is_excluded_by_words(title, node.title, _autocheck_thread_exclude):
                            node.thread_checks[t["url"]] = False
                            self.log(f"{indent}    ⏭ Пропущена (слово-исключение): {title}")
                        else:
                            node.thread_checks[t["url"]] = True
                            autochecked_count += 1
                            self.log(f"{indent}    ✅ Автоотметка темы: {title}")
                    else:
                        node.thread_checks[t["url"]] = False
                        self.log(f"{indent}    + {title}")
                    self.new_threads_found.append((node.title, title, t["url"]))
                if autochecked_count:
                    self.log(
                        f"{indent}    ✅ Автоматически отмечено новых тем: "
                        f"{autochecked_count}"
                    )
            node.threads = threads
            node.thread_count = len(threads)
            autochecked_existing = 0
            unchecked_by_date = 0
            unchecked_by_exclude = 0
            thread_dates = load_thread_dates_cached(self.project_dir)
            cfg_kw = _autocheck_thread_keywords
            for t in threads:
                thr_url = t.get("url", "")
                thr_title = t.get("title", "")
                cached_date = thread_dates.get(norm_thread_url(thr_url), '')
                kw_eval = evaluate_thread_keyword_match(thr_title, cached_date, cfg_kw)
                node_match = should_autocheck_new_node(thr_title)
                if is_excluded_by_words(thr_title, node.title, _autocheck_thread_exclude):
                    # v1.22: исключение бьёт всё — снимаем и старые, и ручные галочки.
                    if node.thread_checks.get(thr_url, False):
                        node.thread_checks[thr_url] = False
                        unchecked_by_exclude += 1
                elif kw_eval == 'uncheck':
                    if node.thread_checks.get(thr_url, False):
                        node.thread_checks[thr_url] = False
                        unchecked_by_date += 1
                elif kw_eval == 'check' or node_match:
                    if not node.thread_checks.get(thr_url, False):
                        node.thread_checks[thr_url] = True
                        autochecked_existing += 1
            if autochecked_existing:
                self.log(f"{indent}    ✅ Автоотметка существующих тем по ключевым словам: {autochecked_existing}")
            if unchecked_by_date:
                self.log(f"{indent}    ⏭ Снято отметок (дата до фильтра): {unchecked_by_date}")
            if unchecked_by_exclude:
                self.log(f"{indent}    ⏭ Снято отметок (исключения): {unchecked_by_exclude}")
            if threads:
                self.log(f"{indent}  📄 Тем: {len(threads)}")
            if em and s_threads:
                em.step_state(s_threads, "done", f"{len(threads)} {plural(len(threads), 'тема', 'темы', 'тем')}")
            total_done[0] += 1
            self.progress(f"Обработано: {total_done[0]}/{total_nodes[0]}")
            if em and root_pid:
                em.step_progress(root_pid, total_done[0], total_nodes[0])
            for child in node.children:
                if self.abort.is_set():
                    break
                dfs(child, depth + 1, emitter_pid=s_node)
            if em and s_node:
                em.step_state(s_node, "done", f"{len(threads)} {plural(len(threads), 'тема', 'темы', 'тем')}, {len(new_children)} {plural(len(new_children), 'новый подраздел', 'новых подраздела', 'новых подразделов')}")

        for node in selected_nodes:
            if self.abort.is_set():
                break
            self.log(f"\n▶ Начинаю раздел: «{node.title}»")
            dfs(node, depth=0, emitter_pid=root_pid)
        all_nodes = []
        for n in selected_nodes:
            all_nodes.extend(flatten_nodes([n]))
        self.log("")
        self.log(f"✅ Этап 2 завершён.")
        self.log(f"   Обработано: {len(all_nodes)} {plural(len(all_nodes), 'раздел', 'раздела', 'разделов')}")
        self.log(f"   Тем найдено:         {sum(n.thread_count for n in all_nodes)}")
        if self.disappeared_selected_threads:
            self.log(f"   ⚠ {len(self.disappeared_selected_threads)} {plural(len(self.disappeared_selected_threads), 'пропавшая выбранная тема', 'пропавшие выбранные темы', 'пропавших выбранных тем')}")
        if self.new_threads_found:
            self.log(f"   🆕 {len(self.new_threads_found)} {plural(len(self.new_threads_found), 'новая найденная тема', 'новые найденные темы', 'новых найденных тем')}")

    def _collect_thread_titles(self, soup: BeautifulSoup,
                               base: str, forum_url: str,
                               emitter=None, parent_id=None) -> list:
        threads = []
        seen = set()
        cur_soup = soup
        url = forum_url
        page = 1
        em = emitter or self.emitter
        pid = parent_id or self.emitter_parent
        while cur_soup and page <= 50:
            s_page = None
            if em and pid:
                s_page = em.step_added(pid, f"Страница {page}", "📄")
                em.step_state(s_page, "running")
            page_threads = extract_threads(cur_soup, base)
            for t in page_threads:
                k = norm_url(t["url"])
                if k not in seen:
                    seen.add(k)
                    threads.append(t)
            if em and s_page:
                em.step_state(s_page, "done", f"{len(page_threads)} {plural(len(page_threads), 'тема', 'темы', 'тем')}")
            next_url = next_page_url(cur_soup, base)
            if not next_url or next_url == url:
                break
            url = next_url
            page += 1
            if self.http.fast_enabled:
                time.sleep(random.uniform(FAST_DELAY_MIN, FAST_DELAY_MAX) * self.http.stealth_mult)
            else:
                time.sleep(random.uniform(PAGE_DELAY_MIN, PAGE_DELAY_MAX) * self.http.stealth_mult)
            html = self.http.get(url, emitter=em, parent_id=pid)
            if not html:
                break
            cur_soup = BeautifulSoup(html, BS_PARSER)
        return threads

def safe_name(s: str, n=90) -> str:
    s = re.sub(r'[<>:"/\\|?*\x00-\x1f]', "", s or "")
    s = re.sub(r"\s+", " ", s).strip(" .")
    return (s[:n].rstrip() + "…" if len(s) > n else s) or "untitled"


def blocks_to_lines(blocks, ind=0, skip_text=False, skip_images=False, ocr_min=0):
    """
    Преобразует список блоков поста в список текстовых строк.

    Параметры:
        blocks:
            Список блоков.
        ind:
            Уровень вложенности. Используется только для рекурсии.
        skip_text:
            Если True — текстовые блоки пропускаются.
        skip_images:
            Если True — изображения/OCR-текст пропускаются.
        ocr_min:
            Минимальная длина распознанного текста изображения.
            Если OCR-текст короче — изображение не добавляется.
    """
    lines = []

    # Защита от некорректного формата входных данных
    if isinstance(blocks, dict):
        blocks = [blocks]
    elif not isinstance(blocks, list):
        return lines

    skip_text = bool(skip_text)
    skip_images = bool(skip_images)

    try:
        ocr_min = int(ocr_min or 0)
    except Exception:
        ocr_min = 0

    for b in blocks:
        if not isinstance(b, dict):
            continue

        t = (b.get("type") or "").strip()

        # ─────────────────────────────────────────────
        # Заголовки
        # ─────────────────────────────────────────────
        if t == "heading":
            if skip_text:
                continue

            v = (b.get("text") or "").strip()

            if v and not is_garbage_line(v):
                lines.append(normalize_text(v))

        # ─────────────────────────────────────────────
        # Обычные абзацы
        # ─────────────────────────────────────────────
        elif t == "paragraph":
            # Если в будущем OCR-текст будет помечаться отдельным флагом,
            # его можно будет включать даже при skip_text=True,
            # если разрешены изображения.
            is_image_text = bool(b.get("from_image") or b.get("is_ocr"))

            if skip_text and not (is_image_text and not skip_images):
                continue

            v = (b.get("text") or "").strip()

            if v and not is_garbage_line(v):
                lines.append(normalize_text(v))

        # ─────────────────────────────────────────────
        # Списки
        # ─────────────────────────────────────────────
        elif t == "list_item":
            if skip_text:
                continue

            v = (b.get("text") or "").strip()

            if v and not is_garbage_line(v):
                lines.append(normalize_text("- " + v))

        # ─────────────────────────────────────────────
        # Цитата
        #
        # Важно:
        # раньше при skip_text=True цитата пропускалась целиком,
        # из-за чего изображения внутри цитаты тоже терялись.
        # Теперь текст цитаты можно пропустить, но рекурсивно
        # обработать вложенные блоки, чтобы найти изображения.
        # ─────────────────────────────────────────────
        elif t == "quote":
            au = (b.get("author") or "").strip()

            if not skip_text:
                label = f"Цитата ({au}):" if au else "Цитата:"

                if not is_garbage_line(label):
                    lines.append(normalize_text(label))

            lines += blocks_to_lines(
                b.get("blocks", []),
                ind + 1,
                skip_text,
                skip_images,
                ocr_min
            )

        # ─────────────────────────────────────────────
        # Спойлер
        #
        # Аналогично цитате:
        # не пропускаем весь спойлер целиком, если включены изображения.
        # ─────────────────────────────────────────────
        elif t == "spoiler":
            lines += blocks_to_lines(
                b.get("blocks", []),
                ind + 1,
                skip_text,
                skip_images,
                ocr_min
            )

        # ─────────────────────────────────────────────
        # Код
        # ─────────────────────────────────────────────
        elif t == "code":
            if skip_text:
                continue

            code_text = b.get("text") or ""

            for ln in code_text.split("\n"):
                ln = ln.strip()

                if ln and not is_garbage_line(ln):
                    lines.append(normalize_text(ln))

        # ─────────────────────────────────────────────
        # Таблицы
        # ─────────────────────────────────────────────
        elif t == "table":
            if skip_text:
                continue

            rows = b.get("rows", [])

            if not isinstance(rows, list):
                continue

            for row in rows:
                if not isinstance(row, list):
                    continue

                cells = []

                for cell in row:
                    if cell is None:
                        cells.append("")
                    else:
                        cells.append(str(cell).strip())

                joined = " | ".join(cells)

                if joined and not is_garbage_line(joined):
                    lines.append(normalize_text(joined))

        # ─────────────────────────────────────────────
        # Изображения / OCR
        # ─────────────────────────────────────────────
        elif t in ("image", "image_ocr", "ocr"):
            if skip_images:
                continue

            ocr_text = (b.get("ocr_text") or "").strip()

            # Если распознавание слишком короткое — пропускаем
            if ocr_min > 0 and len(ocr_text) < ocr_min:
                continue

            if ocr_text:
                for ln in ocr_text.splitlines():
                    ln = ln.strip()

                    if ln and not is_garbage_line(ln):
                        lines.append(normalize_text(ln))

        # ─────────────────────────────────────────────
        # Разделитель
        # ─────────────────────────────────────────────
        elif t == "divider":
            continue

        # ─────────────────────────────────────────────
        # Универсальный fallback для неизвестных блоков
        # ─────────────────────────────────────────────
        else:
            # Если внутри есть вложенные блоки — обрабатываем их рекурсивно.
            # Это позволяет не терять изображения в нестандартных контейнерах.
            if isinstance(b.get("blocks"), list):
                lines += blocks_to_lines(
                    b.get("blocks", []),
                    ind + 1,
                    skip_text,
                    skip_images,
                    ocr_min
                )

            # Если это просто текстовый блок неизвестного типа
            elif not skip_text:
                v = (b.get("text") or "").strip()

                if v and not is_garbage_line(v):
                    lines.append(normalize_text(v))

    return lines


def threads_to_lines(threads, cache, topic_filters=None, abbreviations=None, telegraph=None,
                     ocr_min_global=0, thread_dates=None):
    lines = []
    first = True

    if thread_dates is None:
        try:
            thread_dates = load_thread_dates_cached(Project(FORUM_BASE_URL).dir)
        except Exception:
            thread_dates = {}

    for t in threads:
        k = norm_thread_url(t["url"])
        data = cache.get(k)

        if not data:
            get_logger().warning(f"threads_to_lines: нет в кэше {t['url']} (key={k})")
            continue

        title = data.get("title", "")
        skip_text = False
        skip_images = False

        if topic_filters:
            normalized_title = normalize_title_for_match(title)
            for f in topic_filters:
                kw = (f.get("keyword") or "").strip()
                normalized_kw = normalize_title_for_match(kw)
                if normalized_kw and normalized_kw in normalized_title:
                    skip_images = not f.get("parse_images", True)
                    skip_text = not f.get("parse_text", True)
                    break

        if not first:
            lines.append("=" * 50)

        first = False

        created_date = (data.get("created_date") or "").strip()
        if not created_date:
            created_date = (thread_dates.get(k, "") or "").strip()

        lines.append(f"тема создана {created_date}")
        lines.append("{")

        title_line = normalize_text(title)
        if abbreviations:
            title_line = apply_abbreviations(title_line, abbreviations)
        if telegraph:
            title_line = apply_telegraph(title_line, telegraph)

        lines.append(title_line)
        lines.append(data.get("url", ""))

        for post in data.get("posts", []):
            for ln in blocks_to_lines(
                post.get("blocks", []),
                skip_text=skip_text,
                skip_images=skip_images,
                ocr_min=ocr_min_global
            ):
                if abbreviations:
                    ln = apply_abbreviations(ln, abbreviations)
                if telegraph:
                    ln = apply_telegraph(ln, telegraph)
                lines.append(ln)

        lines.append("}")

    return lines


from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import lxml.etree as etree


def _remove_rsid_attrs(doc):
    """Удаляет rsid* атрибуты которые создают мусорные символы в некоторых просмотрщиках."""
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    for el in doc.element.body.iter():
        for attr in list(el.attrib.keys()):
            if 'rsid' in attr.lower():
                del el.attrib[attr]


def _collapse_paragraphs(doc):
    def para_text(p):
        return ''.join(r.text or '' for r in p.runs).strip()

    changed = True
    while changed:
        changed = False
        paras = list(doc.paragraphs)
        i = 0
        while i < len(paras) - 1:
            if para_text(paras[i]) == '' and para_text(paras[i + 1]) == '':
                p_el = paras[i + 1]._element
                p_el.getparent().remove(p_el)
                changed = True
                break
            i += 1


END_ANSWER_MARKER = "===КОНЕЦ ОТВЕТА==="
DOC_TAIL_TEXT = "Действую согласно промта в начале файла и в конце ответа всегда отвечай ===КОНЕЦ ОТВЕТА===."


class DocxGen:
    def __init__(self):
        self.doc = Document()
        sec = self.doc.sections[0]
        from docx.shared import Cm
        m = Cm(DOC_MARGIN)
        sec.left_margin = sec.right_margin = sec.top_margin = sec.bottom_margin = m
        st = self.doc.styles["Normal"]
        st.font.name = DOC_FONT
        st.font.size = Pt(DOC_SIZE)
        pf = st.paragraph_format
        pf.space_before = pf.space_after = Pt(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.line_spacing = DOC_LS

    def add(self, text=""):
        text = normalize_text(text)
        if not text:
            return
        p = self.doc.add_paragraph()
        pf = p.paragraph_format
        pf.space_before = pf.space_after = Pt(0)
        pf.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        pf.line_spacing_rule = WD_LINE_SPACING.SINGLE
        pf.line_spacing = DOC_LS
        r = p.add_run(text)

    def flatten(self):
        """v1.18: аналог Word-замены ^p -> " " (повторять, пока замен > 0).

        Весь документ схлопывается в один абзац: тексты всех абзацев идут
        подряд через одиночные пробелы. Первый абзац переиспользуется,
        остальные удаляются (sectPr не трогаем).
        """
        paras = list(self.doc.paragraphs)
        if not paras:
            return
        # ^p -> " ": склейка текстов через пробел (пустые пропускаем).
        joined = " ".join((p.text or "").strip() for p in paras if (p.text or "").strip())
        # Цикл "пока замен не будет 0": схлопываем кратные пробелы.
        while "  " in joined:
            joined = joined.replace("  ", " ")
        joined = joined.strip()
        # v1.19: вычищаем "=" и "_" — кроме маркера конца ответа.
        _guard = "\u0001END\u0001"
        joined = joined.replace(END_ANSWER_MARKER, _guard)
        joined = joined.replace("=", "").replace("_", "")
        joined = joined.replace(_guard, END_ANSWER_MARKER)
        # Удаление символов могло склеить пробелы — повторяем цикл до нуля замен.
        while "  " in joined:
            joined = joined.replace("  ", " ")
        joined = joined.strip()
        first = paras[0]
        body = self.doc.element.body
        for p in paras[1:]:
            body.remove(p._element)
        # paragraph.text = ... чистит раны; форматирование абзаца остаётся.
        first.text = joined

    def save(self, path):
        # Удаляем мусорные rsid
        _remove_rsid_attrs(self.doc)

        # Схлопываем абзацы
        _collapse_paragraphs(self.doc)

        # Сохраняем документ
        self.doc.save(path)

def build_server_docx(prompt_text: str, general_lines: list, server_title: str, server_lines: list):
    """
    Порядок слоёв в документе (v1.18, сверху вниз):
      СЛОЙ 1 — ПРОМТ (самый верх);
      СЛОЙ 2 — ОБЩИЕ ПРАВИЛА (правила проекта);
      СЛОЙ 3 — ЗАКОНОДАТЕЛЬНАЯ БАЗА конкретного сервера.
    Добавочный хвостовой шаблон про конец ответа удалён (ещё в v1.15):
    требование о нём уже есть в самом промте.
    Финал: весь документ схлопнут в один абзац — аналог Word-замены
    ^p -> " " (повторять, пока замен не будет 0).
    v1.19: в конце — хвост-приказ; "=" и "_" вычищены,
    кроме маркера ===КОНЕЦ ОТВЕТА===.
    """
    g = DocxGen()
    # ── СЛОЙ 1: промт — САМЫЙ ВЕРХ документа ──
    if prompt_text:
        for ln in prompt_text.split("\n"):
            g.add(ln.rstrip("\r"))
        g.add("=" * 60)
        g.add()
    # ── СЛОЙ 2: общие правила проекта ──
    g.add("ОБЩИЕ ПРАВИЛА")
    g.add()
    for ln in general_lines:
        g.add(ln)
    g.add()
    g.add("=" * 60)
    g.add()
    # ── СЛОЙ 3: законодательная база конкретного сервера ──
    g.add(f"ЗАКОНОДАТЕЛЬНАЯ БАЗА: {server_title}")
    g.add()
    for ln in server_lines:
        g.add(ln)
    # ── v1.19: хвост-приказ в самый конец ──
    g.add(DOC_TAIL_TEXT)
    # ── Финал v1.18: всё в один абзац ──
    g.flatten()
    return g


def build_common_docx(prompt_text: str, general_lines: list):
    """
    Порядок слоёв в документе (v1.18, сверху вниз):
      СЛОЙ 1 — ПРОМТ (самый верх);
      СЛОЙ 2 — ОБЩИЕ ПРАВИЛА.
    Добавочный хвостовой шаблон про конец ответа удалён (ещё в v1.15):
    требование о нём уже есть в самом промте.
    Финал: весь документ схлопнут в один абзац — аналог Word-замены
    ^p -> " " (повторять, пока замен не будет 0).
    v1.19: в конце — хвост-приказ; "=" и "_" вычищены,
    кроме маркера ===КОНЕЦ ОТВЕТА===.
    """
    g = DocxGen()
    # ── СЛОЙ 1: промт — САМЫЙ ВЕРХ документа ──
    if prompt_text:
        for ln in prompt_text.split("\n"):
            g.add(ln.rstrip("\r"))
        g.add("=" * 60)
        g.add()
    # ── СЛОЙ 2: общие правила проекта ──
    g.add("ОБЩИЕ ПРАВИЛА")
    g.add()
    for ln in general_lines:
        g.add(ln)
    # ── v1.19: хвост-приказ в самый конец ──
    g.add(DOC_TAIL_TEXT)
    # ── Финал v1.18: всё в один абзац ──
    g.flatten()
    return g


class Project:
    def __init__(self, root_url: str):
        host = urlparse(root_url).netloc or "forum"
        self.dir = SERVERS_ROOT / safe_name(host, 60)
        self.dir.mkdir(parents=True, exist_ok=True)
        self.struct = self.dir / STRUCTURE_FILE
        self.cfg = self.dir / CONFIG_FILE
        self.cookies = self.dir / COOKIES_FILE
        self.checks = self.dir / CHECKS_FILE
        self.out = self.dir / OUTPUT_SUBDIR
        self.cache_dir = self.dir / CACHE_SUBDIR
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def save_struct(self, nodes) -> bool:
        if not nodes:
            get_logger().warning(
                "save_struct: попытка сохранить ПУСТУЮ структуру отклонена, "
                "существующий structure.json не тронут"
            )
            return False
        self.struct.write_text(
            json.dumps([n.to_dict() for n in nodes], ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        return True

    def load_struct(self):
        if not self.struct.exists():
            return None
        try:
            return [Node.from_dict(d) for d in json.loads(self.struct.read_text(encoding="utf-8"))]
        except Exception:
            return None

    def save_config(self, data: dict):
        cur = self.load_config()
        cur.update(data)
        self.cfg.write_text(json.dumps(cur, ensure_ascii=False, indent=2), encoding="utf-8")

    def load_config(self) -> dict:
        if not self.cfg.exists():
            return {}

        try:
            return json.loads(self.cfg.read_text(encoding="utf-8-sig"))
        except Exception as e:
            get_logger().error(f"Не удалось прочитать конфиг {self.cfg}: {e}")
            return {}

    def save_cookies(self, cookies: list):
        self.cookies.write_text(
            json.dumps(cookies, ensure_ascii=False, indent=2, default=str),
            encoding="utf-8",
        )

    def load_cookies(self) -> list:
        if not self.cookies.exists():
            return []
        try:
            return json.loads(self.cookies.read_text(encoding="utf-8"))
        except Exception:
            return []

    def save_checks(self, data: dict):
        """
        Сохраняет состояние галочек и пометок.
        data = {
            'checks': {url: bool, ...},
            'marks':  {url: mark_value, ...},
            'thread_checks': {thread_url: bool, ...}
        }
        """
        self.checks.write_text(
            json.dumps(data, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def load_checks(self) -> dict:
        if not self.checks.exists():
            return {}
        try:
            return json.loads(self.checks.read_text(encoding="utf-8"))
        except Exception:
            return {}

    def load_thread_cache(self, thread_url: str):
        return load_thread_cache(self.dir, thread_url)

    def save_thread_cache(self, thread_url: str, data: dict, checksum: str, fpo: bool = None):
        save_thread_cache(self.dir, thread_url, data, checksum, fpo)


def compute_html_checksum(html: str) -> str:
    """
    SHA-256 только от стабильного содержимого постов темы.
    Игнорирует динамические элементы: счётчики, токены, время посещения и т.д.
    """
    if not html:
        return ""
    try:
        soup = BeautifulSoup(html, BS_PARSER)

        # Удаляем все динамические блоки перед снятием контрольной суммы
        for sel in (
            "script", "style", "noscript",
            # XenForo динамика
            "[data-xf-init]", ".js-visitorMenuBody",
            ".p-navgroup--member", ".p-navgroup--guest",
            # Счётчики просмотров/ответов — меняются при каждом заходе
            ".structItem-cell--meta", ".structItem-cell--latest",
            ".message-lastPost", ".node-stats",
            # CSRF и скрытые поля
            "input[name='_xfToken']", "input[name='_xfRequestUri']",
            # Реклама и баннеры
            ".p-footer", ".p-header", "header", "footer",
            # Навигация
            "nav", ".p-nav", ".p-breadcrumbs", ".pageNav",
            # Онлайн-индикаторы и аватары (содержат временные ссылки)
            ".message-avatar", ".avatarWrapper",
            # Временные метки "онлайн X минут назад" и т.п.
            ".u-concealed", "[data-last-click]",
        ):
            for tag in soup.select(sel):
                tag.decompose()

        # Собираем только содержимое постов
        post_parts = []
        for article in soup.select("article.message, .js-post, .message--post"):
            # Автор
            author = ""
            for s in (".message-name .username", ".username", "[itemprop='name']"):
                t = article.select_one(s)
                if t:
                    author = t.get_text(strip=True)
                    break

            # Тело поста
            body = (
                article.select_one(".message-body .bbWrapper")
                or article.select_one(".bbWrapper")
                or article.select_one(".message-body")
            )
            body_text = body.get_text(separator=" ", strip=True) if body else ""

            if author or body_text:
                post_parts.append(f"{author}::{body_text}")

        if not post_parts:
            # Fallback: если посты не нашлись — берём заголовок темы
            title_tag = soup.select_one("h1.p-title-value, h1.p-title, h1")
            fallback = title_tag.get_text(strip=True) if title_tag else html[:500]
            normalized = re.sub(r'\s+', ' ', fallback).strip()
            return hashlib.sha256(normalized.encode('utf-8')).hexdigest()

        combined = "\n".join(post_parts)
        normalized = re.sub(r'\s+', ' ', combined).strip()
        return hashlib.sha256(normalized.encode('utf-8')).hexdigest()

    except Exception:
        # При любой ошибке парсинга — используем старый метод как fallback
        normalized = re.sub(r'\s+', ' ', html or '').strip()
        return hashlib.sha256(normalized.encode('utf-8')).hexdigest()


def cache_path_for(project_dir: Path, thread_url: str) -> Path:
    """Путь к файлу кэша для конкретной темы."""
    key = hashlib.md5(thread_url.encode()).hexdigest()
    cache_dir = project_dir / CACHE_SUBDIR
    cache_dir.mkdir(parents=True, exist_ok=True)
    return cache_dir / f"{key}.json"


def load_thread_cache(project_dir: Path, thread_url: str):
    """Загружает кэш темы. Возвращает dict или None."""
    p = cache_path_for(project_dir, thread_url)
    if not p.exists():
        return None
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except Exception:
        return None


def save_thread_cache(project_dir: Path, thread_url: str, data: dict, checksum: str, fpo: bool = None):
    p = cache_path_for(project_dir, thread_url)
    payload = {'checksum': checksum, 'data': data, 'fpo': fpo}
    p.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding='utf-8')


def get_thread_checksum_from_web(http: HttpClient, thread_url: str) -> tuple:
    """
    Загружает HTML темы и вычисляет контрольную сумму.
    Возвращает (html, checksum) или (None, None).
    """
    html = http.get(thread_url)
    if not html:
        return None, None
    return html, compute_html_checksum(html)

_EASYOCR_READER = None


def _detect_image_ext(url: str, data: bytes) -> str:
    if data.startswith(b"\xff\xd8\xff"):
        return ".jpg"
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return ".png"
    if data.startswith(b"GIF87a") or data.startswith(b"GIF89a"):
        return ".gif"
    if data[:4] == b"RIFF" and data[8:12] == b"WEBP":
        return ".webp"
    if data.startswith(b"BM"):
        return ".bmp"
    ext = os.path.splitext(urlparse(url).path)[1].lower()
    if ext in (".jpg", ".jpeg", ".png", ".gif", ".webp", ".bmp"):
        return ext
    return ".img"


def _valid_image_bytes(data) -> bool:
    if not data or len(data) < 50 or len(data) > OCR_IMAGE_MAX_BYTES:
        return False
    low = data[:1000].lstrip().lower()
    if low.startswith(b"<html") or low.startswith(b"<!doctype"):
        return False
    if low.startswith(b"{") or low.startswith(b"["):
        return False
    return True

def _download_image_bytes(url: str, cookies=None, referer: str = FORUM_BASE_URL,
                          timeout: int = 30, ua_variant: int = 0) -> bytes:
    if not url:
        return None
    try:
        if url.lower().startswith("data:"):
            try:
                b64 = url.split(",", 1)[1]
                return base64.b64decode(b64)
            except Exception:
                return None
        if url.startswith("//"):
            url = "https:" + url
        cookie_dict = {}
        try:
            if urlparse(url).netloc.lower().endswith(FORUM_HOST):
                for c in cookies or []:
                    name = c.get("name")
                    if name:
                        cookie_dict[name] = c.get("value", "")
        except Exception:
            cookie_dict = {}
        uas = [
            UA,
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64; rv:132.0) Gecko/20100101 Firefox/132.0",
            "Mozilla/5.0 (iPhone; CPU iPhone OS 17_0 like Mac OS X) AppleWebKit/605.1.15 "
            "(KHTML, like Gecko) Version/17.0 Mobile/15E148 Safari/604.1",
        ]
        headers = {
            "User-Agent": uas[ua_variant % len(uas)],
            "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Referer": referer or FORUM_BASE_URL,
            "Sec-Fetch-Dest": "image",
            "Sec-Fetch-Mode": "no-cors",
            "Sec-Fetch-Site": "cross-site",
        }
        r = requests.get(url, headers=headers, cookies=cookie_dict, timeout=timeout)
        if not r.ok:
            return None
        if not _valid_image_bytes(r.content):
            return None
        return r.content
    except Exception:
        return None

def _download_image_curl(url: str, timeout: int = 15) -> bytes:
    if not url:
        return None
    exe = shutil.which("curl") or ("curl.exe" if sys.platform == "win32" else None)
    if not exe:
        return None
    try:
        p = subprocess.run(
            [exe, "-s", "-S", "-L", "--max-time", str(timeout),
             "-A", UA, "-e", FORUM_BASE_URL, url],
            capture_output=True, timeout=timeout + 5,
        )
        if p.returncode == 0 and _valid_image_bytes(p.stdout):
            return p.stdout
    except Exception:
        pass
    return None

def _download_image_via_proxy(proxy_url: str, cookies=None, timeout: int = 15) -> bytes:
    if not proxy_url:
        return None
    try:
        u = proxy_url
        if u.startswith("//"):
            u = "https:" + u
        if u.startswith("/"):
            u = urljoin(FORUM_BASE_URL, u)
        if not u.startswith("http"):
            return None
        cookie_dict = {}
        for c in cookies or []:
            name = c.get("name")
            if name:
                cookie_dict[name] = c.get("value", "")
        headers = {
            "User-Agent": UA,
            "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
            "Accept-Language": "ru-RU,ru;q=0.9,en-US;q=0.8,en;q=0.7",
            "Referer": FORUM_BASE_URL,
            "Sec-Fetch-Dest": "image",
            "Sec-Fetch-Mode": "no-cors",
            "Sec-Fetch-Site": "same-origin",
        }
        r = requests.get(u, headers=headers, cookies=cookie_dict, timeout=timeout)
        if r.ok and _valid_image_bytes(r.content):
            return r.content
    except Exception:
        pass
    return None

def _download_image_browser_tab(driver, url: str, timeout: int = 25) -> bytes:
    if not url or not url.startswith("http"):
        return None
    try:
        original = driver.current_window_handle
    except Exception:
        return None
    before = set(driver.window_handles)
    try:
        driver.set_script_timeout(timeout)
        driver.execute_script("window.open(arguments[0], '_blank');", url)
    except Exception:
        return None
    new_handle = None
    t0 = time.monotonic()
    while time.monotonic() - t0 < 2:
        diff = set(driver.window_handles) - before
        if diff:
            new_handle = diff.pop()
            break
        time.sleep(0.2)
    if not new_handle:
        return None
    try:
        driver.switch_to.window(new_handle)
        t1 = time.monotonic()
        while time.monotonic() - t1 < 4:
            try:
                if driver.execute_script("return document.readyState;") == "complete":
                    break
            except Exception:
                break
            time.sleep(0.3)
        script = (
            "const cb = arguments[arguments.length - 1];"
            "try {"
            "  fetch(window.location.href, {credentials: 'include'})"
            "    .then(r => r.ok ? r.arrayBuffer() : null)"
            "    .then(buf => {"
            "      if (!buf || buf.byteLength < 50 || buf.byteLength > 20971520) { cb(null); return; }"
            "      const bytes = new Uint8Array(buf);"
            "      let binary = '';"
            "      const chunk = 0x8000;"
            "      for (let i = 0; i < bytes.length; i += chunk) {"
            "        binary += String.fromCharCode.apply(null, bytes.subarray(i, i + chunk));"
            "      }"
            "      cb(btoa(binary));"
            "    })"
            "    .catch(() => cb(null));"
            "} catch (e) { cb(null); }"
        )
        b64 = driver.execute_async_script(script)
        if not b64:
            return None
        raw = base64.b64decode(b64)
        if _valid_image_bytes(raw):
            return raw
        return None
    except Exception:
        return None
    finally:
        try:
            if new_handle in set(driver.window_handles):
                driver.close()
        except Exception:
            pass
        try:
            driver.switch_to.window(original)
        except Exception:
            pass


_last_ocr_engine = ""

def _ocr_image_bytes(data: bytes) -> str:
    global _last_ocr_engine
    if not data or not HAS_PIL:
        _last_ocr_engine = ""
        return ""

    try:
        from PIL import Image
        import io
        img = Image.open(io.BytesIO(data)).convert("RGB")
    except Exception:
        return ""

    # Приоритет 1: EasyOCR
    try:
        global _EASYOCR_READER
        import easyocr
        import numpy as np
        if _EASYOCR_READER is None:
            _EASYOCR_READER = easyocr.Reader(["ru", "en"], verbose=False, gpu=True)
        result = _EASYOCR_READER.readtext(np.array(img), detail=0, paragraph=True)
        lines = [str(x).strip() for x in (result or []) if str(x).strip()]
        if lines:
            _last_ocr_engine = "EasyOCR"
            return "\n".join(lines).strip()
    except Exception:
        pass

    # Приоритет 2: pytesseract
    try:
        import pytesseract
        try:
            txt = pytesseract.image_to_string(img, lang="rus+eng").strip()
        except Exception:
            txt = pytesseract.image_to_string(img).strip()
        if txt:
            _last_ocr_engine = "pytesseract"
            return txt
    except Exception:
        pass
    _last_ocr_engine = ""
    return ""

def _ocr_worth_running(data: bytes) -> bool:
    """Стоит ли гнать движок OCR: отсекает иконки/пустышки (v1.17)."""
    if not data or len(data) < OCR_MIN_BYTES:
        return False
    try:
        from PIL import Image
        import io as _io
        with Image.open(_io.BytesIO(data)) as img:
            w, h = img.size
        if min(w, h) < OCR_MIN_DIM:
            return False
    except Exception:
        pass
    return True


def _ocr_content_key(data: bytes) -> str:
    try:
        return hashlib.sha256(data or b"").hexdigest()[:32]
    except Exception:
        return ""


def _ocr_lookup_by_content(data: bytes, images_dir) -> str | None:
    """Уже распознанный такой же файл (другой URL, то же содержимое)?"""
    if not data:
        return None
    try:
        p = Path(images_dir) / f"c_{_ocr_content_key(data)}.txt"
        if p.exists():
            return p.read_text(encoding="utf-8")
    except Exception:
        pass
    return None


def _ocr_store_by_content(data: bytes, images_dir, text: str):
    try:
        p = Path(images_dir) / f"c_{_ocr_content_key(data)}.txt"
        if not p.exists():
            p.write_text(text or "", encoding="utf-8")
    except Exception:
        pass


def _process_image_block(block: dict, ctx: dict) -> list:
    src = (block.get("src") or "").strip()
    alt = (block.get("alt") or "").strip()
    proxy_url = (block.get("original_src") or "").strip()
    if not src:
        return []
    if ctx.get("count", 0) >= OCR_MAX_IMAGES_PER_POST:
        return []
    ctx["count"] = ctx.get("count", 0) + 1
    url = src.split("#")[0]
    if url.startswith("//"):
        url = "https:" + url
    em = ctx.get("emitter")
    pid = ctx.get("parent_id")
    seq_box = ctx.get("seq_box")
    if isinstance(seq_box, list) and seq_box:
        seq_box[0] += 1
        disp_no = seq_box[0]
    else:
        disp_no = ctx.get("count", 0)
    seen = ctx.get("seen")
    if isinstance(seen, dict) and url in seen:
        if em and pid:
            s_dup = em.step_added(pid, f"🖼 Изображение #{disp_no}", "🖼")
            em.step_state(s_dup, "done", "дубль URL — из памяти")
        return [dict(b) for b in seen[url]]
    s_img = None
    if em and pid:
        s_img = em.step_added(pid, f"🖼 Изображение #{disp_no}", "🖼")
        em.step_state(s_img, "running")
    images_dir = Path(ctx.get("images_dir") or (Path.cwd() / IMAGES_SUBDIR))
    images_dir.mkdir(parents=True, exist_ok=True)
    url_hash = hashlib.sha256(url.encode("utf-8")).hexdigest()
    ocr_path = images_dir / f"{url_hash}.txt"
    meta_path = images_dir / f"{url_hash}.url.txt"
    img_candidates = [
        p for p in images_dir.glob(f"{url_hash}.*")
        if p.suffix.lower() != ".txt"
    ]
    img_path = img_candidates[0] if img_candidates else None
    raw = None
    method = ""
    t_start = time.monotonic()
    if img_path is None:
        s_dl = None
        if em and s_img:
            s_dl = em.step_added(s_img, "Скачивание", "⬇")
            em.step_state(s_dl, "running")
        fetcher = ctx.get("fetcher")
        if callable(fetcher):
            try:
                res = fetcher(url, proxy_url=proxy_url)
            except Exception:
                res = None
            if isinstance(res, tuple):
                raw = res[0]
                method = res[1] or ""
            else:
                raw = res
        if not raw:
            raw = _download_image_bytes(
                url,
                cookies=ctx.get("cookies"),
                referer=ctx.get("referer") or FORUM_BASE_URL,
                timeout=IMAGE_FALLBACK_TIMEOUT,
            )
            if raw:
                method = "requests_fallback"
        elapsed = time.monotonic() - t_start
        if not raw:
            if em and s_dl:
                em.step_state(s_dl, "error", f"не скачалось ({elapsed:.1f}с)")
            if em and s_img:
                em.step_state(s_img, "error", f"не скачалось ({elapsed:.1f}с)")
            if isinstance(ctx.get("seen"), dict):
                ctx["seen"][url] = []
            return []
        if em and s_dl:
            size_kb = len(raw) // 1024
            em.step_state(s_dl, "done", f"{size_kb} КБ, {method or 'direct'}, {elapsed:.1f}с")
        ext = _detect_image_ext(url, raw)
        img_path = images_dir / f"{url_hash}{ext}"
        try:
            img_path.write_bytes(raw)
        except Exception:
            pass
        try:
            meta_path.write_text(url, encoding="utf-8")
        except Exception:
            pass
    else:
        if em and s_img:
            em.step_metric(s_img, "detail", "из кэша")
    if ocr_path.exists():
        try:
            ocr_text = ocr_path.read_text(encoding="utf-8")
        except Exception:
            ocr_text = ""
        if em and s_img:
            em.step_metric(s_img, "detail", "OCR из кэша")
    else:
        if raw is None:
            try:
                raw = img_path.read_bytes()
            except Exception:
                raw = b""
        content_hit = _ocr_lookup_by_content(raw, images_dir) if raw else None
        if content_hit is not None:
            ocr_text = content_hit
            if em and s_img:
                em.step_metric(s_img, "detail", "OCR по содержимому (дубль)")
        elif not raw or not _ocr_worth_running(raw):
            ocr_text = ""
            if em and s_img:
                em.step_metric(s_img, "detail", "пропуск OCR: маленькая")
        else:
            s_ocr = None
            if em and s_img:
                s_ocr = em.step_added(s_img, "OCR распознавание", "🔤")
                em.step_state(s_ocr, "running")
            t_ocr = time.monotonic()
            ocr_text = _ocr_image_bytes(raw) if raw else ""
            elapsed_ocr = time.monotonic() - t_ocr
            if em and s_ocr:
                em.step_state(s_ocr, "done", f"{elapsed_ocr:.1f}с, {len(ocr_text)} симв.")
            _ocr_store_by_content(raw, images_dir, ocr_text)
        try:
            ocr_path.write_text(ocr_text or "", encoding="utf-8")
        except Exception:
            pass
    if em and s_img:
        em.step_state(s_img, "done", f"OCR: {len(ocr_text)} симв.")
    _res_blocks = [
        {
            "type": "image",
            "src": url,
            "alt": alt,
            "ocr_text": ocr_text or "",
        }
    ]
    if isinstance(ctx.get("seen"), dict):
        ctx["seen"][url] = [dict(b) for b in _res_blocks]
    return _res_blocks

def _process_blocks_with_images(blocks: list, ctx: dict) -> list:
    """Рекурсивно проходит по блокам и заменяет type='image' на текст + OCR."""
    out = []
    for b in blocks or []:
        if not isinstance(b, dict):
            continue
        t = b.get("type")
        if t == "image":
            out.extend(_process_image_block(b, ctx))
        elif t in ("quote", "spoiler") and isinstance(b.get("blocks"), list):
            nb = dict(b)
            nb["blocks"] = _process_blocks_with_images(b["blocks"], ctx)
            out.append(nb)
        else:
            out.append(b)
    return out

def enrich_thread_data_with_images(data: dict, project_dir: Path,
                                   cookies=None, base_url: str = "",
                                   fetcher=None, emitter=None, parent_id=None) -> dict:
    if not data or not isinstance(data.get("posts"), list):
        return data
    images_dir = Path(project_dir) / IMAGES_SUBDIR
    base = base_url or data.get("url") or FORUM_BASE_URL
    # v1.17: сквозная нумерация на тему + дедуп одинаковых URL между постами.
    seq_box = [0]
    seen = {}
    for post in data["posts"]:
        ctx = {
            "images_dir": images_dir,
            "cookies": cookies or [],
            "base": base,
            "referer": base_url or FORUM_BASE_URL,
            "fetcher": fetcher,
            "count": 0,
            "seq_box": seq_box,
            "seen": seen,
            "emitter": emitter,
            "parent_id": parent_id,
        }
        post["blocks"] = _process_blocks_with_images(post.get("blocks", []), ctx)
    return data

C = dict(
    bg="#12121f", panel="#1c1c2e", panel2="#22223a",
    accent="#7c6af7", accent2="#4ea8de", green="#3ddc84",
    red="#ff5370", yellow="#ffcb6b", orange="#f78c6c",
    text="#eeeeff", sub="#8888aa", entry="#0d0d1a",
    border="#33335a", sel="#2d2b55",
    lv0="#c792ea", lv1="#82aaff", lv2="#4ec9b0",
    lv3="#dcdcaa", lv4="#ce9178", lv5="#9cdcfe",
    debug="#666688", search="#ffee58",
)
LEVEL_COLORS = [C["lv0"], C["lv1"], C["lv2"], C["lv3"], C["lv4"], C["lv5"]]
ICONS = ["🌐", "📁", "🗂", "📂", "📋", "📄"]
MARK_LABELS = {None: "", "general": "📜 Общие", "server": "🖥 Сервер"}

GARBAGE_PATTERNS = [
    r'\[ИЗОБРАЖЕНИЕ НЕ СКАЧАЛОСЬ:.*?\]',
    r'\[ИЗОБРАЖЕНИЕ:.*?\]',
    r'\[ALT изображения:.*?\]',
    r'\[РАСПОЗНАННЫЙ ТЕКСТ С ИЗОБРАЖЕНИЯ\]',
    r'\(текст на изображении не распознан\)',
    r'\[Спойлер:.*?\]',
    r'\[Спойлер\]',
]

def normalize_text(s):
    s = re.sub(r'\s+', ' ', s)
    return s.strip()

def is_garbage_line(s):
    s = s.strip()
    if not s:
        return True
    for p in GARBAGE_PATTERNS:
        if re.match(p, s, re.IGNORECASE):
            return True
    return False

def apply_abbreviations(s, abbreviations):
    if not s or not abbreviations:
        return s

    items = sorted(
        abbreviations.items(),
        key=lambda x: len((x[0] or "").strip()),
        reverse=True
    )

    for full, abbr in items:
        full = (full or "").strip()
        abbr = (abbr or "").strip()

        if full:
            s = s.replace(full, abbr)

    return s


def apply_telegraph(s, replacements):
    if not s or not replacements:
        return s

    items = sorted(
        replacements.items(),
        key=lambda x: len((x[0] or "").strip()),
        reverse=True
    )

    for full, short in items:
        full = (full or "").strip()
        short = (short or "").strip()

        if full:
            s = s.replace(full, short)

    return s


def normalize_title_for_match(s: str) -> str:
    """
    Нормализует название для сравнения:
    - нижний регистр;
    - ё -> е;
    - убирает лишние пробелы;
    - убирает пунктуацию;
    - оставляет только буквы/цифры/пробелы.
    """
    s = (s or "").lower().replace("ё", "е")
    s = s.replace("_", " ")
    s = re.sub(r"[\u00a0\t]+", " ", s)
    s = re.sub(r"[^\w\s]+", " ", s)
    s = re.sub(r"\s+", " ", s).strip()
    return s


_autocheck_node_keywords = []
_autocheck_thread_keywords = []
_autocheck_thread_exclude = []


def set_autocheck_keywords(node_kw, thread_kw, thread_exclude=None):
    global _autocheck_node_keywords, _autocheck_thread_keywords, _autocheck_thread_exclude
    _autocheck_node_keywords = list(node_kw or [])
    _autocheck_thread_keywords = list(thread_kw or [])
    if thread_exclude is not None:
        _autocheck_thread_exclude = list(thread_exclude)


def is_excluded_by_words(title: str, section: str = "", exclude_words=None) -> bool:
    """True, если слово-исключение входит в название темы или раздела (v1.19).

    Используется и для автоотметки (исключения ключевых слов), и для
    скрытия тем в «Истории». Пустой список = ничего не скрываем.
    """
    if not exclude_words:
        return False
    texts = [normalize_title_for_match(title), normalize_title_for_match(section or "")]
    texts = [x for x in texts if x]
    if not texts:
        return False
    for w in exclude_words:
        if isinstance(w, dict):
            w = w.get("word", "")
        needle = normalize_title_for_match(str(w or ""))
        if not needle:
            continue
        for x in texts:
            if needle in x:
                return True
    return False

def struct_row_hidden(title, section="", q="", excl=None):
    """Строка темы скрыта в «Структуре» (v1.25, чистая функция для тестов).

    q уже в нижнем регистре. Исключение ищется в теме И разделе —
    как в автоотметке («есть в теме/разделе»).
    """
    t = (title or "").lower()
    if q and q not in t:
        return True
    if excl and is_excluded_by_words(title, section, excl):
        return True
    return False

def match_keyword(text: str, entry) -> bool:
    """Совпадение ключевого слова: целое слово + позиция (v1.20).

    entry: str — как раньше (подстрока, anywhere); dict —
      {"keyword": ..., "whole": bool, "pos": "any|start|end"}.
    whole=True: «закон» не совпадает с «законодательство».
    pos="start"/"end": совпадение только в начале/в конце названия.
    """
    if isinstance(entry, dict):
        kw = entry.get("keyword", "")
        whole = bool(entry.get("whole", False))
        pos = (entry.get("pos", "any") or "any").strip().lower()
    else:
        kw = entry
        whole = False
        pos = "any"
    needle = normalize_title_for_match(str(kw or ""))
    if not needle:
        return False
    ntext = normalize_title_for_match(text)
    if not ntext:
        return False
    if pos not in ("any", "start", "end"):
        pos = "any"
    if whole:
        if pos == "start":
            return ntext == needle or ntext.startswith(needle + " ")
        if pos == "end":
            return ntext == needle or ntext.endswith(" " + needle)
        return re.search(r"\b%s\b" % re.escape(needle), ntext) is not None
    if pos == "start":
        return ntext.startswith(needle)
    if pos == "end":
        return ntext.endswith(needle)
    return needle in ntext


def should_autocheck_new_node(title: str) -> bool:
    for item in _autocheck_node_keywords:
        if match_keyword(title, item):
            return True
    return False

def should_autocheck_new_thread(title: str) -> bool:
    for item in _autocheck_thread_keywords:
        if match_keyword(title, item):
            return True
    return False

def final_thread_check_value(*, excluded: bool, kw_eval, was_checked: bool, kw_match: bool) -> bool:
    """Финальное решение по галочке темы (v1.24, чистая функция для тестов).

    Исключение бьёт всё: и восстановление старой отметки, и совпадение
    ключевого слова. Порядок остальных веток — как было исторически.
    """
    if excluded:
        return False
    if kw_eval == 'uncheck':
        return False
    if was_checked:
        return True
    if kw_match:
        return True
    return False

_TIME_DT_RE = re.compile(
    r'<time\b[^>]*\bclass\s*=\s*["\'][^"\']*\bu-dt\b[^"\']*["\'][^>]*>', re.IGNORECASE)
_DATA_DATE_RE = re.compile(r'data-date\s*=\s*["\'](\d{2}\.\d{2}\.\d{4})["\']')
_DT_ATTR_RE = re.compile(r'datetime\s*=\s*["\']([^"\']+)["\']')


def extract_thread_creation_date(html):
    # v1.16: быстрый путь — regex по первому <time class="u-dt"> без полного парсинга.
    if html:
        try:
            m = _TIME_DT_RE.search(html[:200000])
            if m:
                tag = m.group(0)
                dm = _DATA_DATE_RE.search(tag)
                if dm:
                    datetime.strptime(dm.group(1), '%d.%m.%Y')
                    return dm.group(1)
                dtm = _DT_ATTR_RE.search(tag)
                if dtm:
                    dt = datetime.fromisoformat(dtm.group(1)[:19])
                    return dt.strftime('%d.%m.%Y')
        except Exception:
            pass
    try:
        soup = BeautifulSoup(html, BS_PARSER)
        time_el = soup.find('time', class_='u-dt')
        if not time_el:
            return ''
        data_date = time_el.get('data-date', '')
        if data_date:
            datetime.strptime(data_date, '%d.%m.%Y')
            return data_date
        dt_attr = time_el.get('datetime', '')
        if dt_attr:
            dt = datetime.fromisoformat(dt_attr[:19])
            return dt.strftime('%d.%m.%Y')
    except Exception:
        pass
    return ''


def thread_dates_path(project_dir: Path) -> Path:
    return project_dir / "thread_dates.json"


def load_thread_dates(project_dir: Path) -> dict:
    p = thread_dates_path(project_dir)
    if not p.exists():
        return {}
    try:
        return json.loads(p.read_text(encoding='utf-8'))
    except Exception:
        return {}


_thread_dates_lock = threading.Lock()
_thread_dates_cache = {"key": None, "mtime": 0.0, "data": {}}


def save_thread_dates(project_dir: Path, dates: dict):
    p = thread_dates_path(project_dir)
    existing = load_thread_dates(project_dir)
    existing.update(dates)
    p.write_text(json.dumps(existing, ensure_ascii=False, indent=2), encoding='utf-8')
    with _thread_dates_lock:
        try:
            _thread_dates_cache["key"] = str(p.resolve())
            _thread_dates_cache["mtime"] = p.stat().st_mtime
            _thread_dates_cache["data"] = dict(existing)
        except Exception:
            _thread_dates_cache["key"] = None


def load_thread_dates_cached(project_dir: Path) -> dict:
    """load_thread_dates + кэш в памяти (защита по mtime). Семантика та же, только быстрее."""
    try:
        p = thread_dates_path(project_dir)
        key = str(p.resolve())
    except Exception:
        return load_thread_dates(project_dir)
    with _thread_dates_lock:
        if _thread_dates_cache.get("key") == key:
            try:
                if p.exists() and p.stat().st_mtime == _thread_dates_cache.get("mtime"):
                    return _thread_dates_cache["data"]
            except Exception:
                pass
    data = load_thread_dates(project_dir)
    with _thread_dates_lock:
        try:
            _thread_dates_cache["key"] = key
            _thread_dates_cache["mtime"] = p.stat().st_mtime if p.exists() else 0.0
            _thread_dates_cache["data"] = data
        except Exception:
            pass
    return data


def evaluate_thread_keyword_match(thr_title, thread_date, keywords_config):
    if not keywords_config:
        return None
    if not normalize_title_for_match(thr_title):
        return None
    has_match = False
    should_check = False
    for entry in keywords_config:
        if isinstance(entry, dict):
            date_filter = entry.get('date', '').strip()
        else:
            date_filter = ''
        if not match_keyword(thr_title, entry):
            continue
        has_match = True
        if not date_filter or not thread_date:
            should_check = True
            break
        try:
            filter_dt = datetime.strptime(date_filter, '%d.%m.%Y')
            thread_dt = datetime.strptime(thread_date, '%d.%m.%Y')
            if thread_dt >= filter_dt:
                should_check = True
                break
        except ValueError:
            should_check = True
            break
    if not has_match:
        return None
    return 'check' if should_check else 'uncheck'

def history_layers(scopes, server):
    """Слои темы в «Истории» (v1.23): подмножество {"general", "server"}.

    Без привязки — общие. Темы чужих серверов — пусто (скрыты).
    """
    sc = [s for s in (scopes or []) if s]
    if not sc:
        return {"general"}
    out = set()
    if "general" in sc:
        out.add("general")
    if server and f"server:{server}" in sc:
        out.add("server")
    return out


class HistoryExcludeWindow(tk.Toplevel):
    """Слова-исключения для «Истории» (v1.19): тема со словом в названии скрывается."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("🚫 Слова-исключения «Истории»")
        self.configure(bg=C["panel"])
        self.geometry("520x500")
        self.transient(parent)
        self.grab_set()
        self._entries = []

        tk.Label(self, text="Темы с этими словами в названии\nне показываются в «Истории»:",
                 bg=C["panel"], fg=C["sub"], font=("Segoe UI", 9),
                 justify="left").pack(anchor="w", padx=10, pady=(8, 2))

        scroll = tk.Frame(self, bg=C["panel"])
        scroll.pack(fill="both", expand=True, padx=10, pady=5)
        canvas = tk.Canvas(scroll, bg=C["panel"], highlightthickness=0)
        vsb = ttk.Scrollbar(scroll, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        frame = tk.Frame(canvas, bg=C["panel"])
        self._frame = frame
        win = canvas.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))

        try:
            words = list(self.app.project.load_config().get("history_exclude_words", []) or [])
        except Exception:
            words = []
        for w in words:
            self._add_entry(str(w))
        if not words:
            self._add_entry("")

        btns = tk.Frame(self, bg=C["panel"])
        btns.pack(fill="x", padx=10, pady=5)
        self.app._mk_btn(btns, "+ Слово", lambda: self._add_entry(""), C["green"], small=True).pack(side="left", padx=2)

        bottom = tk.Frame(self, bg=C["panel"])
        bottom.pack(fill="x", padx=10, pady=(0, 8))
        self.app._mk_btn(bottom, "Сохранить", self._save, C["green"]).pack(side="left", padx=2)
        self.app._mk_btn(bottom, "Отмена", self.destroy, C["red"]).pack(side="left", padx=2)

    def _add_entry(self, value=""):
        row = tk.Frame(self._frame, bg=C["panel"])
        row.pack(fill="x", pady=1)
        e = tk.Entry(row, bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                     relief="flat", font=("Segoe UI", 9),
                     highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        e.pack(side="left", fill="x", expand=True, padx=(2, 4), ipady=2)
        e.insert(0, value)
        self.app._mk_btn(row, "✕", lambda r=row: self._remove(r), C["red"], small=True).pack(side="right", padx=2)
        self._entries.append((row, e))

    def _remove(self, row):
        self._entries = [(r, e) for r, e in self._entries if r is not row]
        try:
            row.destroy()
        except Exception:
            pass

    def _save(self):
        words = [e.get().strip() for _, e in self._entries if e.get().strip()]
        try:
            self.app.project.save_config({"history_exclude_words": words})
        except Exception:
            pass
        self.app._log(f"🚫 Слов-исключений «Истории»: {len(words)}", "ok")
        self.destroy()


class StructureExcludeWindow(tk.Toplevel):
    """Слова-исключения для «Структуры» (v1.25): тема скрывается из дерева."""

    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("🚫 Слова-исключения «Структуры»")
        self.configure(bg=C["panel"])
        self.geometry("520x500")
        self.transient(parent)
        self.grab_set()
        self._entries = []

        tk.Label(self, text="Темы с этими словами в названии\nне показываются в «Структуре форума»:",
                 bg=C["panel"], fg=C["sub"], font=("Segoe UI", 9),
                 justify="left").pack(anchor="w", padx=10, pady=(8, 2))

        scroll = tk.Frame(self, bg=C["panel"])
        scroll.pack(fill="both", expand=True, padx=10, pady=5)
        canvas = tk.Canvas(scroll, bg=C["panel"], highlightthickness=0)
        vsb = ttk.Scrollbar(scroll, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        frame = tk.Frame(canvas, bg=C["panel"])
        self._frame = frame
        win = canvas.create_window((0, 0), window=frame, anchor="nw")
        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))

        try:
            words = list(self.app.project.load_config().get("structure_exclude_words", []) or [])
        except Exception:
            words = []
        for w in words:
            self._add_entry(str(w))
        if not words:
            self._add_entry("")

        btns = tk.Frame(self, bg=C["panel"])
        btns.pack(fill="x", padx=10, pady=5)
        self.app._mk_btn(btns, "+ Слово", lambda: self._add_entry(""), C["green"], small=True).pack(side="left", padx=2)

        bottom = tk.Frame(self, bg=C["panel"])
        bottom.pack(fill="x", padx=10, pady=(0, 8))
        self.app._mk_btn(bottom, "Сохранить", self._save, C["green"]).pack(side="left", padx=2)
        self.app._mk_btn(bottom, "Отмена", self.destroy, C["red"]).pack(side="left", padx=2)

    def _add_entry(self, value=""):
        row = tk.Frame(self._frame, bg=C["panel"])
        row.pack(fill="x", pady=1)
        e = tk.Entry(row, bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                     relief="flat", font=("Segoe UI", 9),
                     highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        e.pack(side="left", fill="x", expand=True, padx=(2, 4), ipady=2)
        e.insert(0, value)
        self.app._mk_btn(row, "✕", lambda r=row: self._remove(r), C["red"], small=True).pack(side="right", padx=2)
        self._entries.append((row, e))

    def _remove(self, row):
        self._entries = [(r, e) for r, e in self._entries if r is not row]
        try:
            row.destroy()
        except Exception:
            pass

    def _save(self):
        words = [e.get().strip() for _, e in self._entries if e.get().strip()]
        try:
            self.app.project.save_config({"structure_exclude_words": words})
        except Exception:
            pass
        self.app._log(f"🚫 Слов-исключений «Структуры»: {len(words)}", "ok")
        self.destroy()


class HistoryWindow(tk.Toplevel):
    """Окно «История»: слева темы, справа журнал по датам (v1.17)."""

    def __init__(self, app):
        super().__init__(app)
        self.app = app
        self.title("📜 История изменений тем")
        self.configure(bg=C["panel"])
        self.geometry("1000x640")
        self.minsize(760, 480)
        self._all = []
        self._keys = []
        self._cur = None
        self._newest_first = tk.BooleanVar(value=True)

        top = tk.Frame(self, bg=C["panel"])
        top.pack(fill="x", padx=8, pady=(8, 4))
        tk.Label(top, text="🔍", bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 10)).pack(side="left")
        self.v_search = tk.StringVar()
        self.v_search.trace_add("write", lambda *_: self._apply_filter())
        tk.Entry(top, textvariable=self.v_search, width=34, bg=C["entry"],
                 fg=C["text"], insertbackground=C["text"], relief="flat",
                 font=("Segoe UI", 9), highlightthickness=1,
                 highlightbackground=C["border"],
                 highlightcolor=C["accent"]).pack(side="left", padx=4, ipady=3)
        tk.Checkbutton(top, text="Сначала новые", variable=self._newest_first,
                       command=self._render_right, bg=C["panel"], fg=C["sub"],
                       activebackground=C["panel"], selectcolor=C["entry"],
                       font=("Segoe UI", 9)).pack(side="left", padx=10)
        # v1.23: переключателя слоёв нет — общие всегда сверху, ниже сервер.
        try:
            _hv = self.app.project.load_config().get("history_view", {}) or {}
        except Exception:
            _hv = {}
        tk.Label(top, text="🖥 Сервер:", bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(10, 2))
        self.v_hserver = tk.StringVar(value=_hv.get("server", ""))
        self.cb_hserver = ttk.Combobox(top, textvariable=self.v_hserver,
                                       state="readonly", width=22,
                                       font=("Segoe UI", 9))
        self.cb_hserver.pack(side="left", padx=2)
        self.cb_hserver.bind("<<ComboboxSelected>>", lambda *_: self._on_hserver())
        app._mk_btn(top, "🚫 Исключения…", self._open_exclude, C["accent2"], small=True).pack(side="left", padx=6)
        self.v_count = tk.StringVar(value="")
        tk.Label(top, textvariable=self.v_count, bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 9)).pack(side="right")

        paned = tk.PanedWindow(self, orient="horizontal", bg=C["border"],
                               sashwidth=4, sashrelief="flat")
        paned.pack(fill="both", expand=True, padx=8, pady=4)

        left = tk.Frame(paned, bg=C["panel"])
        paned.add(left, minsize=260)
        sb = tk.Scrollbar(left, orient="vertical")
        sb.pack(side="right", fill="y")
        self.lb = tk.Listbox(left, bg=C["entry"], fg=C["text"],
                             selectbackground=C["sel"], selectforeground="white",
                             font=("Segoe UI", 9), relief="flat",
                             highlightthickness=0, yscrollcommand=sb.set)
        self.lb.pack(side="left", fill="both", expand=True)
        sb.config(command=self.lb.yview)
        self.lb.bind("<<ListboxSelect>>", lambda *_: self._on_select())

        right = tk.Frame(paned, bg=C["panel"])
        paned.add(right, minsize=400)
        self.txt = scrolledtext.ScrolledText(right, bg=C["entry"], fg=C["text"],
                                             font=("Consolas", 9), wrap="word",
                                             relief="flat", highlightthickness=0)
        self.txt.pack(fill="both", expand=True)
        self.txt.tag_configure("h_date", foreground=C["accent"], font=("Consolas", 10, "bold"))
        self.txt.tag_configure("post", foreground=C["accent2"], font=("Consolas", 9, "bold"))
        self.txt.tag_configure("was", foreground=C["red"])
        self.txt.tag_configure("became", foreground=C["green"])
        self.txt.tag_configure("add", foreground=C["green"])
        self.txt.tag_configure("dlt", foreground=C["orange"])
        self.txt.tag_configure("ctx", foreground=C["sub"])
        self.txt.tag_configure("title", foreground=C["text"], font=("Consolas", 11, "bold"))

        bar = tk.Frame(self, bg=C["panel"])
        bar.pack(fill="x", padx=8, pady=(0, 8))
        app._mk_btn(bar, "🌐 Открыть тему", self._open_topic, C["accent2"]).pack(side="left", padx=2)
        app._mk_btn(bar, "📋 Копировать", self._copy, C["accent2"]).pack(side="left", padx=2)
        app._mk_btn(bar, "💾 Сохранить…", self._save, C["green"]).pack(side="left", padx=2)
        app._mk_btn(bar, "🗑 Очистить тему", self._clear_topic, C["red"]).pack(side="left", padx=2)
        app._mk_btn(bar, "🗑 Всю историю…", self._clear_all, C["red"]).pack(side="left", padx=2)
        app._mk_btn(bar, "✖ Закрыть", self.destroy, C["panel"]).pack(side="right", padx=2)

        self._reload()
        if self._keys:
            self.lb.selection_set(0)
            self._on_select()

    def _reload(self):
        try:
            with self.app._history_lock:
                topics = dict((self.app._history or {}).get("topics", {}))
        except Exception:
            topics = {}
        try:
            self._exclude_words = list(self.app.project.load_config().get("history_exclude_words", []) or [])
        except Exception:
            self._exclude_words = []
        rows = []
        servers = set()
        for key, t in topics.items():
            t = t or {}
            evts = t.get("events", [])
            if not evts:
                continue
            scopes = list(t.get("scopes", []) or [])
            for s in scopes:
                if str(s).startswith("server:"):
                    servers.add(str(s)[7:])
            rows.append((key, t.get("title") or "(без названия)",
                         t.get("updated") or "", len(evts),
                         t.get("url") or "", scopes))
        rows.sort(key=lambda r: r[2], reverse=True)
        self._all = rows
        try:
            for sn in (self.app._server_root_nodes() or []):
                if getattr(sn, "title", ""):
                    servers.add(sn.title)
        except Exception:
            pass
        srv_list = sorted(servers)
        try:
            self.cb_hserver["values"] = srv_list
            if self.v_hserver.get() not in srv_list:
                self.v_hserver.set(srv_list[0] if srv_list else "")
        except Exception:
            pass
        self._apply_filter()

    def _apply_filter(self):
        q = (self.v_search.get() or "").strip().lower()
        server = (self.v_hserver.get() or "") if hasattr(self, "v_hserver") else ""
        excl = getattr(self, "_exclude_words", []) or []
        self._keys = []
        self.lb.delete(0, "end")
        gen, srv = [], []
        for key, title, updated, n, _url, scopes in self._all:
            if is_excluded_by_words(title, "", excl):
                continue
            if q and q not in title.lower():
                continue
            layers = history_layers(scopes, server)
            if "general" in layers:
                gen.append((key, title, updated, n))
            if "server" in layers:
                srv.append((key, title, updated, n))
        for key, title, updated, n in gen:
            self._keys.append(key)
            day = (updated or "")[:10]
            self.lb.insert("end", f"📜 {title}  ·  {n} {plural(n, 'событие', 'события', 'событий')}  ·  {day}")
        if server:
            self._keys.append(None)
            self.lb.insert("end", f"───── 🖥 Сервер «{server}» · тем: {len(srv)} ─────")
        for key, title, updated, n in srv:
            self._keys.append(key)
            day = (updated or "")[:10]
            self.lb.insert("end", f"🖥 {title}  ·  {n} {plural(n, 'событие', 'события', 'событий')}  ·  {day}")
        self.v_count.set(f"📜 Общих: {len(gen)}   🖥 Сервера: {len(srv)}")
        if self._keys:
            pick = next((i for i, k in enumerate(self._keys) if k is not None), 0)
            self.lb.selection_clear(0, "end")
            self.lb.selection_set(pick)
            self._on_select()
        else:
            self._cur = None
            self._render_right()

    def _on_hserver(self):
        try:
            self.app.project.save_config({"history_view": {
                "server": self.v_hserver.get()}})
        except Exception:
            pass
        self._apply_filter()

    def _open_exclude(self):
        win = HistoryExcludeWindow(self, self.app)
        try:
            self.wait_window(win)
        except Exception:
            pass
        self._reload()

    def _on_select(self):
        sel = self.lb.curselection()
        if not sel:
            return
        idx = sel[0]
        if 0 <= idx < len(self._keys):
            if self._keys[idx] is None:
                self._cur = None
                self._render_right()
                return
            self._cur = self._keys[idx]
            self._render_right()

    def _cur_topic(self):
        if not self._cur:
            return None
        try:
            with self.app._history_lock:
                return (self.app._history.get("topics", {}) or {}).get(self._cur)
        except Exception:
            return None

    def _render_right(self):
        t = self._cur_topic()
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        if not t:
            self.txt.insert("end", "Выберите тему слева.", "ctx")
            self.txt.config(state="disabled")
            return
        title = t.get("title") or "(без названия)"
        self.txt.insert("end", f"{title}\n", "title")
        if t.get("url"):
            self.txt.insert("end", f"{t['url']}\n", "ctx")
        self.txt.insert("end", "\n")
        for ev in history_sorted_events(t, self._newest_first.get()):
            for text, tag in history_event_segments(ev):
                self.txt.insert("end", text + "\n", tag)
            self.txt.insert("end", "\n")
        self.txt.config(state="disabled")

    def _open_topic(self):
        t = self._cur_topic()
        url = (t or {}).get("url", "")
        if url:
            try:
                webbrowser.open(url)
            except Exception:
                pass

    def _copy(self):
        try:
            s = self.txt.get("1.0", "end-1c")
            self.clipboard_clear()
            self.clipboard_append(s)
        except Exception:
            pass

    def _save(self):
        t = self._cur_topic()
        if not t:
            return
        p = filedialog.asksaveasfilename(
            title="Сохранить историю темы", defaultextension=".txt",
            filetypes=[("Text", "*.txt"), ("All", "*.*")],
            initialfile=f"history_{safe_name(t.get('title', 'topic') or 'topic', 40)}.txt",
        )
        if not p:
            return
        try:
            Path(p).write_text(self.txt.get("1.0", "end-1c"), encoding="utf-8")
        except Exception as e:
            messagebox.showerror("История", str(e))

    def _clear_topic(self):
        if not self._cur:
            return
        t = self._cur_topic()
        name = (t or {}).get("title", "")
        if not messagebox.askyesno("История", f"Удалить всю историю темы?\n«{name}»"):
            return
        self.app._clear_history_topic(self._cur)
        self._cur = None
        self._reload()

    def _clear_all(self):
        if not messagebox.askyesno("История", "Удалить ВСЮ историю изменений?"):
            return
        self.app._clear_history_all()
        self._cur = None
        self._reload()


class RichLog(tk.Frame):
    def __init__(self, parent, **kw):
        super().__init__(parent, bg=C["panel"], **kw)
        self._build()
        self._search_pos = "1.0"

    def _build(self):
        tb = tk.Frame(self, bg=C["panel2"])
        tb.pack(fill="x", side="top")

        def mk(text, cmd, color, tip=""):
            b = tk.Button(tb, text=text, command=cmd,
                          bg=color, fg="white", activebackground=color,
                          font=("Segoe UI", 8, "bold"),
                          relief="flat", cursor="hand2", padx=6, pady=1, bd=0)
            b.pack(side="left", padx=1, pady=1)
            if tip:
                self._tooltip(b, tip)
            return b

        mk("📋 Копир. всё", self.copy_all, C["accent2"], "Скопировать весь лог")
        mk("📄 Копир. выд.", self.copy_sel, C["accent"], "Скопировать выделение")
        mk("💾 Сохранить", self.save_to_file, C["green"], "Сохранить лог в файл")
        mk("🗑 Очистить", self.clear, C["red"], "Очистить лог")
        mk("⬇ Вниз", self.scroll_end, C["panel"], "Прокрутить в конец")

        tk.Label(tb, text="  🔍", bg=C["panel2"], fg=C["sub"], font=("Segoe UI", 9)).pack(side="left", padx=(6, 0))
        self.v_search = tk.StringVar()
        e = tk.Entry(tb, textvariable=self.v_search, width=18,
                     bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                     relief="flat", font=("Segoe UI", 8),
                     highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        e.pack(side="left", padx=2, ipady=2)
        e.bind("<Return>", lambda *_: self.find_next())
        e.bind("<KeyRelease>", lambda *_: self._highlight_all())

        mk("↓", self.find_next, C["panel"], "Следующее вхождение")
        mk("↑", self.find_prev, C["panel"], "Предыдущее вхождение")

        self.v_autoscroll = tk.BooleanVar(value=True)
        tk.Checkbutton(tb, text="Автопрокрутка", variable=self.v_autoscroll,
                       bg=C["panel2"], fg=C["sub"], activebackground=C["panel2"],
                       selectcolor=C["entry"], font=("Segoe UI", 8)).pack(side="right", padx=6)

        self.txt = scrolledtext.ScrolledText(
            self, state="disabled", wrap="word",
            bg=C["entry"], fg=C["text"], font=("Consolas", 8),
            insertbackground=C["text"], relief="flat", highlightthickness=0, undo=False,
        )
        self.txt.pack(fill="both", expand=True, padx=2, pady=(0, 2))

        self.txt.tag_configure("ok", foreground=C["green"])
        self.txt.tag_configure("err", foreground=C["red"])
        self.txt.tag_configure("warn", foreground=C["yellow"])
        self.txt.tag_configure("info", foreground=C["accent2"])
        self.txt.tag_configure("head", foreground=C["accent"])
        self.txt.tag_configure("debug", foreground=C["debug"])
        self.txt.tag_configure("normal", foreground=C["text"])
        self.txt.tag_configure("search", background=C["search"], foreground="#000000")
        self.txt.tag_configure("search_current", background="#ff9800", foreground="#000000")

        self.txt.bind("<Control-c>", lambda e: (self.copy_sel(), "break"))
        self.txt.bind("<Control-C>", lambda e: (self.copy_sel(), "break"))
        self.txt.bind("<Control-a>", lambda e: (self._select_all(), "break"))
        self.txt.bind("<Control-A>", lambda e: (self._select_all(), "break"))
        self.txt.bind("<Control-f>", lambda e: (e.widget.focus_set(), "break"))
        self.txt.bind("<Button-3>", self._show_ctx_menu)

        self._ctx = tk.Menu(self.txt, tearoff=0, bg=C["panel"], fg=C["text"],
                             activebackground=C["sel"], activeforeground="white")
        self._ctx.add_command(label="📄 Копировать выделение", command=self.copy_sel)
        self._ctx.add_command(label="📋 Копировать всё", command=self.copy_all)
        self._ctx.add_command(label="✅ Выделить всё", command=self._select_all)
        self._ctx.add_separator()
        self._ctx.add_command(label="💾 Сохранить в файл…", command=self.save_to_file)
        self._ctx.add_command(label="🗑 Очистить лог", command=self.clear)

    def _show_ctx_menu(self, event):
        try:
            self._ctx.tk_popup(event.x_root, event.y_root)
        finally:
            self._ctx.grab_release()

    def _tooltip(self, widget, text):
        tip = {"win": None}

        def enter(_):
            if tip["win"]:
                return
            x = widget.winfo_rootx() + 20
            y = widget.winfo_rooty() + widget.winfo_height() + 2
            w = tk.Toplevel(widget)
            w.wm_overrideredirect(True)
            w.wm_geometry(f"+{x}+{y}")
            tk.Label(w, text=text, bg="#333", fg="white", font=("Segoe UI", 8), padx=6, pady=2).pack()
            tip["win"] = w

        def leave(_):
            if tip["win"]:
                tip["win"].destroy()
                tip["win"] = None

        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)

    def log(self, msg: str, tag="normal"):
        def _do():
            self.txt.config(state="normal")
            self.txt.insert("end", msg + "\n", tag)
            if self.v_autoscroll.get():
                self.txt.see("end")
            self.txt.config(state="disabled")
        try:
            self.after(0, _do)
        except Exception:
            pass

    def _select_all(self):
        self.txt.tag_add("sel", "1.0", "end-1c")
        self.txt.focus_set()

    def copy_sel(self):
        try:
            s = self.txt.selection_get()
        except Exception:
            s = ""
        if not s:
            self.copy_all()
            return
        self.clipboard_clear()
        self.clipboard_append(s)

    def copy_all(self):
        s = self.txt.get("1.0", "end-1c")
        self.clipboard_clear()
        self.clipboard_append(s)

    def clear(self):
        self.txt.config(state="normal")
        self.txt.delete("1.0", "end")
        self.txt.config(state="disabled")
        self._search_pos = "1.0"

    def scroll_end(self):
        self.txt.see("end")

    def save_to_file(self):
        p = filedialog.asksaveasfilename(
            title="Сохранить лог", defaultextension=".log",
            filetypes=[("Log files", "*.log *.txt"), ("All files", "*.*")],
            initialfile=f"forum_log_{time.strftime('%Y%m%d_%H%M%S')}.log",
        )
        if not p:
            return
        try:
            Path(p).write_text(self.txt.get("1.0", "end-1c"), encoding="utf-8")
            messagebox.showinfo("Сохранено", f"Лог сохранён:\n{p}")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def _highlight_all(self):
        q = self.v_search.get()
        self.txt.tag_remove("search", "1.0", "end")
        self.txt.tag_remove("search_current", "1.0", "end")
        if not q:
            return
        start = "1.0"
        while True:
            idx = self.txt.search(q, start, nocase=True, stopindex="end")
            if not idx:
                break
            end = f"{idx}+{len(q)}c"
            self.txt.tag_add("search", idx, end)
            start = end

    def find_next(self):
        q = self.v_search.get()
        if not q:
            return
        self.txt.tag_remove("search_current", "1.0", "end")
        idx = self.txt.search(q, self._search_pos, nocase=True, stopindex="end")
        if not idx:
            self._search_pos = "1.0"
            idx = self.txt.search(q, "1.0", nocase=True, stopindex="end")
            if not idx:
                return
        end = f"{idx}+{len(q)}c"
        self.txt.tag_add("search_current", idx, end)
        self.txt.see(idx)
        self._search_pos = end

    def find_prev(self):
        q = self.v_search.get()
        if not q:
            return
        self.txt.tag_remove("search_current", "1.0", "end")
        idx = self.txt.search(q, self._search_pos, nocase=True, backwards=True, stopindex="1.0")
        if not idx:
            idx = self.txt.search(q, "end", nocase=True, backwards=True, stopindex="1.0")
            if not idx:
                return
        end = f"{idx}+{len(q)}c"
        self.txt.tag_add("search_current", idx, end)
        self.txt.see(idx)
        self._search_pos = idx

class ProcessDispatcher(tk.Frame):
    STATE_ICONS = {
        "pending": "⏳",
        "running": "🔄",
        "done": "✅",
        "error": "❌",
        "warn": "⚠️",
        "skipped": "⏭",
        "paused": "⏸",
        "cancelled": "⛔",
        "blocked": "🚫",
        "changed": "♻️",
        "new": "🆕",
        "cached": "🔒",
    }

    def __init__(self, parent, **kw):
        super().__init__(parent, bg=C["panel"], **kw)
        self._nodes = {}
        self._emitter = None
        self._build()

    def _build(self):
        hdr = tk.Frame(self, bg=C["panel2"])
        hdr.pack(fill="x")
        tk.Label(hdr, text="📊 Диспетчер процессов", bg=C["panel2"], fg=C["accent"],
                 font=("Segoe UI", 9, "bold")).pack(side="left", padx=6, pady=2)
        self._mk_btn(hdr, "⊞", self._expand_all)
        self._mk_btn(hdr, "⊟", self._collapse_all)
        self._mk_btn(hdr, "🗑", self._clear_finished)
        self._mk_btn(hdr, "🗑 Всё", self.clear_all)
        self._mk_btn(hdr, "📋", self.copy_all)
        self.tree = ttk.Treeview(self, style="FM.Treeview", show="tree headings", selectmode="browse")
        self.tree["columns"] = ("status", "progress", "time", "detail")
        self.tree.column("#0", width=280, minwidth=140, stretch=True)
        self.tree.column("status", width=36, minwidth=28, stretch=False, anchor="center")
        self.tree.column("progress", width=110, minwidth=60, stretch=False, anchor="center")
        self.tree.column("time", width=56, minwidth=40, stretch=False, anchor="center")
        self.tree.column("detail", width=140, minwidth=70, stretch=True, anchor="w")
        self.tree.heading("#0", text="Процесс / Шаг")
        self.tree.heading("status", text="")
        self.tree.heading("progress", text="Прогресс")
        self.tree.heading("time", text="Время")
        self.tree.heading("detail", text="Детали")
        vsb = ttk.Scrollbar(self, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self.tree.bind("<Control-c>", lambda e: (self.copy_all(), "break"))
        self.tree.bind("<Control-C>", lambda e: (self.copy_all(), "break"))
        for st in self.STATE_ICONS:
            color = C["sub"]
            if st == "running":
                color = C["yellow"]
            elif st == "done":
                color = C["green"]
            elif st == "error":
                color = C["red"]
            elif st == "warn":
                color = C["orange"]
            elif st == "changed":
                color = C["orange"]
            elif st == "new":
                color = C["accent2"]
            self.tree.tag_configure(st, foreground=color)

    def _mk_btn(self, parent, text, cmd):
        b = tk.Button(parent, text=text, command=cmd, bg=C["panel"], fg=C["text"],
                      font=("Segoe UI", 8), relief="flat", cursor="hand2", padx=4, pady=1)
        b.pack(side="left", padx=2)
        return b

    def attach(self, emitter):
        self._emitter = emitter
        self._poll()

    def _poll(self):
        if self._emitter is not None:
            try:
                while True:
                    evt = self._emitter.q.get_nowait()
                    self._handle_event(evt)
            except queue.Empty:
                pass
            except Exception:
                pass
        self.after(50, self._poll)

    def _handle_event(self, evt):
        t = evt.get("type")
        if t == "process_started":
            self._add_node(evt["id"], "", evt.get("label", ""), "🚀", "pending")
        elif t == "step_added":
            self._add_node(evt["id"], evt.get("parent", ""), evt.get("label", ""), evt.get("icon", ""), "pending")
        elif t == "step_state":
            self._update_node(evt["id"], state=evt.get("state"), detail=evt.get("detail"))
        elif t == "step_metric":
            self._update_node(evt["id"], metric_key=evt.get("key"), metric_val=evt.get("value"))
        elif t == "process_finished":
            elapsed = float(evt.get("elapsed") or 0)
            self._update_node(
                evt["id"],
                state=evt.get("state"),
                detail=f"⏱ {elapsed:.1f}с",
                metric_key="elapsed",
                metric_val=f"{elapsed:.1f}с"
            )

    def _add_node(self, node_id, parent_id, label, icon, state):
        parent_iid = self._nodes.get(parent_id, "")
        try:
            iid = self.tree.insert(parent_iid, "end", text=f"{icon} {label}",
                                   values=(self.STATE_ICONS.get(state, ""), "", "", ""),
                                   tags=(state,), open=True)
            self._nodes[node_id] = iid
            if parent_iid:
                self.tree.item(parent_iid, open=True)
        except Exception:
            pass

    def _update_node(self, node_id, state=None, detail=None, metric_key=None, metric_val=None):
        iid = self._nodes.get(node_id)
        if not iid:
            return
        try:
            vals = list(self.tree.item(iid, "values"))
            if state:
                vals[0] = self.STATE_ICONS.get(state, state)
                self.tree.item(iid, tags=(state,))
            if detail:
                vals[3] = detail
            if metric_key == "elapsed":
                vals[2] = str(metric_val)
            elif metric_key == "progress":
                vals[1] = str(metric_val)
            self.tree.item(iid, values=vals)
        except Exception:
            pass

    def _expand_all(self):
        def exp(iid):
            self.tree.item(iid, open=True)
            for ch in self.tree.get_children(iid):
                exp(ch)
        for iid in self.tree.get_children():
            exp(iid)

    def _collapse_all(self):
        def col(iid):
            self.tree.item(iid, open=False)
            for ch in self.tree.get_children(iid):
                col(ch)
        for iid in self.tree.get_children():
            col(iid)

    def _clear_finished(self):
        for iid in list(self.tree.get_children()):
            tags = self.tree.item(iid, "tags")
            if any(t in tags for t in ("done", "error", "cancelled")):
                self.tree.delete(iid)
                for nid, tid in list(self._nodes.items()):
                    if tid == iid:
                        del self._nodes[nid]

    def clear_all(self):
        for iid in self.tree.get_children():
            self.tree.delete(iid)
        self._nodes.clear()

    def copy_all(self):
        lines = ["Процесс / Шаг | Статус | Прогресс | Время | Детали"]
        for iid in self.tree.get_children():
            self._collect_copy_item(iid, 0, lines)
        self.clipboard_clear()
        self.clipboard_append("\n".join(lines))

    def _collect_copy_item(self, iid, level, lines):
        text = str(self.tree.item(iid, "text") or "")
        values = list(self.tree.item(iid, "values") or [])
        while len(values) < 4:
            values.append("")
        values = [str(v) for v in values[:4]]
        lines.append(("    " * level) + text + " | " + " | ".join(values))
        for child in self.tree.get_children(iid):
            self._collect_copy_item(child, level + 1, lines)

class KeywordConfigWindow(tk.Toplevel):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("Ключевые слова автоотметки")
        self.configure(bg=C["panel"])
        self.geometry("650x650")
        self.transient(parent)
        self.grab_set()
        self._entries_node = []
        self._entries_thread = []
        self._entries_exclude = []

        scroll = tk.Frame(self, bg=C["panel"])
        scroll.pack(fill="both", expand=True, padx=10, pady=5)

        canvas = tk.Canvas(scroll, bg=C["panel"], highlightthickness=0)
        vsb = ttk.Scrollbar(scroll, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._frame = tk.Frame(canvas, bg=C["panel"])
        self._canvas = canvas
        self._canvas_window = canvas.create_window((0, 0), window=self._frame, anchor="nw")
        self._frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(self._canvas_window, width=e.width))

        cfg = self.app.project.load_config()

        tk.Label(self._frame, text="Ключевые слова для разделов:",
                 bg=C["panel"], fg=C["accent"], font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(5, 2))
        node_kw = cfg.get("auto_check_section_keywords", [])
        for kw in node_kw:
            if isinstance(kw, dict):
                self._add_node_entry(kw.get("keyword", ""), kw.get("whole", False), kw.get("pos", "any"))
            else:
                self._add_node_entry(str(kw))
        if not node_kw:
            self._add_node_entry("")
        btns1 = tk.Frame(self._frame, bg=C["panel"])
        btns1.pack(anchor="w", pady=5)
        self.app._mk_btn(btns1, "+ Раздел", lambda: self._add_node_entry(""), C["green"], small=True).pack(side="left", padx=2)

        tk.Label(self._frame, text="Ключевые слова для тем:",
                 bg=C["panel"], fg=C["accent2"], font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10, 2))
        thread_kw = cfg.get("auto_check_thread_keywords", [])
        for item in thread_kw:
            if isinstance(item, dict):
                self._add_thread_entry(item.get("keyword", ""), item.get("date", ""),
                                       item.get("whole", False), item.get("pos", "any"))
            else:
                self._add_thread_entry(str(item), "")
        if not thread_kw:
            self._add_thread_entry("", "")
        btns2 = tk.Frame(self._frame, bg=C["panel"])
        btns2.pack(anchor="w", pady=5)
        self.app._mk_btn(btns2, "+ Тема", lambda: self._add_thread_entry("", ""), C["green"], small=True).pack(
            side="left", padx=2)

        tk.Label(self._frame, text="🚫 Исключения (есть в теме/разделе — не отмечать):",
                 bg=C["panel"], fg=C["red"], font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(10, 2))
        thread_ex = cfg.get("auto_check_thread_exclude", [])
        for w in thread_ex:
            self._add_exclude_entry(str(w) if not isinstance(w, dict) else w.get("word", ""))
        if not thread_ex:
            self._add_exclude_entry("")
        btns3 = tk.Frame(self._frame, bg=C["panel"])
        btns3.pack(anchor="w", pady=5)
        self.app._mk_btn(btns3, "+ Исключение", lambda: self._add_exclude_entry(""), C["green"], small=True).pack(
            side="left", padx=2)

        bottom = tk.Frame(self, bg=C["panel"])
        bottom.pack(fill="x", padx=10, pady=5)
        self.app._mk_btn(bottom, "Сохранить", self._save, C["green"]).pack(side="left", padx=2)
        self.app._mk_btn(bottom, "Отмена", self.destroy, C["red"]).pack(side="left", padx=2)

    def _add_node_entry(self, value="", whole=False, pos="any"):
        row = tk.Frame(self._frame, bg=C["panel"])
        row.pack(fill="x", pady=1)
        e = tk.Entry(row, bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                     relief="flat", font=("Segoe UI", 9),
                     highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        e.pack(side="left", fill="x", expand=True, padx=(2, 4), ipady=2)
        e.insert(0, value)
        wv = tk.BooleanVar(value=bool(whole))
        tk.Checkbutton(row, text="целое", variable=wv, bg=C["panel"], fg=C["sub"],
                       activebackground=C["panel"], selectcolor=C["entry"],
                       font=("Segoe UI", 8)).pack(side="left", padx=2)
        pv = tk.StringVar(value={"any": "везде", "start": "в начале", "end": "в конце"}.get(pos, "везде"))
        ttk.Combobox(row, textvariable=pv, values=["везде", "в начале", "в конце"],
                     state="readonly", width=9, font=("Segoe UI", 8)).pack(side="left", padx=2)
        self.app._mk_btn(row, "✕", lambda r=row, en=e: self._remove_entry(r, en, self._entries_node), C["red"], small=True).pack(side="right", padx=2)
        self._entries_node.append((row, e, wv, pv))

    def _add_thread_entry(self, value="", date="", whole=False, pos="any"):
        row = tk.Frame(self._frame, bg=C["panel"])
        row.pack(fill="x", pady=1)
        e = tk.Entry(row, bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                     relief="flat", font=("Segoe UI", 9),
                     highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent2"])
        e.pack(side="left", fill="x", expand=True, padx=(2, 4), ipady=2)
        e.insert(0, value)
        tk.Label(row, text="с:", bg=C["panel"], fg=C["accent2"], font=("Segoe UI", 9)).pack(side="left", padx=(4, 0))
        d = tk.Entry(row, bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                     relief="flat", font=("Segoe UI", 9), width=12,
                     highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent2"])
        d.pack(side="left", padx=2, ipady=2)
        d.insert(0, date)
        tk.Label(row, text="(дд.мм.гггг)", bg=C["panel"], fg=C["sub"], font=("Segoe UI", 8)).pack(side="left",
                                                                                                  padx=(0, 4))
        wv = tk.BooleanVar(value=bool(whole))
        tk.Checkbutton(row, text="целое", variable=wv, bg=C["panel"], fg=C["sub"],
                       activebackground=C["panel"], selectcolor=C["entry"],
                       font=("Segoe UI", 8)).pack(side="left", padx=2)
        pv = tk.StringVar(value={"any": "везде", "start": "в начале", "end": "в конце"}.get(pos, "везде"))
        ttk.Combobox(row, textvariable=pv, values=["везде", "в начале", "в конце"],
                     state="readonly", width=9, font=("Segoe UI", 8)).pack(side="left", padx=2)
        self.app._mk_btn(row, "✕", lambda r=row, en=e: self._remove_entry(r, en, self._entries_thread), C["red"],
                         small=True).pack(side="right", padx=2)
        self._entries_thread.append((row, e, d, wv, pv))

    def _add_exclude_entry(self, value=""):
        row = tk.Frame(self._frame, bg=C["panel"])
        row.pack(fill="x", pady=1)
        e = tk.Entry(row, bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                     relief="flat", font=("Segoe UI", 9),
                     highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["red"])
        e.pack(side="left", fill="x", expand=True, padx=(2, 4), ipady=2)
        e.insert(0, value)
        self.app._mk_btn(row, "✕", lambda r=row, en=e: self._remove_entry(r, en, self._entries_exclude), C["red"], small=True).pack(side="right", padx=2)
        self._entries_exclude.append((row, e))

    def _remove_entry(self, row, entry, lst):
        for i, item in enumerate(lst):
            if item[0] is row:
                lst.pop(i)
                break
        row.destroy()

    def _save(self):
        _pos_back = {"в начале": "start", "в конце": "end"}
        node_kw = []
        for item in self._entries_node:
            _kw = item[1].get().strip()
            if _kw:
                node_kw.append({"keyword": _kw, "whole": bool(item[2].get()),
                                "pos": _pos_back.get(item[3].get(), "any")})
        thread_kw = []
        for item in self._entries_thread:
            kw_entry = item[1]
            date_entry = item[2]
            kw = kw_entry.get().strip()
            if kw:
                thread_kw.append({"keyword": kw, "date": date_entry.get().strip(),
                                  "whole": bool(item[3].get()),
                                  "pos": _pos_back.get(item[4].get(), "any")})
        thread_ex = [e.get().strip() for _, e in self._entries_exclude if e.get().strip()]
        self.app.project.save_config({
            "auto_check_section_keywords": node_kw,
            "auto_check_thread_keywords": thread_kw,
            "auto_check_thread_exclude": thread_ex,
        })
        set_autocheck_keywords(node_kw, thread_kw, thread_ex)
        try:
            self.app._apply_exclusions_to_checks()
        except Exception:
            pass
        self.app._log("Ключевые слова автоотметки сохранены и применены", "ok")
        self.destroy()

class TopicFilterWindow(tk.Toplevel):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("Фильтр тем")
        self.configure(bg=C["panel"])
        self.geometry("750x550")
        self.transient(parent)
        self.grab_set()
        self._rows = []
        scroll = tk.Frame(self, bg=C["panel"])
        scroll.pack(fill="both", expand=True, padx=10, pady=5)
        canvas = tk.Canvas(scroll, bg=C["panel"], highlightthickness=0)
        vsb = ttk.Scrollbar(scroll, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")
        self._frame = tk.Frame(canvas, bg=C["panel"])
        self._canvas = canvas
        self._canvas_window = canvas.create_window((0, 0), window=self._frame, anchor="nw")
        self._frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(self._canvas_window, width=e.width))
        ocr_top = tk.Frame(self._frame, bg=C["panel"])
        ocr_top.pack(fill="x", pady=(2, 6))
        tk.Label(ocr_top, text="Порог OCR (симв.):", bg=C["panel"], fg=C["sub"], font=("Segoe UI", 9)).pack(side="left")
        cfg = self.app.project.load_config()
        self._ocr_entry = tk.Entry(ocr_top, bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                                   relief="flat", font=("Segoe UI", 9), width=8,
                                   highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        self._ocr_entry.pack(side="left", padx=(6, 0), ipady=2)
        self._ocr_entry.insert(0, str(cfg.get("ocr_min_chars", 0)))
        filters = cfg.get("topic_filters", [])
        for f in filters:
            self._add_row(f)
        if not filters:
            self._add_row({})
        btns = tk.Frame(self._frame, bg=C["panel"])
        btns.pack(anchor="w", pady=5)
        self.app._mk_btn(btns, "+ Добавить", lambda: self._add_row({}), C["green"], small=True).pack(side="left", padx=2)
        bottom = tk.Frame(self, bg=C["panel"])
        bottom.pack(fill="x", padx=10, pady=5)
        self.app._mk_btn(bottom, "Сохранить", self._save, C["green"]).pack(side="left", padx=2)
        self.app._mk_btn(bottom, "Отмена", self.destroy, C["red"]).pack(side="left", padx=2)

    def _add_row(self, data=None):
        data = data or {}
        row = tk.Frame(self._frame, bg=C["panel2"], padx=5, pady=5)
        row.pack(fill="x", pady=2)
        r1 = tk.Frame(row, bg=C["panel2"])
        r1.pack(fill="x")
        tk.Label(r1, text="Ключевая фраза:", bg=C["panel2"], fg=C["sub"], font=("Segoe UI", 8)).pack(side="left")
        kw_entry = tk.Entry(r1, bg=C["entry"], fg=C["text"], insertbackground=C["text"],
                            relief="flat", font=("Segoe UI", 9),
                            highlightthickness=1, highlightbackground=C["border"], highlightcolor=C["accent"])
        kw_entry.pack(side="left", fill="x", expand=True, padx=5, ipady=2)
        kw_entry.insert(0, data.get("keyword", ""))
        r2 = tk.Frame(row, bg=C["panel2"])
        r2.pack(fill="x", pady=(3, 0))
        img_var = tk.BooleanVar(value=data.get("parse_images", True))
        txt_var = tk.BooleanVar(value=data.get("parse_text", True))
        tk.Checkbutton(r2, text="изображение", variable=img_var,
                       bg=C["panel2"], fg=C["text"], activebackground=C["panel2"],
                       selectcolor=C["entry"], font=("Segoe UI", 8)).pack(side="left", padx=5)
        tk.Checkbutton(r2, text="текст", variable=txt_var,
                       bg=C["panel2"], fg=C["text"], activebackground=C["panel2"],
                       selectcolor=C["entry"], font=("Segoe UI", 8)).pack(side="left", padx=5)
        self.app._mk_btn(r2, "✕", lambda r=row: self._remove_row(r), C["red"], small=True).pack(side="right", padx=2)
        self._rows.append((row, kw_entry, img_var, txt_var))

    def _remove_row(self, row):
        for i, (r, *_) in enumerate(self._rows):
            if r is row:
                self._rows.pop(i)
                break
        row.destroy()

    def _save(self):
        filters = []
        for row, kw_e, img_v, txt_v in self._rows:
            kw = kw_e.get().strip()
            if not kw:
                continue
            filters.append({
                "keyword": kw,
                "parse_images": img_v.get(),
                "parse_text": txt_v.get(),
            })
        try:
            ocr_min = int(self._ocr_entry.get().strip())
        except ValueError:
            ocr_min = 0
        self.app.project.save_config({"topic_filters": filters, "ocr_min_chars": ocr_min})
        self.app._topic_filters = filters
        self.app._log("Фильтр тем сохранён", "ok")
        self.destroy()

class ParticipationWindow(tk.Toplevel):
    def __init__(self, parent, app):
        super().__init__(parent)
        self.app = app
        self.title("Участвующие разделы")
        self.configure(bg=C["panel"])
        self.geometry("680x620")
        self.transient(parent)
        self.grab_set()

        self._rows = []

        top = tk.Frame(self, bg=C["panel"])
        top.pack(fill="x", padx=10, pady=5)

        tk.Label(
            top,
            text="Разделы и все внутренние подразделы, участвующие в Этапе 2 и формировании документов:",
            bg=C["panel"], fg=C["accent"], font=("Segoe UI", 9, "bold"),
            wraplength=640, justify="left"
        ).pack(anchor="w")

        btns = tk.Frame(top, bg=C["panel"])
        btns.pack(anchor="w", pady=6)

        self.app._mk_btn(btns, "✅ Выбрать все", self._select_all, C["green"], small=True).pack(side="left", padx=2)
        self.app._mk_btn(btns, "☐ Убрать все", self._unselect_all, C["red"], small=True).pack(side="left", padx=2)

        scroll = tk.Frame(self, bg=C["panel"])
        scroll.pack(fill="both", expand=True, padx=10, pady=5)

        canvas = tk.Canvas(scroll, bg=C["panel"], highlightthickness=0)
        vsb = ttk.Scrollbar(scroll, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        self._frame = tk.Frame(canvas, bg=C["panel"])
        self._canvas = canvas
        self._canvas_window = canvas.create_window((0, 0), window=self._frame, anchor="nw")

        self._frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(self._canvas_window, width=e.width))

        selected = self.app._participation_selected_urls()
        items = self.app._participation_items()

        if not items:
            tk.Label(
                self._frame,
                text="Нет разделов. Сначала выполните сканирование.",
                bg=C["panel"], fg=C["yellow"], font=("Segoe UI", 9)
            ).pack(anchor="w", pady=10)

        for node, depth in items:
            self._add_row(node, depth, selected)

        bottom = tk.Frame(self, bg=C["panel"])
        bottom.pack(fill="x", padx=10, pady=5)

        self.app._mk_btn(bottom, "Сохранить", self._save, C["green"]).pack(side="left", padx=2)
        self.app._mk_btn(bottom, "Отмена", self.destroy, C["red"]).pack(side="left", padx=2)

    def _add_row(self, node, depth, selected):
        row = tk.Frame(self._frame, bg=C["panel"])
        row.pack(fill="x", pady=1)

        var = tk.BooleanVar(value=True if selected is None else node.url in selected)
        indent = "    " * depth

        cb = tk.Checkbutton(
            row,
            text=f"{indent}{node.title}",
            variable=var,
            bg=C["panel"], fg=C["text"], activebackground=C["panel"],
            selectcolor=C["entry"], font=("Segoe UI", 9),
            anchor="w", justify="left", padx=2, pady=0, relief="flat"
        )

        cb.pack(fill="x", anchor="w")
        self._rows.append((node, var))

    def _select_all(self):
        for _, var in self._rows:
            var.set(True)

    def _unselect_all(self):
        for _, var in self._rows:
            var.set(False)

    def _save(self):
        selected_urls = []
        seen = set()

        for node, var in self._rows:
            if var.get():
                url = node.url
                if url not in seen:
                    seen.add(url)
                    selected_urls.append(url)

        self.app.project.save_config({"participation_sections": selected_urls})
        self.app._log(f"🎯 Участвующих разделов сохранено: {len(selected_urls)}", "ok")
        self.destroy()

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"Форум-Менеджер {PARSER_VERSION}")
        self.geometry("1400x900")
        self.minsize(1000, 640)
        self.configure(bg=C["bg"])

        self.http = HttpClient()
        self.http.cookie_harvester = self._harvest_browser_cookies
        self.http.cookies_persister = self._persist_harvested_cookies
        self.http.browser_ensurer = self._ensure_fetch_browser
        self.executor = ThreadPoolExecutor(max_workers=MAX_WORKERS)
        self.project = Project(FORUM_BASE_URL)
        self.nodes = []
        self._scanner = None
        self._login_browser = None
        self._fetch_browser = None
        self._fetch_browser_lock = threading.Lock()
        self._running = False
        self._logged_in = False
        self._abort = threading.Event()
        self.http.abort_event = self._abort  # v1.21: «Стоп» прерывает и паузы рубильника
        self._checks = {}
        self._nmap = {}
        self._depths = {}
        self._thread_map = {}
        self._thread_checks = {}
        self._history = {"topics": {}}
        self._history_lock = threading.Lock()
        self._history_loaded_for = None
        self._sel_url = None
        self._show_debug = tk.BooleanVar(value=True)
        self._emitter = None
        cfg = self.project.load_config()
        self._prompt_path = str(cfg.get("prompt_path") or "").strip()
        self._topic_filters = cfg.get("topic_filters", [])
        self._abbreviations = cfg.get("abbreviations", {})
        self._telegraph = cfg.get("telegraph_replacements", {})
        self.v_fast = tk.BooleanVar(value=bool(cfg.get("fast_path_enabled", FAST_PATH_ENABLED_DEFAULT)))
        self.v_fast.trace_add("write", lambda *_: self._on_fast_toggle())
        self.http.fast_enabled = self.v_fast.get()
        self.v_stealth = tk.BooleanVar(value=bool(cfg.get("stealth_mode", False)))
        self.v_stealth.trace_add("write", lambda *_: self._on_stealth_toggle())
        self._apply_stealth(self.v_stealth.get(), silent=True)
        node_kw = cfg.get("auto_check_section_keywords", [])
        thread_kw = cfg.get("auto_check_thread_keywords", [])
        thread_ex = cfg.get("auto_check_thread_exclude", [])
        set_autocheck_keywords(node_kw, thread_kw, thread_ex)
        try:
            self._workers_n = max(1, min(64, int(cfg.get("workers", MAX_WORKERS))))
        except Exception:
            self._workers_n = MAX_WORKERS
        if self._workers_n != MAX_WORKERS:
            try:
                _old_ex = self.executor
                self.executor = ThreadPoolExecutor(max_workers=self._workers_n)
                _old_ex.shutdown(wait=False)
            except Exception:
                pass
        if not cfg.get("abbreviations"):
            self.project.save_config({"abbreviations": {}})
        if not cfg.get("telegraph_replacements"):
            self.project.save_config({"telegraph_replacements": {}})
        if not cfg.get("auto_check_section_keywords"):
            self.project.save_config({"auto_check_section_keywords": ["ОБРАЩЕНИЕ О ДАЧЕ РАЗЪЯСНЕНИЙ"]})
        if not cfg.get("auto_check_thread_keywords"):
            self.project.save_config({"auto_check_thread_keywords": ["ОБРАЩЕНИЕ О ДАЧЕ РАЗЪЯСНЕНИЙ"]})
        if not cfg.get("topic_filters"):
            self.project.save_config({"topic_filters": []})
        if "fast_path_enabled" not in cfg:
            self.project.save_config({"fast_path_enabled": FAST_PATH_ENABLED_DEFAULT})
        if "stealth_mode" not in cfg:
            self.project.save_config({"stealth_mode": False})
        self._build()
        self.protocol("WM_DELETE_WINDOW", self._on_close)
        self._update_login_indicator()
        self.after(300, self._auto_load_state)
        self.after(2000, self._poll_transport_stats)
        self.after(600, self._update_history_badge)

    def _build(self):
        self._build_topbar()
        self._build_main()
        self._build_statusbar()

    def _mk_btn(self, parent, text, cmd, color, small=False):
        fs = 8 if small else 9
        pd = (5, 2) if small else (10, 4)
        return tk.Button(parent, text=text, command=cmd,
                          bg=color, fg="white", activebackground=color,
                          font=("Segoe UI", fs, "bold"), relief="flat",
                          cursor="hand2", padx=pd[0], pady=pd[1])

    def _server_dir_names(self, server_nodes):
        name_counts = {}

        for sn in server_nodes:
            base_dir = safe_name(sn.title) or "server"
            name_counts[base_dir] = name_counts.get(base_dir, 0) + 1

        used_server_dirs = set()
        result = {}

        for sn in server_nodes:
            base_dir = safe_name(sn.title) or "server"
            dir_name = base_dir

            if name_counts.get(base_dir, 0) > 1:
                digest = hashlib.md5((sn.url or sn.title or "").encode("utf-8")).hexdigest()[:8]
                dir_name = f"{base_dir}_{digest}"

                counter = 1
                while dir_name in used_server_dirs:
                    counter += 1
                    dir_name = f"{base_dir}_{digest}_{counter}"
            else:
                counter = 1
                while dir_name in used_server_dirs:
                    counter += 1
                    dir_name = f"{base_dir}_{counter}"

            used_server_dirs.add(dir_name)
            result[id(sn)] = dir_name

        return result

    def _uncheck_threads_by_url(self, urls_to_uncheck):
        urls_set = set(urls_to_uncheck)
        changed = False
        for thr_iid, thr in self._thread_map.items():
            norm_url = norm_thread_url(thr.get('url', ''))
            if norm_url not in urls_set:
                continue
            var = self._thread_checks.get(thr_iid)
            if var is None:
                continue
            if not var.get():
                continue
            var.set(False)
            changed = True
            try:
                vals = list(self.tree.item(thr_iid, "values"))
                vals[0] = "☐"
                self.tree.item(thr_iid, values=vals)
            except Exception:
                pass
            parent_iid = self.tree.parent(thr_iid)
            node = self._nmap.get(parent_iid)
            if node:
                node.thread_checks[thr.get('url', '')] = False
        if changed:
            self._propagate_thread_selection()
            self._save_state()
            self._log(f"⏭ Снято чекбоксов с тем (дата до фильтра): {len(urls_set)}", "warn")

    def _build_topbar(self):
        outer = tk.Frame(self, bg=C["panel"])
        outer.pack(fill="x")

        bar = tk.Frame(outer, bg=C["panel"], pady=6, padx=10)
        bar.pack(fill="x")

        tk.Label(bar, text="Форум:", bg=C["panel"], fg=C["sub"], font=("Segoe UI", 9)).pack(side="left")
        tk.Label(bar, text=FORUM_HOST, bg=C["panel"], fg=C["yellow"], font=("Segoe UI", 10, "bold")).pack(side="left", padx=(4, 14))

        self.v_auth = tk.StringVar(value="🔒 Не авторизован")
        self.lbl_auth = tk.Label(bar, textvariable=self.v_auth, bg=C["panel"], fg=C["red"],
                                 font=("Segoe UI", 9, "bold"))
        self.lbl_auth.pack(side="left", padx=(0, 14))

        self.v_transport = tk.StringVar(value="⚡ cookies")
        self.lbl_transport = tk.Label(bar, textvariable=self.v_transport, bg=C["panel"], fg=C["green"],
                                      font=("Segoe UI", 8, "bold"))
        self.lbl_transport.pack(side="left", padx=(0, 14))

        self.btn_login = self._mk_btn(bar, "🔑 Войти", self._do_login, C["orange"])
        self.btn_check_cookies = self._mk_btn(bar, "🍪 Проверить cookies", self._check_fast_session_click, C["green"])
        self.btn_scan = self._mk_btn(bar, "🔍 Сканировать", self._do_scan, C["accent"])
        self.btn_stop = self._mk_btn(bar, "⏹ Стоп", self._do_stop, C["red"])
        self.btn_load = self._mk_btn(bar, "💾 Загрузить", self._do_load, C["accent2"])
        self.btn_reset_profile = self._mk_btn(bar, "🧹 Сброс профиля", self._reset_browser_profile, C["red"])
        self.btn_reset_profile.pack(side="left", padx=3)
        self.btn_keywords = self._mk_btn(bar, "🔑 Ключевые слова", self._open_keyword_config, C["accent2"])
        self.btn_keywords.pack(side="left", padx=3)
        self.btn_topic_filter = self._mk_btn(bar, "🔍 Фильтр тем", self._open_topic_filter, C["accent2"])
        self.btn_topic_filter.pack(side="left", padx=3)

        for b in (self.btn_login, self.btn_check_cookies, self.btn_scan, self.btn_stop, self.btn_load):
            b.pack(side="left", padx=3)
        self.btn_stop.config(state="disabled")

        bar2 = tk.Frame(outer, bg=C["panel"], pady=8, padx=10)
        bar2.pack(fill="x")

        self.btn_mark_general = self._mk_btn(bar2, "📜 Общие правила",
                                             lambda: self._mark_selected("general"), C["accent2"])
        self.btn_mark_server = self._mk_btn(bar2, "🖥 Сервер",
                                            lambda: self._mark_selected("server"), C["accent"])
        self.btn_unmark = self._mk_btn(bar2, "🚫 Снять пометку",
                                       lambda: self._mark_selected(None), C["red"])
        self.btn_scan_deep = self._mk_btn(bar2, "🔍 Этап 2: Сканировать выбранные",
                                          self._do_scan_deep, C["orange"])
        self.btn_generate = self._mk_btn(bar2, "🚀 Сформировать документы",
                                         self._do_generate, C["green"])
        self.btn_history = self._mk_btn(bar2, "📜 История",
                                        self._show_history, C["accent2"])
        for b in (self.btn_mark_general, self.btn_mark_server,
                  self.btn_unmark, self.btn_scan_deep, self.btn_generate,
                  self.btn_history):
            b.pack(side="left", padx=3)

        self.v_marks_info = tk.StringVar(value="📜 Общие: 0   🖥 Сервер: 0")
        tk.Label(bar2, textvariable=self.v_marks_info, bg=C["panel"], fg=C["sub"], font=("Segoe UI", 9)).pack(side="right", padx=10)

    def _on_tree_double_click(self, event):
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        thr = self._thread_map.get(iid)
        if thr and thr.get("url"):
            try:
                webbrowser.open(thr["url"])
            except Exception:
                pass

    def _build_main(self):
        paned = tk.PanedWindow(self, orient="horizontal", bg=C["border"], sashwidth=4, sashrelief="flat")
        paned.pack(fill="both", expand=True, padx=6, pady=6)

        left = tk.Frame(paned, bg=C["panel"])
        paned.add(left, minsize=460)

        hdr = tk.Frame(left, bg=C["accent"], pady=5, padx=8)
        hdr.pack(fill="x")
        tk.Label(hdr, text="📋  Структура форума", bg=C["accent"], fg="white", font=("Segoe UI", 10, "bold")).pack(
            side="left")
        btn_row = tk.Frame(hdr, bg=C["accent"])
        btn_row.pack(side="right")
        self._mk_btn(btn_row, "✅ Все", self._check_all, C["green"], small=True).pack(side="left", padx=2)
        self._mk_btn(btn_row, "☐ Сброс", self._uncheck_all, C["red"], small=True).pack(side="left", padx=2)
        self._mk_btn(btn_row, "➕ Раскрыть", self._expand_all, C["panel2"], small=True).pack(side="left", padx=2)
        self._mk_btn(btn_row, "➖ Свернуть", self._collapse_all, C["panel2"], small=True).pack(side="left", padx=2)

        leg = tk.Frame(left, bg=C["panel2"], pady=3)
        leg.pack(fill="x")
        labels = ["Корень", "Раздел", "Подраздел", "Ур.3", "Ур.4", "Ур.5+"]
        for icon, color, lbl in zip(ICONS, LEVEL_COLORS, labels):
            tk.Label(leg, text=f"{icon} {lbl}", fg=color, bg=C["panel2"], font=("Segoe UI", 7)).pack(side="left",
                                                                                                     padx=5)

        self.v_count = tk.StringVar(value="Разделов: 0")
        tk.Label(left, textvariable=self.v_count, bg=C["panel"], fg=C["sub"], font=("Segoe UI", 8)).pack(anchor="w",
                                                                                                         padx=8, pady=2)

        wrap = tk.Frame(left, bg=C["panel"])
        wrap.pack(fill="both", expand=True, padx=4, pady=(0, 4))
        style = ttk.Style(self)
        style.theme_use("clam")
        style.configure("FM.Treeview", background=C["panel"], foreground=C["text"],
                        fieldbackground=C["panel"], rowheight=24, font=("Segoe UI", 9), borderwidth=0)
        style.configure("FM.Treeview.Heading", background=C["panel2"], foreground=C["sub"],
                        font=("Segoe UI", 8, "bold"), relief="flat")
        style.map("FM.Treeview", background=[("selected", C["sel"])], foreground=[("selected", "white")])

        self.tree = ttk.Treeview(wrap, style="FM.Treeview", selectmode="browse", show="tree headings")
        self.tree["columns"] = ("check", "mark", "threads", "kids")
        self.tree.column("#0", width=440, minwidth=200, stretch=True)
        self.tree.column("check", width=60, minwidth=50, stretch=False, anchor="center")
        self.tree.column("mark", width=100, minwidth=80, stretch=False, anchor="center")
        self.tree.column("threads", width=60, minwidth=50, stretch=False, anchor="center")
        self.tree.column("kids", width=70, minwidth=50, stretch=False, anchor="center")
        self.tree.heading("#0", text="Раздел / Подраздел")
        self.tree.heading("check", text="Выбор")
        self.tree.heading("mark", text="Пометка")
        self.tree.heading("threads", text="Тем")
        self.tree.heading("kids", text="Дочерних")
        vsb = ttk.Scrollbar(wrap, orient="vertical", command=self.tree.yview)
        hsb = ttk.Scrollbar(wrap, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        sframe = tk.Frame(wrap, bg=C["panel"])
        sframe.grid(row=2, column=0, columnspan=2, sticky="ew", pady=(4, 0))
        tk.Label(sframe, text="🔍", bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 9)).pack(side="left", padx=(2, 0))
        self.v_tsearch = tk.StringVar()
        self.v_tsearch.trace_add("write", lambda *_: self._apply_struct_filter())
        tk.Entry(sframe, textvariable=self.v_tsearch, bg=C["entry"], fg=C["text"],
                 insertbackground=C["text"], relief="flat", font=("Segoe UI", 9),
                 highlightthickness=1, highlightbackground=C["border"],
                 highlightcolor=C["accent"]).pack(side="left", fill="x", expand=True, padx=4, ipady=2)
        self.v_tcount = tk.StringVar(value="")
        tk.Label(sframe, textvariable=self.v_tcount, bg=C["panel"], fg=C["sub"],
                 font=("Segoe UI", 8)).pack(side="left", padx=4)
        self._mk_btn(sframe, "🚫 Исключения…", self._open_struct_exclude,
                     C["accent2"], small=True).pack(side="left", padx=4)
        self._struct_hidden = []
        wrap.columnconfigure(0, weight=1)
        wrap.rowconfigure(0, weight=1)
        for i, color in enumerate(LEVEL_COLORS):
            self.tree.tag_configure(f"lv{i}", foreground=color)
        self.tree.tag_configure("checked", foreground=C["green"])
        self.tree.tag_configure("unchecked", foreground=C["text"])
        self.tree.tag_configure("thread", foreground=C["sub"])
        self.tree.bind("<Button-1>", self._on_tree_click)
        self.tree.bind("<<TreeviewSelect>>", self._on_tree_select)
        self.tree.bind("<Double-1>", self._on_tree_double_click)

        right = tk.Frame(paned, bg=C["panel"])
        paned.add(right, minsize=380)
        right.rowconfigure(2, weight=2)
        right.rowconfigure(3, weight=1)
        right.columnconfigure(0, weight=1)

        info_f = tk.LabelFrame(right, text="  ℹ  Выбранный раздел  ", bg=C["panel"], fg=C["accent"],
                               font=("Segoe UI", 9, "bold"), bd=1, relief="groove")
        info_f.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 3))
        self.v_sel_title = tk.StringVar(value="— не выбрано —")
        tk.Label(info_f, textvariable=self.v_sel_title, bg=C["panel"], fg=C["yellow"],
                 font=("Segoe UI", 9, "bold"), wraplength=340, justify="left").pack(anchor="w", padx=8, pady=(4, 2))

        cfg_f = tk.LabelFrame(right, text="  ⚙  Глобальные настройки  ", bg=C["panel"], fg=C["accent"],
                              font=("Segoe UI", 9, "bold"), bd=1, relief="groove")
        cfg_f.grid(row=1, column=0, sticky="ew", padx=6, pady=(0, 3))
        cfg_f.columnconfigure(1, weight=1)
        tk.Label(cfg_f, text="Промт:", bg=C["panel"], fg=C["sub"], font=("Segoe UI", 8)).grid(row=0, column=0,
                                                                                              sticky="w", padx=8,
                                                                                              pady=4)
        self.v_prompt_path = tk.StringVar(value=self._prompt_path or "— не выбран —")
        tk.Label(cfg_f, textvariable=self.v_prompt_path, bg=C["panel"], fg=C["text"],
                 font=("Segoe UI", 8), wraplength=200, justify="left").grid(row=0, column=1, sticky="w", padx=4)
        self._mk_btn(cfg_f, "📄 Выбрать", self._choose_prompt, C["accent2"], small=True).grid(row=0, column=2, padx=6)
        self.v_fpo = tk.BooleanVar(value=DEFAULT_FIRST_POST_ONLY)
        tk.Checkbutton(cfg_f, text="Только первый пост темы", variable=self.v_fpo,
                       bg=C["panel"], fg=C["text"], activebackground=C["panel"],
                       selectcolor=C["entry"], font=("Segoe UI", 8)).grid(row=1, column=0, columnspan=3, sticky="w",
                                                                          padx=8, pady=2)
        tk.Checkbutton(cfg_f, text="Показывать DEBUG-лог фильтрации ссылок", variable=self._show_debug,
                       bg=C["panel"], fg=C["sub"], activebackground=C["panel"],
                       selectcolor=C["entry"], font=("Segoe UI", 8)).grid(row=2, column=0, columnspan=3, sticky="w",
                                                                          padx=8, pady=2)
        tk.Label(cfg_f, text="Потоков:", bg=C["panel"], fg=C["sub"], font=("Segoe UI", 8)).grid(row=3, column=0,
                                                                                                sticky="w", padx=8,
                                                                                                pady=4)
        self.v_workers = tk.IntVar(value=getattr(self, "_workers_n", MAX_WORKERS))
        tk.Spinbox(cfg_f, from_=1, to=64, textvariable=self.v_workers, width=5,
                   bg=C["entry"], fg=C["text"], relief="flat").grid(row=3, column=1, sticky="w")
        self._mk_btn(cfg_f, "Применить", self._apply_workers, C["accent2"], small=True).grid(row=3, column=2, padx=6)
        tk.Checkbutton(cfg_f, text="⚡ Быстрый режим (cookies, браузер — запасной)", variable=self.v_fast,
                       bg=C["panel"], fg=C["text"], activebackground=C["panel"],
                       selectcolor=C["entry"], font=("Segoe UI", 8)).grid(row=4, column=0, columnspan=3, sticky="w",
                                                                          padx=8, pady=2)
        tk.Checkbutton(cfg_f, text="🐢 Тихий режим (паузы ×3, после бана)", variable=self.v_stealth,
                       bg=C["panel"], fg=C["text"], activebackground=C["panel"],
                       selectcolor=C["entry"], font=("Segoe UI", 8)).grid(row=5, column=0, columnspan=3, sticky="w",
                                                                          padx=8, pady=2)

        disp_f = tk.LabelFrame(right, text="  📊  Диспетчер процессов  ", bg=C["panel"], fg=C["green"],
                               font=("Segoe UI", 9, "bold"), bd=1, relief="groove")
        disp_f.grid(row=2, column=0, sticky="nsew", padx=6, pady=(0, 3))
        disp_f.rowconfigure(0, weight=1)
        disp_f.columnconfigure(0, weight=1)
        self.dispatcher = ProcessDispatcher(disp_f)
        self.dispatcher.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

        log_f = tk.LabelFrame(right, text="  📜  Лог  ", bg=C["panel"], fg=C["accent2"],
                              font=("Segoe UI", 9, "bold"), bd=1, relief="groove")
        log_f.grid(row=3, column=0, sticky="nsew", padx=6, pady=(0, 6))
        log_f.rowconfigure(0, weight=1)
        log_f.columnconfigure(0, weight=1)
        self.rlog = RichLog(log_f)
        self.rlog.grid(row=0, column=0, sticky="nsew", padx=2, pady=2)

    def _build_statusbar(self):
        bar = tk.Frame(self, bg=C["panel2"], pady=3)
        bar.pack(fill="x", side="bottom")
        self.v_status = tk.StringVar(value="Готов. Сначала нажмите «🔑 Войти», затем «🔍 Сканировать».")
        tk.Label(bar, textvariable=self.v_status, bg=C["panel2"], fg=C["sub"], font=("Segoe UI", 8), anchor="w").pack(
            side="left", padx=10)

    def _log(self, msg: str, tag="normal"):
        self.rlog.log(msg, tag)

    def _debug(self, msg: str):
        if self._show_debug.get():
            self.rlog.log(msg, "debug")
        get_logger().debug(msg)

    def _status(self, msg: str):
        self.after(0, lambda: self.v_status.set(msg))

    def _set_busy(self, on: bool):
        self._running = on

        def _do():
            extra = []
            if getattr(self, "btn_participation", None):
                extra.append(self.btn_participation)
            buttons = (
                          self.btn_scan, self.btn_login, self.btn_load,
                          self.btn_mark_general, self.btn_mark_server,
                          self.btn_unmark, self.btn_scan_deep, self.btn_generate,
                          self.btn_keywords, self.btn_topic_filter
                      ) + tuple(extra)
            if on:
                for b in buttons:
                    b.config(state="disabled")
                self.btn_stop.config(state="normal")
            else:
                for b in buttons:
                    b.config(state="normal")
                self.btn_stop.config(state="disabled")

        self.after(0, _do)

    def _start_emitter(self, label, total=0):
        self._emitter = ProcessEmitter()
        self.dispatcher.clear_all()
        self.dispatcher.attach(self._emitter)
        return self._emitter.process_started(label, total)

    def _on_fast_toggle(self):
        try:
            val = bool(self.v_fast.get())
        except Exception:
            return
        self.http.fast_enabled = val
        try:
            self.project.save_config({"fast_path_enabled": val})
        except Exception:
            pass
        self._update_transport_indicator()
        self._log(f"⚡ Быстрый режим {'ВКЛ (cookies + браузер-запасной)' if val else 'ВЫКЛ (только браузер)'}",
                  "info")

    def _on_stealth_toggle(self):
        try:
            val = bool(self.v_stealth.get())
        except Exception:
            return
        self._apply_stealth(val)
        try:
            self.project.save_config({"stealth_mode": val})
        except Exception:
            pass

    def _apply_stealth(self, val: bool, silent: bool = False):
        """v1.21: 🐢 Тихий режим — все паузы ×3 (и браузерные, и быстрые)."""
        mult = 3.0 if val else 1.0
        self.http.stealth_mult = mult
        try:
            self.http._rate.pace_mult = mult
        except Exception:
            pass
        if not silent:
            self._log(f"🐢 Тихий режим {'ВКЛ: паузы ×3 — медленно, но незаметно' if val else 'ВЫКЛ: обычный темп'}",
                      "info")

    def _update_transport_indicator(self):
        try:
            st = self.http.get_stats()
        except Exception:
            return
        if not st.get("enabled"):
            self.v_transport.set("🌐 только браузер")
            self.lbl_transport.config(fg=C["yellow"])
            return
        bleft = st.get("breaker_left", 0)
        if bleft > 0:
            mm, ss = divmod(int(bleft), 60)
            self.v_transport.set(f"⏸ форум отдыхает {mm}:{ss:02d} · ⚡{st['fast']} 🌐{st['browser']}")
            self.lbl_transport.config(fg=C["yellow"])
            return
        left = st.get("browser_only_left", 0)
        if left > 0:
            mm, ss = divmod(int(left), 60)
            self.v_transport.set(f"🌐 браузер-only {mm}:{ss:02d} · ⚡{st['fast']} 🌐{st['browser']}")
            self.lbl_transport.config(fg=C["yellow"])
        else:
            self.v_transport.set(f"⚡cookies {st['fast']} · 🌐{st['browser']} · ⛔{st['fast_blocks']}")
            self.lbl_transport.config(fg=C["green"])

    def _poll_transport_stats(self):
        try:
            self._update_transport_indicator()
        except Exception:
            pass
        try:
            self.after(2000, self._poll_transport_stats)
        except Exception:
            pass

    def _persist_harvested_cookies(self, cookies):
        """v1.19: свежие cookies из браузера сразу пишем в файл (не только в память)."""
        try:
            if cookies:
                self.project.save_cookies(list(cookies))
                get_logger().debug(f"harvest: {len(cookies)} cookies сохранены в файл")
        except Exception:
            pass

    def _harvest_browser_cookies(self):
        if not self._fetch_browser_lock.acquire(blocking=False):
            return None
        try:
            browser = self._fetch_browser
            if not browser or not browser.driver:
                return None
            try:
                return browser.export_cookies()
            except Exception:
                return None
        finally:
            self._fetch_browser_lock.release()

    def _update_login_indicator(self):
        has_cookies = self.project.cookies.exists()
        with self._fetch_browser_lock:
            has_fallback = bool(self._fetch_browser and self._fetch_browser.driver)
        suffix = "  🌐(браузер-фолбэк активен)" if has_fallback else ""
        if self._logged_in:
            self.v_auth.set(f"🔓 Авторизован{suffix}")
            self.lbl_auth.config(fg=C["green"])
        elif has_cookies:
            self.v_auth.set(f"🔑 Cookies сохранены{suffix}")
            self.lbl_auth.config(fg=C["yellow"])
        else:
            self.v_auth.set("🔒 Не авторизован")
            self.lbl_auth.config(fg=C["red"])

    def _open_struct_exclude(self):
        try:
            win = StructureExcludeWindow(self, self)
        except Exception:
            return
        try:
            self.wait_window(win)
        except Exception:
            pass
        try:
            self._apply_struct_filter()
        except Exception:
            pass

    def _apply_struct_filter(self):
        """Поиск по темам/разделам в дереве + исключения структуры (v1.25)."""
        try:
            q = (self.v_tsearch.get() or "").strip().lower()
        except Exception:
            return
        try:
            excl = list((self.project.load_config() or {}).get("structure_exclude_words", []) or [])
        except Exception:
            excl = []
        # Возвращаем спрятанное: по родителям, в порядке исходных индексов.
        try:
            items = list(getattr(self, "_struct_hidden", []) or [])
        except Exception:
            items = []
        try:
            items.sort(key=lambda e: (str(e[1]), e[2]))
        except Exception:
            pass
        for iid, par, idx in items:
            try:
                self.tree.move(iid, par, idx)
            except Exception:
                pass
        self._struct_hidden = []
        try:
            self.v_tcount.set("")
        except Exception:
            pass
        if not q and not excl:
            return
        hidden = []
        hidden_excl = 0

        def hide(iid):
            try:
                par = self.tree.parent(iid)
                idx = self.tree.index(iid)
            except Exception:
                return
            hidden.append((iid, par, idx))
            try:
                self.tree.detach(iid)
            except Exception:
                pass

        try:
            thr_items = list((self._thread_map or {}).items())
        except Exception:
            thr_items = []
        shown = 0
        for thr_iid, thr in thr_items:
            try:
                title = thr.get("title", "") or ""
            except Exception:
                title = ""
            try:
                _pn = (self._nmap or {}).get(self.tree.parent(thr_iid))
                _sec = _pn.title if _pn else ""
            except Exception:
                _sec = ""
            if struct_row_hidden(title, _sec, q, excl):
                hide(thr_iid)
                try:
                    if excl and is_excluded_by_words(title, _sec, excl):
                        hidden_excl += 1
                except Exception:
                    pass
            else:
                shown += 1
        try:
            node_iids = list((self._nmap or {}).keys())
        except Exception:
            node_iids = []
        try:
            node_iids.sort(key=lambda i: (self._depths or {}).get(i, 0), reverse=True)
        except Exception:
            pass
        for iid in node_iids:
            try:
                node = (self._nmap or {}).get(iid)
                _nt = ((node.title if node else "") or "").lower()
                own_q = bool(q) and q in _nt
                kids = self.tree.get_children(iid)
            except Exception:
                continue
            if own_q or kids:
                continue
            if not q:
                # Только исключения: непросканированные разделы не трогаем.
                try:
                    _has = bool(node.threads)
                except Exception:
                    _has = False
                if not _has:
                    continue
            hide(iid)
        self._struct_hidden = hidden
        try:
            _lbl = f"Найдено тем: {shown}" if q else f"Тем: {shown}"
            if hidden_excl:
                _lbl += f" · скрыто: {hidden_excl}"
            self.v_tcount.set(_lbl if (q or hidden_excl) else "")
        except Exception:
            pass

    def _build_tree(self, nodes):
        self._struct_hidden = []
        self.tree.delete(*self.tree.get_children())
        self._checks.clear()
        self._nmap.clear()
        self._depths.clear()
        self._thread_map.clear()
        self._thread_checks.clear()
        total = [0]

        def insert(parent_iid: str, node: Node, depth: int):
            total[0] += 1

            raw_iid = node.url
            iid = raw_iid
            suffix = 0
            while True:
                try:
                    self.tree.item(iid)
                    suffix += 1
                    iid = f"{raw_iid}::{suffix}"
                except tk.TclError:
                    break

            self._nmap[iid] = node
            self._nmap[node.url] = node
            self._checks[iid] = tk.BooleanVar(value=False)
            self._depths[iid] = depth
            node._tree_iid = iid

            lv = min(depth, len(LEVEL_COLORS) - 1)
            icon = ICONS[min(depth, len(ICONS) - 1)]
            kids = len(node.children)

            self.tree.insert(
                parent_iid, "end",
                iid=iid,
                text=f"{icon}  {node.title}",
                values=(
                    "☐",
                    MARK_LABELS.get(node.mark, ""),
                    str(node.thread_count) if node.thread_count else "",
                    str(kids) if kids else "",
                ),
                open=(depth < 2),
                tags=(f"lv{lv}", "unchecked"),
            )

            # ── Темы ──────────────────────────────────────────────────
            for idx, t in enumerate(node.threads):
                thr_iid = f"THR::{iid}::{idx}"
                self._thread_map[thr_iid] = t
                sticky = "📌 " if t.get("sticky") else ""
                thr_url = t.get("url", "")

                # Состояние галочки темы берём из node.thread_checks
                thr_checked = node.thread_checks.get(thr_url, False)
                self._thread_checks[thr_iid] = tk.BooleanVar(value=thr_checked)
                thr_sym = "✅" if thr_checked else "☐"

                self.tree.insert(
                    iid, "end",
                    iid=thr_iid,
                    text=f"📄  {sticky}{t.get('title', '')}",
                    values=(thr_sym, "", "", ""),
                    tags=("thread",),
                )

            # ── Рекурсия ──────────────────────────────────────────────
            for child in node.children:
                insert(iid, child, depth + 1)

        for n in nodes:
            insert("", n, 0)

        # Автоматически проставляем "Выбор" узлам, у которых есть
        # выбранные темы (по умолчанию все темы считаются выбранными)
        self._propagate_thread_selection()

        self.after(0, lambda: self.v_count.set(f"Разделов: {total[0]}"))
        self.after(0, self._refresh_marks_info)
        self.after(0, self._apply_struct_filter)

    def _refresh_marks_info(self):
        g = len(find_marked_nodes(self.nodes, "general")) if self.nodes else 0
        s = len(find_marked_nodes(self.nodes, "server")) if self.nodes else 0
        self.v_marks_info.set(f"📜 Общие: {g}   🖥 Сервер: {s}")

    def _set_check(self, iid: str, val: bool, recursive=True):
        if iid not in self._checks:
            return
        self._checks[iid].set(val)
        sym = "✅" if val else "☐"
        tag2 = "checked" if val else "unchecked"
        lv = min(self._depths.get(iid, 0), len(LEVEL_COLORS) - 1)
        try:
            vals = list(self.tree.item(iid, "values"))
            vals[0] = sym
            self.tree.item(iid, values=vals, tags=(f"lv{lv}", tag2))
        except Exception:
            pass

        # Синхронизируем галочки тем этого узла
        node = self._nmap.get(iid)
        for child_iid in self.tree.get_children(iid):
            if child_iid.startswith("THR::"):
                thr = self._thread_map.get(child_iid)
                # v1.23: исключения бьют и отметку раздела — такие темы не ставим
                # (иначе клик по разделу / «Все» / восстановление состояния
                # возвращали бы галочки, снятые на Этапе 2).
                eff = val
                if eff and thr is not None and node is not None:
                    try:
                        if is_excluded_by_words(thr.get("title", ""), node.title,
                                               _autocheck_thread_exclude):
                            eff = False
                    except Exception:
                        pass
                # Обновляем BooleanVar и отображение темы
                thr_var = self._thread_checks.get(child_iid)
                if thr_var is not None:
                    thr_var.set(eff)
                    thr_sym = "✅" if eff else "☐"
                    try:
                        thr_vals = list(self.tree.item(child_iid, "values"))
                        thr_vals[0] = thr_sym
                        self.tree.item(child_iid, values=thr_vals)
                    except Exception:
                        pass
                # Обновляем node.thread_checks
                if node is not None and thr:
                    node.thread_checks[thr.get("url", "")] = eff
            elif recursive:
                # Рекурсия в дочерние подразделы
                self._set_check(child_iid, val, recursive=True)

    def _set_check_visual_only(self, iid: str, val: bool):
        """
        Обновляет чекбокс узла БЕЗ рекурсивного изменения его тем/детей.
        Используется когда галочка подраздела должна проставиться
        автоматически из-за выбора вложенной темы.
        """
        if iid not in self._checks:
            return
        self._checks[iid].set(val)
        sym = "✅" if val else "☐"
        tag2 = "checked" if val else "unchecked"
        lv = min(self._depths.get(iid, 0), len(LEVEL_COLORS) - 1)
        try:
            vals = list(self.tree.item(iid, "values"))
            vals[0] = sym
            self.tree.item(iid, values=vals, tags=(f"lv{lv}", tag2))
        except Exception:
            pass

    def _is_node_checked(self, node) -> bool:
        """Проверяет, установлена ли галочка у данного узла в дереве."""
        iid = getattr(node, "_tree_iid", None)
        if iid is None:
            return False
        var = self._checks.get(iid)
        if var is None:
            return False
        return bool(var.get())

    def _propagate_thread_selection(self):
        """
        Автоматически ставит "Выбор" любому узлу (и всем его родителям
        вплоть до корня), если внутри него самого или в его потомках
        есть хотя бы одна выбранная тема. Уже стоящие галочки не снимает.
        Тема считается выбранной если явно не снята (по умолчанию True).
        """

        def walk(node) -> bool:
            has_selected = False
            for t in node.threads:
                if node.thread_checks.get(t.get("url", ""), False):
                    has_selected = True
            for c in node.children:
                if walk(c):
                    has_selected = True
            if has_selected:
                iid = getattr(node, "_tree_iid", None)
                if iid and iid in self._checks and not self._checks[iid].get():
                    self._set_check_visual_only(iid, True)
            return has_selected

        for n in self.nodes or []:
            walk(n)

    def _is_thread_checked(self, node: "Node", thread: dict) -> bool:
        """
        Проверяет выбор отдельной темы.
        По умолчанию тема считается НЕ выбранной, пока пользователь
        явно не поставит галочку (запись True в thread_checks).
        """
        thr_url = thread.get('url', '')
        return node.thread_checks.get(thr_url, False)

    def _set_mark(self, iid: str, value):
        node = self._nmap.get(iid)
        if not node:
            return
        node.mark = value
        try:
            vals = list(self.tree.item(iid, "values"))
            vals[1] = MARK_LABELS.get(value, "")
            self.tree.item(iid, values=vals)
        except Exception:
            pass

    def _get_ancestor_iids(self, iid):
        """Список iid всех родителей узла (от ближайшего к корню)."""
        ids = []
        cur = self.tree.parent(iid)
        while cur:
            ids.append(cur)
            cur = self.tree.parent(cur)
        return ids

    def _get_descendant_iids(self, iid):
        """Список iid всех потомков узла (без тем)."""
        ids = []

        def walk(i):
            for ch in self.tree.get_children(i):
                if ch.startswith("THR::"):
                    continue
                ids.append(ch)
                walk(ch)

        walk(iid)
        return ids

    def _selected_section_iid(self):
        sel = self.tree.selection()
        if not sel:
            return None

        iid = sel[0]

        if iid.startswith("THR::"):
            parent_iid = self.tree.parent(iid)
            if not parent_iid:
                return None

            self.tree.selection_set(parent_iid)
            iid = parent_iid

        return iid

    def _mark_selected(self, value):
        iid = self._selected_section_iid()

        if not iid:
            self._log("⚠ Сначала выберите раздел в дереве", "warn")
            self._status("⚠ Выберите раздел в дереве")
            return

        node = self._nmap.get(iid)
        if not node:
            self._log("⚠ Выбранный элемент не является разделом", "warn")
            return

        if value == "server":
            depth = self._depths.get(iid, -1)
            if depth != 0 and node.mark != "server":
                self._log(
                    f"⚠ Пропущено «{node.title}»: пометка «Сервер» доступна "
                    f"только для узлов уровня «Корень»", "warn"
                )
                return

        old_mark = node.mark

        if value is None:
            new_mark = None
        elif old_mark == value:
            new_mark = None
        else:
            new_mark = value

        if old_mark == new_mark:
            self._log(f"ℹ Пометка «{node.title}» уже пустая", "info")
            return

        self._set_mark(iid, new_mark)

        if new_mark is None:
            self._log(f"🏷 Пометка снята с «{node.title}»", "info")
        else:
            label = {"general": "Общие правила", "server": "Сервер"}[new_mark]
            self._log(f"🏷 «{node.title}» помечен как «{label}»", "info")

        self._refresh_marks_info()
        self._on_tree_select(None)

        if self.nodes:
            self.project.save_struct(self.nodes)

        self._save_state()

    def _auto_load_state(self):
        """Автоматически загружает последнее сохранённое состояние при запуске."""
        nodes = self.project.load_struct()
        if not nodes:
            self._log("ℹ Нет сохранённой структуры. Выполните сканирование.", "info")
            return
        self.nodes = nodes
        self._build_tree(nodes)
        # Восстанавливаем галочки и пометки
        self._restore_checks()
        cnt = len(flatten_nodes(nodes))
        self._status(f"✅ Автозагрузка: {cnt} разделов")
        self._log(f"💾 Автозагружено {cnt} разделов из сохранённого файла", "ok")

    def _restore_checks(self):
        """Восстанавливает галочки, пометки и выбор тем из файла checks.json."""
        saved = self.project.load_checks()
        if not saved:
            return
        checks = saved.get('checks', {})
        marks = saved.get('marks', {})
        t_chks = saved.get('thread_checks', {})

        # Восстанавливаем галочки разделов
        for iid, var in self._checks.items():
            node = self._nmap.get(iid)
            if node is None:
                continue
            # Ищем по URL узла
            url = node.url
            if url in checks:
                val = checks[url]
                self._set_check(iid, val, recursive=False)

        # Восстанавливаем пометки
        for iid in list(self._nmap.keys()):
            node = self._nmap.get(iid)
            if node is None:
                continue
            url = node.url
            if url in marks:
                node.mark = marks[url]
                try:
                    vals = list(self.tree.item(iid, "values"))
                    vals[1] = MARK_LABELS.get(node.mark, "")
                    self.tree.item(iid, values=vals)
                except Exception:
                    pass

        # Восстанавливаем галочки тем
        for thr_iid, var in self._thread_checks.items():
            thr = self._thread_map.get(thr_iid)
            if thr is None:
                continue
            thr_url = thr.get('url', '')
            if thr_url in t_chks:
                val = t_chks[thr_url]
                # v1.24: рестор тоже уважает исключения + синхрон модели
                # (раньше ставил визуал мимо списка и рассинхронизировал модель).
                try:
                    _pn = self._nmap.get(self.tree.parent(thr_iid))
                    if val and is_excluded_by_words(thr.get('title', ''),
                                                   _pn.title if _pn else '',
                                                   _autocheck_thread_exclude):
                        val = False
                except Exception:
                    pass
                var.set(val)
                sym = "✅" if val else "☐"
                try:
                    vals = list(self.tree.item(thr_iid, "values"))
                    vals[0] = sym
                    self.tree.item(thr_iid, values=vals)
                except Exception:
                    pass
                try:
                    _pn2 = self._nmap.get(self.tree.parent(thr_iid))
                    if _pn2 is not None:
                        _pn2.thread_checks[thr_url] = val
                except Exception:
                    pass

        self._propagate_thread_selection()
        self._refresh_marks_info()
        self._log("✅ Состояние восстановлено из сохранённого файла", "ok")

    def _apply_pre_scan_selection(self, before_snapshot: dict):
        """
        Восстанавливает галочки строго по снимку состояния ДО Этапа 2.
        Всё, чего не было в снимке:
          - новые разделы/подразделы;
          - новые темы;
        остаётся/становится снятым.

        Исключения:
          - если новый раздел/подраздел подходит под список
            AUTO_CHECK_NEW_TITLE_KEYWORDS, он автоматически отмечается;
          - если тема подходит под список
            AUTO_CHECK_NEW_THREAD_TITLE_KEYWORDS, она автоматически отмечается.
        """
        checks = before_snapshot.get("checks", {}) or {}
        marks = before_snapshot.get("marks", {}) or {}
        t_chks = before_snapshot.get("thread_checks", {}) or {}

        def node_key(u: str) -> str:
            u = (u or "").strip()
            if not u:
                return ""

            # Якорные категории вида "/#category.97" нельзя нормализовать
            # через norm_url(), иначе срежется часть после '#'.
            if u.startswith("#") or u.startswith("/#"):
                return u.lower().rstrip("/")

            return norm_url(u)

        checked_nodes = {
            node_key(u)
            for u, v in checks.items()
            if v and u
        }

        # Все узлы, которые уже существовали ДО текущего сканирования.
        # Если URL узла отсутствует в known_nodes — узел найден впервые.
        known_nodes = {
            node_key(u)
            for u in checks.keys()
            if u
        }

        marks_nodes = {
            node_key(u): v
            for u, v in marks.items()
            if u
        }

        # Темы, которые уже существовали ДО текущего сканирования.
        known_threads = {
            norm_thread_url(u)
            for u in t_chks.keys()
            if u
        }

        # Темы, которые были отмечены ДО текущего сканирования.
        checked_threads = {
            norm_thread_url(u)
            for u, v in t_chks.items()
            if v and u
        }

        # ── Разделы/подразделы ─────────────────────────────────────
        autochecked_new_nodes = []

        for iid, var in self._checks.items():
            node = self._nmap.get(iid)
            if node is None:
                continue

            key = node_key(node.url)

            # Якорные категории /#... обычно не являются реальными
            # сканируемыми разделами, поэтому их автоотметку пропускаем.
            is_anchor = (
                    (node.url or "").startswith("#")
                    or (node.url or "").startswith("/#")
            )

            if key in checked_nodes:
                new_checked = True
            elif not is_anchor and should_autocheck_new_node(node.title):
                # Проверяем по ключевым словам ВСЕГДА, даже если раздел уже известен
                new_checked = True
                autochecked_new_nodes.append(node.title)
            else:
                new_checked = False

            self._set_check_visual_only(iid, new_checked)

            # Заодно восстанавливаем пометку, если она была в снимке.
            if key in marks_nodes:
                new_mark = marks_nodes[key]
                if node.mark != new_mark:
                    node.mark = new_mark
                    try:
                        vals = list(self.tree.item(iid, "values"))
                        vals[1] = MARK_LABELS.get(new_mark, "")
                        self.tree.item(iid, values=vals)
                    except Exception:
                        pass

        # ── Темы ───────────────────────────────────────────────────
        autochecked_threads = []
        blocked_by_exclude = 0
        thread_dates = load_thread_dates_cached(self.project.dir)
        for thr_iid, var in self._thread_checks.items():
            thr = self._thread_map.get(thr_iid)
            if thr is None:
                continue

            thr_url = thr.get("url", "")
            thr_title = thr.get("title", "")

            if not thr_url:
                var.set(False)
                try:
                    vals = list(self.tree.item(thr_iid, "values"))
                    vals[0] = "☐"
                    self.tree.item(thr_iid, values=vals)
                except Exception:
                    pass
                continue

            key = norm_thread_url(thr_url)
            is_new_thread = key not in known_threads

            cached_date = thread_dates.get(key, '')
            kw_eval = evaluate_thread_keyword_match(thr_title, cached_date, _autocheck_thread_keywords)
            node_match = should_autocheck_new_node(thr_title)
            # v1.24: исключения бьют и финальное восстановление (иначе оно
            # воскрешало галочки, снятые сканером, и ставило новые мимо списка).
            try:
                _pn = self._nmap.get(self.tree.parent(thr_iid))
                _pt = _pn.title if _pn else ""
                _excluded = is_excluded_by_words(thr_title, _pt, _autocheck_thread_exclude)
            except Exception:
                _excluded = False
            _was = key in checked_threads
            _kwm = (kw_eval == 'check' or node_match or should_autocheck_new_thread(thr_title))
            if _excluded and kw_eval != 'uncheck' and (_was or _kwm):
                blocked_by_exclude += 1
            val = final_thread_check_value(excluded=_excluded, kw_eval=kw_eval,
                                           was_checked=_was, kw_match=_kwm)
            if val and _kwm and not _was:
                if thr_title and thr_title not in autochecked_threads:
                    autochecked_threads.append(thr_title)

            var.set(val)
            sym = "✅" if val else "☐"

            try:
                vals = list(self.tree.item(thr_iid, "values"))
                vals[0] = sym
                self.tree.item(thr_iid, values=vals)
            except Exception:
                pass

            parent_iid = self.tree.parent(thr_iid)
            parent_node = self._nmap.get(parent_iid)

            if parent_node is not None:
                parent_node.thread_checks[thr_url] = val

        if autochecked_new_nodes:
            self._log(
                f"🆕✅ Автоматически отмечено новых разделов по списку исключений: "
                f"{len(autochecked_new_nodes)}",
                "info"
            )
            for title in autochecked_new_nodes[:20]:
                self._log(f"    + {title}", "info")

            if len(autochecked_new_nodes) > 20:
                self._log(
                    f"    … и ещё {len(autochecked_new_nodes) - 20}",
                    "info"
                )

        if autochecked_threads:
            self._log(
                f"✅ Автоматически отмечено тем по списку исключений: "
                f"{len(autochecked_threads)}",
                "info"
            )
            for title in autochecked_threads[:20]:
                self._log(f"    + {title}", "info")

            if len(autochecked_threads) > 20:
                self._log(
                    f"    … и ещё {len(autochecked_threads) - 20}",
                    "info"
                )

        if blocked_by_exclude:
            self._log(f"⏭ Снято отметок (исключения, финал): {blocked_by_exclude}", "info")

        # Чтобы родительские разделы визуально отметились, если внутри есть
        # выбранные темы.
        self._propagate_thread_selection()
        self._refresh_marks_info()

    def _snapshot_state(self) -> dict:
        """Снимок текущего состояния 'Выбор' / пометок / выбора тем по URL."""
        checks, marks, t_chks = {}, {}, {}
        for iid, var in self._checks.items():
            node = self._nmap.get(iid)
            if node is None:
                continue
            checks[node.url] = bool(var.get())
            marks[node.url] = node.mark
        for thr_iid, var in self._thread_checks.items():
            thr = self._thread_map.get(thr_iid)
            if thr:
                t_chks[thr.get("url", "")] = bool(var.get())
        return {"checks": checks, "marks": marks, "thread_checks": t_chks}

    def _title_by_url(self, url: str) -> str:
        node = self._nmap.get(url)
        return node.title if node else url

    def _diff_state(self, before: dict, after: dict) -> list:
        lines = []
        b_checks, a_checks = before.get("checks", {}), after.get("checks", {})
        b_marks, a_marks = before.get("marks", {}), after.get("marks", {})
        b_t, a_t = before.get("thread_checks", {}), after.get("thread_checks", {})

        for url, val in a_checks.items():
            if url in b_checks and b_checks[url] != val:
                lines.append(f"Выбор «{self._title_by_url(url)}»: {b_checks[url]} → {val}")

        for url, val in a_marks.items():
            if url in b_marks and b_marks[url] != val:
                bl = MARK_LABELS.get(b_marks[url], "нет")
                al = MARK_LABELS.get(val, "нет")
                lines.append(f"Пометка «{self._title_by_url(url)}»: {bl} → {al}")

        for url, val in a_t.items():
            if url in b_t and b_t[url] != val:
                lines.append(f"Выбор темы изменён: {url}: {b_t[url]} → {val}")

        return lines



    def _save_state(self):
        """Сохраняет текущее состояние галочек, пометок и выбора тем."""
        checks = {}
        marks = {}
        for iid, var in self._checks.items():
            node = self._nmap.get(iid)
            if node is None:
                continue
            checks[node.url] = bool(var.get())
            if node.mark is not None:
                marks[node.url] = node.mark

        t_chks = {}
        for thr_iid, var in self._thread_checks.items():
            thr = self._thread_map.get(thr_iid)
            if thr is None:
                continue
            t_chks[thr.get('url', '')] = bool(var.get())

        self.project.save_checks({
            'checks': checks,
            'marks': marks,
            'thread_checks': t_chks,
        })

    def _on_tree_click(self, event):
        col = self.tree.identify_column(event.x)
        iid = self.tree.identify_row(event.y)
        if not iid:
            return
        if col == "#1":
            # Клик по чекбоксу темы
            if iid.startswith("THR::"):
                var = self._thread_checks.get(iid)
                if var is not None:
                    new_val = not var.get()
                    var.set(new_val)
                    sym = "✅" if new_val else "☐"
                    try:
                        vals = list(self.tree.item(iid, "values"))
                        vals[0] = sym
                        self.tree.item(iid, values=vals)
                    except Exception:
                        pass

                    # Обновляем thread_checks в узле
                    thr = self._thread_map.get(iid)
                    if thr:
                        parent_iid = self.tree.parent(iid)
                        node = self._nmap.get(parent_iid)
                        if node:
                            node.thread_checks[thr.get('url', '')] = new_val

                    # Если тема включена — проставляем галочку родительскому
                    # подразделу и всем его предкам вплоть до корня.
                    # Если тема снята — проверяем: если у родителя не осталось
                    # ни одной выбранной темы И ни одного выбранного потомка,
                    # снимаем его галочку тоже.
                    parent_iid = self.tree.parent(iid)
                    if new_val:
                        # Включаем всех предков (только визуально, без рекурсии вниз)
                        cur_iid = parent_iid
                        while cur_iid and cur_iid in self._checks:
                            self._set_check_visual_only(cur_iid, True)
                            cur_iid = self.tree.parent(cur_iid)
                    else:
                        # Пересчитываем — возможно у предка ещё есть выбранные темы
                        self._propagate_thread_selection()

                    self._save_state()
                return

            # Якорные категории без чекбокса — пропускаем
            if iid not in self._checks:
                return
            cur = self._checks[iid].get()
            self._set_check(iid, not cur, recursive=True)
            self._save_state()

    def _on_tree_select(self, _event):
        sel = self.tree.selection()
        if not sel:
            return
        url = sel[0]
        node = self._nmap.get(url)
        if node:
            self._sel_url = url
            mark_txt = MARK_LABELS.get(node.mark, "") or "нет"
            self.v_sel_title.set(f"{node.title}\n{node.url}\nТем: {node.thread_count}   Пометка: {mark_txt}")
            return
        thr = self._thread_map.get(url)
        if thr:
            self._sel_url = None
            self.v_sel_title.set(f"📄 {thr.get('title', '')}\n{thr.get('url', '')}\n(двойной клик — открыть в браузере)")

    def _check_all(self):
        for iid in self._checks:
            self._set_check(iid, True, recursive=False)

    def _uncheck_all(self):
        for iid in self._checks:
            self._set_check(iid, False, recursive=False)

    def _expand_all(self):
        def expand(iid):
            self.tree.item(iid, open=True)
            for ch in self.tree.get_children(iid):
                expand(ch)
        for iid in self.tree.get_children():
            expand(iid)

    def _collapse_all(self):
        def collapse(iid):
            self.tree.item(iid, open=False)
            for ch in self.tree.get_children(iid):
                collapse(ch)
        for iid in self.tree.get_children():
            collapse(iid)

    def _choose_prompt(self):
        p = filedialog.askopenfilename(
            title="Выберите файл промта",
            filetypes=[("Текстовые файлы", "*.txt"), ("Все файлы", "*.*")],
        )
        if p:
            self._prompt_path = p
            self.v_prompt_path.set(p)
            self.project.save_config({"prompt_path": p})
            self._log(f"📄 Промт выбран: {p}", "info")

    def _load_prompt_text(self):
        if not self._prompt_path:
            return ""

        prompt_path = str(self._prompt_path).strip()

        if not prompt_path:
            return ""

        p = Path(prompt_path)

        try:
            return p.read_text(encoding="utf-8-sig")
        except Exception:
            try:
                return p.read_text(encoding="cp1251")
            except Exception as e:
                self._log(f"⚠ Не удалось прочитать промт: {p} ({e})", "warn")
                get_logger().warning(f"prompt load failed {p}: {e}")
                return ""

    def _apply_workers(self):
        if self._running:
            messagebox.showinfo("Занято", "Дождитесь завершения текущей операции.")
            return
        n = max(1, min(64, self.v_workers.get()))
        try:
            self.project.save_config({"workers": n})
        except Exception:
            pass
        old = self.executor
        self.executor = ThreadPoolExecutor(max_workers=n)
        old.shutdown(wait=False)
        self._log(
            f"⚙ Потоков установлено: {n} "
            f"(быстрый путь ⚡cookies реально параллелится; глобальный рейт-лимитер "
            f"держит безопасный темп — сайт не перегружается; "
            f"запросы через браузер 🌐 всегда строго последовательны)",
            "info"
        )
        if n > SAFE_WORKERS_FAST:
            self._log(
                f"⚠ Потоков больше {SAFE_WORKERS_FAST}: риск бана резко растёт. "
                f"Снизьте до 1–{SAFE_WORKERS_FAST} или включите 🐢 Тихий режим.", "warn"
            )

    def _ensure_cookies_loaded(self) -> bool:
        if self.http.get_cookies():
            return True
        cookies = self.project.load_cookies()
        if cookies:
            self.http.set_cookies(cookies)
            self._log(f"🔑 Загружено cookies из файла: {len(cookies)}", "info")
            return True
        return False

    def _verify_fast_session_async(self, reason=""):
        """Проверка быстрого пути в фоне с вердиктом в лог (v1.19)."""
        def worker():
            try:
                res = self.http.verify_fast_session()
            except Exception as e:
                try:
                    n = len(self.http.get_cookies())
                except Exception:
                    n = 0
                res = {"ok": False, "cause": f"error:{type(e).__name__}", "kb": 0,
                       "marker": "", "cookies_n": n, "xf_user": False}
            self.after(0, lambda: self._report_fast_verify(res, reason))
        threading.Thread(target=worker, daemon=True).start()

    def _check_fast_session_click(self):
        if self._running:
            messagebox.showinfo("Занято", "Дождитесь завершения текущей операции.")
            return
        if not self.http.get_cookies():
            try:
                cks = self.project.load_cookies()
            except Exception:
                cks = []
            if cks:
                self.http.set_cookies(cks)
        if not self.http.get_cookies():
            messagebox.showwarning("Нет cookies", "Сначала нажмите «🔑 Войти».")
            return
        self._log("🍪 Проверка быстрого пути: запрашиваю главную форума через ⚡cookies…", "info")
        self._verify_fast_session_async("ручная проверка")

    def _report_fast_verify(self, res, reason=""):
        res = res or {}
        where = f" ({reason})" if reason else ""
        n = res.get("cookies_n", 0)
        xu = "есть" if res.get("xf_user") else "НЕТ"
        cause = res.get("cause", "")
        if res.get("ok"):
            self._log(f"✅ Быстрый путь РАБОТАЕТ{where}: форум узнал вас по cookies "
                      f"(маркер: {res.get('marker')}, cookies: {n}, xf_user: {xu}). "
                      f"Запросы пойдут через ⚡.", "ok")
        elif cause == "no_cookies":
            self._log(f"❌ Нет cookies для проверки{where}: нажмите «🔑 Войти».", "err")
        elif cause in ("login_form", "redirect_login"):
            _a = res.get("analysis") if isinstance(res.get("analysis"), dict) else {}
            _extra = ""
            if _a:
                _extra = (f" Разбор страницы: заголовок «{_a.get('title', '?')}», "
                          f"маркеров входа: {len(_a.get('markers', []))} "
                          f"(сильных: {len(_a.get('strong', []))}), "
                          f"контент: {'есть' if _a.get('content_ok') else 'нет'}.")
            self._log(f"❌ Форум НЕ узнал сессию{where} ({cause}, cookies: {n}, xf_user: {xu}).{_extra} "
                      f"Быстрый путь будет падать в браузер. Нажмите «🔑 Войти» заново "
                      f"(с галочкой «Запомнить меня»).", "err")
        else:
            self._log(f"⚠ Быстрый путь под вопросом{where}: {cause} "
                      f"(cookies: {n}, xf_user: {xu}). Если форум требует вход — "
                      f"нажмите «🔑 Войти» заново.", "warn")

    def _do_login(self):
        if self._running:
            return
        self._set_busy(True)
        self._log("=" * 50, "head")
        self._log("🔑 Открытие браузера для авторизации...", "head")

        def worker():
            t0 = time.monotonic()
            pid = self._start_emitter("🔑 Авторизация")
            em = self._emitter
            try:
                s_init = em.step_added(pid, "Инициализация браузера", "🌐")
                em.step_state(s_init, "running")
                browser = Browser()
                ok = browser.start(log_cb=lambda s: self._log(s, "info"), emitter=em, parent_id=s_init)
                if not ok:
                    em.step_state(s_init, "error", "Chrome не запущен")
                    em.process_finished("error", time.monotonic() - t0)
                    self._set_busy(False)
                    return
                em.step_state(s_init, "done")
                self._login_browser = browser
                s_open = em.step_added(pid, "Открытие страницы входа", "🌐")
                em.step_state(s_open, "running")
                try:
                    browser.driver.get(FORUM_LOGIN_URL)
                    em.step_state(s_open, "done")
                except Exception as e:
                    em.step_state(s_open, "error", str(e)[:80])
                    self._log(f"⚠ {e}", "warn")
                s_wait = em.step_added(pid, "Ожидание ручного входа пользователя", "⏳")
                em.step_state(s_wait, "running")
                self.after(500, self._prompt_login_complete)
            except Exception as e:
                self._log(f"❌ {e}", "err")
                get_logger().exception(e)
                em.process_finished("error", time.monotonic() - t0)
                self._set_busy(False)

        threading.Thread(target=worker, daemon=True).start()

    def _reset_browser_profile(self):
        if self._running:
            messagebox.showinfo("Занято", "Дождитесь завершения текущей операции.")
            return
        r = messagebox.askyesno(
            "Сбросить профиль браузера",
            "Это удалит сохранённый профиль Chrome (куки/кэш антибот-проверки).\n"
            "Понадобится войти на форум заново. Продолжить?"
        )
        if not r:
            return
        try:
            if PERSISTENT_PROFILE.exists():
                shutil.rmtree(PERSISTENT_PROFILE, ignore_errors=True)
            self._log("🧹 Профиль браузера сброшен. Нажмите «🔑 Войти» заново.", "warn")
        except Exception as e:
            self._log(f"⚠ Не удалось сбросить профиль: {e}", "err")

    def _prompt_login_complete(self):
        top = tk.Toplevel(self)
        top.title("Авторизация")
        top.configure(bg=C["panel"])
        top.geometry("500x260")
        top.resizable(False, False)
        top.transient(self)
        top.grab_set()

        tk.Label(top, text="🔑  Авторизация на форуме", bg=C["panel"], fg=C["accent"],
                 font=("Segoe UI", 13, "bold")).pack(pady=(20, 10))

        info = (
            "1. Перейдите в окно Chrome, которое только что открылось.\n"
            "2. Войдите на форуме своим логином и паролем.\n"
            "3. Убедитесь, что вы видите свой профиль в шапке форума.\n"
            "4. Вернитесь сюда и нажмите «✅ Я вошёл»."
        )
        tk.Label(top, text=info, bg=C["panel"], fg=C["text"], font=("Segoe UI", 9), justify="left").pack(padx=20)

        btns = tk.Frame(top, bg=C["panel"])
        btns.pack(pady=15)

        def on_ok():
            top.destroy()
            self._save_session_cookies()

        def on_cancel():
            top.destroy()
            self._log("↩ Авторизация отменена", "warn")
            try:
                if self._login_browser:
                    self._login_browser.close()
            except Exception:
                pass
            self._login_browser = None
            self._set_busy(False)

        self._mk_btn(btns, "✅ Я вошёл", on_ok, C["green"]).pack(side="left", padx=6)
        self._mk_btn(btns, "❌ Отмена", on_cancel, C["red"]).pack(side="left", padx=6)

    def _save_session_cookies(self):
        def worker():
            t0 = time.monotonic()
            em = self._emitter
            pid = em._process_id if em else None
            try:
                s_check = None
                if em and pid:
                    s_check = em.step_added(pid, "Проверка авторизации", "🍪")
                    em.step_state(s_check, "running")
                browser = self._login_browser
                logged = browser.check_logged_in(FORUM_BASE_URL, log_cb=lambda s: self._log(s, "info"))
                cookies = browser.export_cookies()
                if em and s_check:
                    if logged:
                        em.step_state(s_check, "done", f"xf_user найден, {len(cookies)} cookies")
                    else:
                        em.step_state(s_check, "warn", "Авторизация не подтверждена")
                s_export = None
                if em and pid:
                    s_export = em.step_added(pid, "Экспорт и сохранение cookies", "📦")
                    em.step_state(s_export, "running")
                self.project.save_cookies(cookies)
                self.http.set_cookies(cookies)
                if not has_xf_session_cookie(cookies):
                    self._log("⚠ В cookies НЕТ xf_user — похоже, вход был БЕЗ галочки «Запомнить меня»: "
                              "быстрый путь (⚡cookies) перестанет работать после закрытия браузера. "
                              "Войдите заново с галочкой.", "warn")
                self._logged_in = logged
                self._update_login_indicator()
                names = sorted(set(c.get("name", "") for c in cookies))
                self._log(
                    f"🍪 Экспортировано cookies: {len(cookies)} ({', '.join(names[:10])}{'…' if len(names) > 10 else ''})",
                    "debug")
                if logged:
                    self._log("✅ Авторизация подтверждена, cookies сохранены", "ok")
                    self._status("✅ Авторизация успешна. Теперь можно сканировать.")
                else:
                    self._log("⚠ Cookies сохранены, но авторизация не подтверждена", "warn")
                if em and s_export:
                    em.step_state(s_export, "done", f"{len(cookies)} cookies")
                s_fb = None
                if em and pid:
                    s_fb = em.step_added(pid, "Привязка браузера как фолбэк", "🌐")
                    em.step_state(s_fb, "running")
                with self._fetch_browser_lock:
                    self._fetch_browser = browser
                self.http.browser_fallback = self._browser_fetch
                self._login_browser = None
                self._log("🌐 Браузер оставлен активным как фолбэк для обхода антибота при сканировании", "info")
                if em and s_fb:
                    em.step_state(s_fb, "done")
                # v1.19: сразу проверяем, узнаёт ли форум cookies быстрым путём.
                self._verify_fast_session_async("после входа")
                if em:
                    em.process_finished("done", time.monotonic() - t0)
            except Exception as e:
                self._log(f"❌ {e}", "err")
                get_logger().exception(e)
                if em:
                    em.process_finished("error", time.monotonic() - t0)
            finally:
                self._set_busy(False)

        threading.Thread(target=worker, daemon=True).start()

    def _browser_fetch(self, url: str):
        """
        Вызывается из HttpClient для получения HTML через реальный браузер.

        ⚠ КРИТИЧНО: физически это ОДНО окно/вкладка Selenium — оно не может
        обслуживать два запроса одновременно. Раньше лок держался только на
        момент чтения self._fetch_browser, а сам driver.get(url) выполнялся
        БЕЗ защиты. Из-за этого при параллельном сканировании нескольких
        выбранных разделов (Scanner.scan_deep запускает DFS каждого корня
        в отдельном потоке) один поток успевал перенаправить вкладку на
        свой URL прямо во время ожидания загрузки у другого потока —
        в результате оба потока получали содержимое ОДНОЙ и той же
        (последней загруженной) страницы. Это и проявлялось как
        "дублирование" названий разделов/тем при множественном выборе.

        Исправление: лок держится на ВСЁ время запроса (навигация + ожидание
        загрузки), поэтому запросы к браузеру теперь строго очередь,
        один за другим, независимо от того, сколько потоков сканирует.
        """
        with self._fetch_browser_lock:
            browser = self._fetch_browser
            if not browser or not browser.driver:
                return None
            try:
                return browser.get(url)
            except Exception as e:
                get_logger().warning(f"_browser_fetch({url}): {e}")
                return None

    def _ensure_fetch_browser(self, log_cb=None, emitter=None, parent_id=None) -> bool:
        cb = log_cb or self._log
        em = emitter
        with self._fetch_browser_lock:
            if self._fetch_browser and self._fetch_browser.driver:
                return True
        cookies = self.http.get_cookies() or self.project.load_cookies()
        if cookies:
            self.http.set_cookies(cookies)
        else:
            cb("⚠ Сохранённых cookies нет — пробую запустить браузер на сохранённом профиле...", "warn")
        cb("🌐 Запускаю фоновый браузер для обхода антибота "
           "(используются сохранённые cookies, повторный логин не нужен)...", "info")
        s_start = None
        if em and parent_id:
            s_start = em.step_added(parent_id, "Запуск нового браузера", "🌐")
            em.step_state(s_start, "running")
        browser = Browser()
        ok = browser.start(log_cb=lambda s: self._log(s, "debug"), emitter=em, parent_id=s_start)
        if not ok or not browser.driver:
            cb("❌ Не удалось поднять фоновый браузер-фолбэк", "err")
            if em and s_start:
                em.step_state(s_start, "error")
            return False
        if em and s_start:
            em.step_state(s_start, "done")
        s_nav = None
        if em and parent_id:
            s_nav = em.step_added(parent_id, "Переход на форум и применение cookies", "🌐")
            em.step_state(s_nav, "running")
        try:
            browser.driver.get(FORUM_BASE_URL)
            time.sleep(1)
            applied = 0
            for c in cookies:
                name = c.get("name")
                if not name:
                    continue
                try:
                    browser.driver.add_cookie({
                        "name": name,
                        "value": c.get("value", ""),
                        "domain": (c.get("domain") or FORUM_HOST).lstrip("."),
                        "path": c.get("path", "/"),
                    })
                    applied += 1
                except Exception:
                    pass
            browser.driver.get(FORUM_BASE_URL)
            time.sleep(2)
            cb(f"  🍪 В браузер применено cookies: {applied}/{len(cookies)}", "debug")
            try:
                _logged = browser.check_logged_in(FORUM_BASE_URL, log_cb=lambda s: None)
            except Exception:
                _logged = False
            if _logged:
                try:
                    _fresh = browser.export_cookies()
                except Exception:
                    _fresh = []
                if _fresh:
                    self.http.set_cookies(_fresh)
                    self._logged_in = True
                    try:
                        self.project.save_cookies(_fresh)
                    except Exception:
                        pass
                    cb(f"  ✅ Сессия в браузере активна, cookies обновлены ({len(_fresh)})", "ok")
                    try:
                        self._update_login_indicator()
                    except Exception:
                        pass
            else:
                cb("  ⚠ В браузере нет активной сессии — для закрытых разделов нужен ручной вход («🔑 Войти»)",
                   "warn")
            if em and s_nav:
                em.step_state(s_nav, "done", f"{applied}/{len(cookies)} cookies")
        except Exception as e:
            cb(f"⚠ Ошибка применения cookies в браузере: {e}", "warn")
            if em and s_nav:
                em.step_state(s_nav, "warn", str(e)[:60])
        with self._fetch_browser_lock:
            self._fetch_browser = browser
        self.http.browser_fallback = self._browser_fetch
        cb("✅ Фоновый браузер-фолбэк готов и будет использоваться при блокировках", "ok")
        return True

    def _do_scan(self):
        if self._running:
            return
        self._ensure_cookies_loaded()
        if not self.http.get_cookies():
            r = messagebox.askyesno(
                "Нет авторизации",
                "Cookies не найдены. Продолжить без авторизации?"
            )
            if not r:
                return
        self._abort.clear()
        self._set_busy(True)
        self._log("=" * 50, "head")
        self._log("🚀 ЭТАП 1: Сканирование верхнего уровня форума", "head")
        self._log("   После завершения — пометьте разделы и нажмите «Этап 2»", "info")

        def worker():
            t0 = time.monotonic()
            pid = self._start_emitter("🔍 Этап 1: Сканирование")
            em = self._emitter
            try:
                self.http.fast_enabled = self.v_fast.get()
                self.http.ensure_real_ua()
                s_tr = em.step_added(pid, "Выбор транспорта", "⚡")
                em.step_state(s_tr, "running")
                if self.http.fast_enabled and self.http.get_cookies():
                    em.step_state(s_tr, "done", "⚡ cookies, браузер — по требованию")
                    self._log("⚡ Быстрый режим: запросы через cookies, браузер поднимется только при блокировках",
                              "info")
                else:
                    s_fb = em.step_added(s_tr, "Проверка фолбэк-браузера", "🌐")
                    em.step_state(s_fb, "running")
                    if not self._ensure_fetch_browser(emitter=em, parent_id=s_fb):
                        em.step_state(s_fb, "error", "Браузер недоступен")
                        em.step_state(s_tr, "error", "Нет транспорта")
                        em.process_finished("error", time.monotonic() - t0)
                        self._log(
                            "❌ Сканирование отменено: нет cookies для быстрого пути, браузер недоступен. "
                            "Сначала нажмите «🔑 Войти».", "err"
                        )
                        self._status("❌ Нет транспорта — сканирование отменено")
                        return
                    em.step_state(s_fb, "done")
                    em.step_state(s_tr, "done", "🌐 только браузер")
                s_snap = em.step_added(pid, "Снимок состояния ДО", "📸")
                em.step_state(s_snap, "running")
                before_snapshot = self._snapshot_state()
                em.step_state(s_snap, "done")
                scanner = Scanner(
                    self.http, self.executor, self._abort,
                    log_cb=lambda s: self._log(s),
                    progress_cb=lambda s: self._status(s),
                    debug_cb=lambda s: self._debug(s),
                    emitter=em, parent_id=pid,
                    project_dir=self.project.dir,
                )
                self._scanner = scanner
                nodes = scanner.scan_top_level()
                if not nodes:
                    self._log(
                        "❌ Этап 1: пустой результат сканирования — "
                        "СУЩЕСТВУЮЩАЯ структура НЕ изменена и НЕ перезаписана.", "err"
                    )
                    self._status("❌ Этап 1: сбой, старые данные сохранены")
                    em.process_finished("error", time.monotonic() - t0)
                    return
                s_merge = em.step_added(pid, "Слияние с предыдущим состоянием", "🔄")
                em.step_state(s_merge, "running")
                merge_preserve_state(nodes, self.nodes)
                em.step_state(s_merge, "done")
                self.nodes = nodes
                s_save = em.step_added(pid, "Сохранение структуры в JSON", "💾")
                em.step_state(s_save, "running")
                if self.project.save_struct(nodes):
                    self._log(f"💾 Структура сохранена: {self.project.struct}", "ok")
                    em.step_state(s_save, "done")
                else:
                    em.step_state(s_save, "warn", "Пустая структура")
                cnt = len(nodes)
                s_build = em.step_added(pid, "Перестроение дерева в GUI", "🖥")
                em.step_state(s_build, "running")

                def _after_build():
                    self._build_tree(nodes)
                    self._restore_checks()
                    self._save_state()
                    after_snapshot = self._snapshot_state()
                    diff = self._diff_state(before_snapshot, after_snapshot)
                    self._log_state_diff(diff, "Этап 1")
                    if em:
                        em.step_state(s_build, "done")
                        em.process_finished("done", time.monotonic() - t0)

                self.after(0, _after_build)
                self._status(
                    f"✅ Этап 1 завершён: {cnt} {plural(cnt, 'раздел', 'раздела', 'разделов')}. "
                    f"Пометьте нужные → «Этап 2»."
                )
            except Exception as e:
                self._log(f"❌ {e}", "err")
                get_logger().exception(e)
                if em:
                    em.process_finished("error", time.monotonic() - t0)
            finally:
                self._set_busy(False)
                self._scanner = None

        threading.Thread(target=worker, daemon=True).start()

    def _open_participation_window(self):
        ParticipationWindow(self, self)

    def _participation_items(self):
        items = []

        def walk(node, depth):
            items.append((node, depth))
            for child in node.children:
                walk(child, depth + 1)

        for node in self.nodes or []:
            walk(node, 0)

        return items

    def _participation_config_present(self):
        cfg = self.project.load_config()
        return "participation_sections" in cfg

    def _participation_selected_urls(self):
        cfg = self.project.load_config()

        if "participation_sections" not in cfg:
            return None

        return set(cfg.get("participation_sections", []))

    def _participation_allowed_ids(self):
        selected = self._participation_selected_urls()

        if selected is None:
            return None

        allowed = set()

        def walk(node, parent_allowed):
            current_allowed = parent_allowed or (node.url in selected)

            if current_allowed:
                allowed.add(id(node))

            for child in node.children:
                walk(child, current_allowed)

        for node in self.nodes or []:
            walk(node, False)

        return allowed

    def _node_participates(self, node, allowed_ids=None):
        if allowed_ids is None:
            allowed_ids = self._participation_allowed_ids()

        if allowed_ids is None:
            return True

        return id(node) in allowed_ids

    def _subtree_has_participation(self, node, allowed_ids=None):
        if allowed_ids is None:
            allowed_ids = self._participation_allowed_ids()

        if allowed_ids is None:
            return True

        def walk(n):
            if id(n) in allowed_ids:
                return True

            for child in n.children:
                if walk(child):
                    return True

            return False

        return walk(node)

    def _participation_top_selected_nodes(self):
        selected = self._participation_selected_urls()

        if selected is None:
            return []

        top = []

        def walk(node, parent_active):
            current_active = parent_active or (node.url in selected)

            if current_active and not parent_active:
                top.append(node)

            for child in node.children:
                walk(child, current_active)

        for node in self.nodes or []:
            walk(node, False)

        return top

    def _is_anchor_node(self, node):
        return bool(node) and ((node.url or "").startswith("/#") or (node.url or "").startswith("#"))

    def _server_root_nodes(self):
        return [node for node in self.nodes or [] if node.mark == "server"]

    def _ask_selected_servers(self, server_nodes, action_text="Продолжить"):
        result = {"nodes": None}

        cfg = self.project.load_config()
        saved_urls = set(cfg.get("selected_server_urls", []))
        has_saved = "selected_server_urls" in cfg

        top = tk.Toplevel(self)
        top.title("Выбор серверов")
        top.configure(bg=C["panel"])
        top.geometry("560x560")
        top.transient(self)
        top.grab_set()

        tk.Label(
            top,
            text="Выберите серверы для обработки:",
            bg=C["panel"], fg=C["accent"], font=("Segoe UI", 10, "bold")
        ).pack(anchor="w", padx=10, pady=(10, 2))

        btns_top = tk.Frame(top, bg=C["panel"])
        btns_top.pack(anchor="w", padx=10, pady=(0, 5))

        rows = []

        def select_all():
            for _, var in rows:
                var.set(True)

        def unselect_all():
            for _, var in rows:
                var.set(False)

        self._mk_btn(btns_top, "Выбрать все", select_all, C["green"], small=True).pack(side="left", padx=2)
        self._mk_btn(btns_top, "Убрать все", unselect_all, C["red"], small=True).pack(side="left", padx=2)

        scroll = tk.Frame(top, bg=C["panel"])
        scroll.pack(fill="both", expand=True, padx=10, pady=5)

        canvas = tk.Canvas(scroll, bg=C["panel"], highlightthickness=0)
        vsb = ttk.Scrollbar(scroll, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="right", fill="y")

        frame = tk.Frame(canvas, bg=C["panel"])
        win = canvas.create_window((0, 0), window=frame, anchor="nw")

        frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfig(win, width=e.width))

        for node in server_nodes:
            var = tk.BooleanVar(value=True if not has_saved else node.url in saved_urls)

            cb = tk.Checkbutton(
                frame,
                text=node.title,
                variable=var,
                bg=C["panel"], fg=C["text"], activebackground=C["panel"],
                selectcolor=C["entry"], font=("Segoe UI", 9),
                anchor="w"
            )

            cb.pack(fill="x", anchor="w", pady=1)
            rows.append((node, var))

        bottom = tk.Frame(top, bg=C["panel"])
        bottom.pack(fill="x", padx=10, pady=5)

        def on_ok():
            chosen = [node for node, var in rows if var.get()]

            if not chosen:
                messagebox.showwarning(
                    "Ничего не выбрано",
                    "Отметьте хотя бы один сервер.",
                    parent=top
                )
                return

            self.project.save_config({"selected_server_urls": [n.url for n in chosen]})
            result["nodes"] = chosen
            top.destroy()

        def on_cancel():
            result["nodes"] = None
            top.destroy()

        self._mk_btn(bottom, action_text, on_ok, C["green"]).pack(side="left", padx=2)
        self._mk_btn(bottom, "Отмена", on_cancel, C["red"]).pack(side="left", padx=2)

        top.protocol("WM_DELETE_WINDOW", on_cancel)
        top.wait_window()

        return result["nodes"]

    def _selected_nodes_for_stage2(self, selected_servers):
        all_server_iids = set()

        for node in self._server_root_nodes():
            iid = getattr(node, "_tree_iid", None)
            if iid:
                all_server_iids.add(iid)

        selected_server_iids = set()

        for node in selected_servers:
            iid = getattr(node, "_tree_iid", None)
            if iid:
                selected_server_iids.add(iid)

        def owner_server_iid(iid):
            cur = iid

            while cur:
                if cur in all_server_iids:
                    return cur
                cur = self.tree.parent(cur)

            return None

        def allowed_iid(iid):
            own = owner_server_iid(iid)

            if own is None:
                return True

            return own in selected_server_iids

        seen = set()
        selected = []

        for iid, var in self._checks.items():
            if not var.get():
                continue

            node = self._nmap.get(iid)
            if node is None:
                continue

            if self._is_anchor_node(node):
                continue

            if not allowed_iid(iid):
                continue

            parent_iid = self.tree.parent(iid)
            is_child_of_selected = False

            while parent_iid:
                parent_node = self._nmap.get(parent_iid)
                parent_var = self._checks.get(parent_iid)

                if parent_var and parent_var.get() and not self._is_anchor_node(parent_node) and allowed_iid(
                        parent_iid):
                    is_child_of_selected = True
                    break

                parent_iid = self.tree.parent(parent_iid)

            if is_child_of_selected:
                continue

            node_id = id(node)

            if node_id in seen:
                continue

            seen.add(node_id)
            selected.append(node)

        return selected

    def _filter_nodes_by_selected_servers(self, nodes, selected_servers):
        all_server_iids = set()

        for node in self._server_root_nodes():
            iid = getattr(node, "_tree_iid", None)
            if iid:
                all_server_iids.add(iid)

        selected_server_iids = set()

        for node in selected_servers:
            iid = getattr(node, "_tree_iid", None)
            if iid:
                selected_server_iids.add(iid)

        def owner_server_iid(iid):
            cur = iid

            while cur:
                if cur in all_server_iids:
                    return cur
                cur = self.tree.parent(cur)

            return None

        result = []

        for node in nodes:
            iid = getattr(node, "_tree_iid", None)

            if not iid:
                result.append(node)
                continue

            own = owner_server_iid(iid)

            if own is None or own in selected_server_iids:
                result.append(node)

        return result

    def _do_scan_deep(self):
        if self._running:
            return
        if not self.nodes:
            messagebox.showinfo(
                "Нет данных",
                "Сначала выполните Этап 1 (кнопка «🔍 Сканировать»)."
            )
            return
        server_nodes = self._server_root_nodes()
        if not server_nodes:
            messagebox.showwarning(
                "Нет серверов",
                "Отметьте хотя бы один корневой раздел как «Сервер»."
            )
            return
        selected_servers = self._ask_selected_servers(server_nodes, "Сканировать")
        if selected_servers is None:
            self._log("↩ Выбор серверов отменён", "warn")
            return
        selected = self._selected_nodes_for_stage2(selected_servers)
        if not selected:
            messagebox.showwarning(
                "Ничего не выбрано",
                "В выбранных серверах нет разделов с отметкой «Выбор»."
            )
            return
        self.http.fast_enabled = self.v_fast.get()
        if not (self.http.fast_enabled and self.http.get_cookies()) and not self._ensure_fetch_browser():
            messagebox.showerror(
                "Нет транспорта",
                "Нет cookies для быстрого пути и браузер недоступен.\nСначала нажмите «🔑 Войти»."
            )
            return
        self._abort.clear()
        self._set_busy(True)
        self._log("=" * 50, "head")
        self._log(f"🔍 ЭТАП 2: Глубокое сканирование {len(selected)} {plural(len(selected), 'раздел', 'раздела', 'разделов')}", "head")

        def worker():
            t0 = time.monotonic()
            _sel_total = sum(len(flatten_nodes([n])) for n in selected)
            pid = self._start_emitter(f"🔍 Этап 2: Глубокое сканирование ({len(selected)} {plural(len(selected), 'раздел', 'раздела', 'разделов')})", _sel_total)
            em = self._emitter
            self.http.fast_enabled = self.v_fast.get()
            self.http.ensure_real_ua()
            try:
                s_snap = em.step_added(pid, "Снимок состояния ДО", "📸")
                em.step_state(s_snap, "running")
                before_snapshot = self._snapshot_state()
                em.step_state(s_snap, "done")
                scanner = Scanner(
                    self.http, self.executor, self._abort,
                    log_cb=lambda s: self._log(s),
                    progress_cb=lambda s: self._status(s),
                    debug_cb=lambda s: self._debug(s),
                    emitter=em, parent_id=pid,
                )
                self._scanner = scanner
                all_known = flatten_nodes(self.nodes)
                scanner._visited = {
                    norm_url(n.url)
                    for n in all_known
                    if not self._is_anchor_node(n)
                }
                for n in selected:
                    scanner._visited.discard(norm_url(n.url))
                scanner.scan_deep(selected)
                s_save = em.step_added(pid, "Сохранение структуры", "💾")
                em.step_state(s_save, "running")
                self.project.save_struct(self.nodes)
                self._log(f"💾 Структура обновлена: {self.project.struct}", "ok")
                em.step_state(s_save, "done")
                disappeared = list(scanner.disappeared_selected_threads)
                new_found = list(scanner.new_threads_found)
                s_build = em.step_added(pid, "Перестроение дерева и восстановление", "🖥")
                em.step_state(s_build, "running")

                def _after_build():
                    self._build_tree(self.nodes)
                    self._apply_pre_scan_selection(before_snapshot)
                    self._save_state()
                    after_snapshot = self._snapshot_state()
                    diff = self._diff_state(before_snapshot, after_snapshot)
                    self._log_state_diff(diff, "Этап 2")
                    self._record_scan_history(disappeared, new_found)
                    self._save_history()
                    self.after(0, self._update_history_badge)
                    if em:
                        em.step_state(s_build, "done")
                        em.process_finished("done", time.monotonic() - t0)

                self.after(0, _after_build)
                total = sum(len(flatten_nodes([n])) for n in selected)
                self._status(f"✅ Этап 2 завершён. Просканировано: {total} {plural(total, 'узел', 'узла', 'узлов')}")
            except Exception as e:
                self._log(f"❌ {e}", "err")
                get_logger().exception(e)
                if em:
                    em.process_finished("error", time.monotonic() - t0)
            finally:
                self._set_busy(False)
                self._scanner = None

        threading.Thread(target=worker, daemon=True).start()

    def _do_stop(self):
        self._abort.set()
        if self._scanner:
            self._scanner.abort.set()
        self._log("⏹ Остановка...", "warn")
        if self._emitter:
            self._emitter.process_finished("cancelled")

    def _do_load(self):
        nodes = self.project.load_struct()
        if not nodes:
            messagebox.showinfo("Нет данных", "Структура не найдена.\nСначала выполните сканирование.")
            return
        self.nodes = nodes
        self._build_tree(nodes)
        self._restore_checks()  # ← без этого «Выбор» сбрасывался
        cnt = len(flatten_nodes(nodes))
        self._status(f"✅ Загружено. Разделов: {cnt}")
        self._log(f"💾 Загружено {cnt} разделов из файла", "ok")

    def _fetch_image_bytes(self, url: str, proxy_url: str = ""):
        t0 = time.monotonic()
        deadline = min(IMAGE_DOWNLOAD_DEADLINE, 12)
        cookies = self.http.get_cookies()

        def time_left():
            return deadline - (time.monotonic() - t0)

        def attempt(name, fn, budget):
            if time_left() <= 1:
                return None, ""
            try:
                data = fn(int(min(budget, max(3, time_left()))))
            except Exception:
                data = None
            if data:
                return data, name
            return None, ""

        if proxy_url and proxy_url.startswith("/"):
            proxy_url = urljoin(FORUM_BASE_URL, proxy_url)

        host = (urlparse(url).netloc or "").lower()
        is_forum_host = host.endswith(FORUM_HOST.lower())

        if not is_forum_host:
            data, m = attempt("requests_direct",
                              lambda b: _download_image_bytes(url, cookies=cookies, timeout=b), 6)
            if data:
                return data, m

            data, m = attempt("curl",
                              lambda b: _download_image_curl(url, timeout=b), 5)
            if data:
                return data, m

            data, m = attempt("requests_firefox_ua",
                              lambda b: _download_image_bytes(url, cookies=cookies, timeout=b, ua_variant=1), 4)
            if data:
                return data, m

            data, m = attempt("requests_mobile_ua",
                              lambda b: _download_image_bytes(url, cookies=cookies, timeout=b, ua_variant=2), 4)
            if data:
                return data, m

        if is_forum_host and self.http.fast_enabled:
            data, m = attempt("cookies_session",
                              lambda b: self.http.download_bytes(url, timeout=b), 6)
            if data:
                return data, m
            if proxy_url:
                data, m = attempt("cookies_session_proxy",
                                  lambda b: self.http.download_bytes(proxy_url, timeout=b), 5)
                if data:
                    return data, m

        if proxy_url:
            data, m = attempt("forum_proxy",
                              lambda b: _download_image_via_proxy(proxy_url, cookies=cookies, timeout=b), 6)
            if data:
                return data, m

        if not self._fetch_browser_lock.acquire(timeout=IMAGE_LOCK_WAIT):
            return None, ""

        try:
            browser = self._fetch_browser
            if browser and browser.driver:
                drv = browser.driver

                if proxy_url:
                    data, m = attempt("browser_tab_proxy",
                                      lambda b: _download_image_browser_tab(drv, proxy_url, timeout=b), 12)
                    if data:
                        return data, m

                data, m = attempt("browser_tab",
                                  lambda b: _download_image_browser_tab(drv, url, timeout=b), 12)
                if data:
                    return data, m
        finally:
            self._fetch_browser_lock.release()

        return None, ""

    def _open_keyword_config(self):
        KeywordConfigWindow(self, self)

    def _open_topic_filter(self):
        TopicFilterWindow(self, self)

    def _do_generate(self):
        if self._running:
            return
        general_nodes = find_marked_nodes(self.nodes, "general")
        server_nodes = self._server_root_nodes()
        if not server_nodes:
            messagebox.showwarning(
                "Нет серверов",
                "Отметьте хотя бы один корневой раздел как «Сервер»."
            )
            return
        selected_servers = self._ask_selected_servers(server_nodes, "Сформировать")
        if selected_servers is None:
            self._log("↩ Выбор серверов отменён", "warn")
            return
        general_nodes = self._filter_nodes_by_selected_servers(general_nodes, selected_servers)
        server_nodes = selected_servers
        if not server_nodes:
            messagebox.showwarning(
                "Нет серверов",
                "Выберите хотя бы один сервер."
            )
            return
        self.http.fast_enabled = self.v_fast.get()
        if not (self.http.fast_enabled and self.http.get_cookies()) and not self._ensure_fetch_browser():
            messagebox.showerror(
                "Нет транспорта",
                "Нет cookies для быстрого пути и браузер недоступен.\n"
                "Сначала нажмите «🔑 Войти»."
            )
            return
        self._abort.clear()
        self._set_busy(True)
        self._log("=" * 50, "head")
        self._log(
            f"🚀 Формирование документов: "
            f"{len(general_nodes)} {plural(len(general_nodes), 'общий раздел', 'общих раздела', 'общих разделов')}, "
            f"{len(server_nodes)} {plural(len(server_nodes), 'сервер', 'сервера', 'серверов')}", "head"
        )

        def worker():
            t0 = time.monotonic()
            pid = self._start_emitter("🚀 Сформировать документы")
            em = self._emitter
            try:
                self.http.fast_enabled = self.v_fast.get()
                self.http.ensure_real_ua()
                s_tr = em.step_added(pid, "Выбор транспорта", "⚡")
                em.step_state(s_tr, "running")
                if self.http.fast_enabled and self.http.get_cookies():
                    em.step_state(s_tr, "done", "⚡ cookies, браузер — по требованию")
                    self._log("⚡ Быстрый режим: темы качаются через cookies, браузер — только при блокировках",
                              "info")
                else:
                    self._ensure_fetch_browser(emitter=em, parent_id=s_tr)
                    em.step_state(s_tr, "done", "🌐 браузер")

                s_collect = em.step_added(pid, "Сбор выбранных тем", "📜")
                em.step_state(s_collect, "running")
                general_threads = self._collect_checked_threads(general_nodes)
                self._log(f"📜 Тем в «Общих правилах»: {len(general_threads)}", "info")
                server_data = []
                for sn in server_nodes:
                    lst = self._collect_checked_threads([sn])
                    server_data.append((sn, lst))
                    self._log(f"🖥 «{sn.title}»: тем {len(lst)}", "info")
                em.step_state(s_collect, "done", f"Общих: {len(general_threads)}, {len(server_data)} {plural(len(server_data), 'сервер', 'сервера', 'серверов')}")

                s_dirs = em.step_added(pid, "Создание выходных директорий", "📂")
                em.step_state(s_dirs, "running")
                out = self.project.out
                out.mkdir(parents=True, exist_ok=True)
                common_dir = out / COMMON_OUTPUT_SUBDIR
                common_dir.mkdir(parents=True, exist_ok=True)
                servers_dir = out / SERVERS_OUTPUT_SUBDIR
                servers_dir.mkdir(parents=True, exist_ok=True)
                server_dir_names = self._server_dir_names(server_nodes)
                common_topics_dir = common_dir / TOPICS_SUBDIR
                common_topics_dir.mkdir(parents=True, exist_ok=True)
                server_topics_dirs = {}
                for sn in server_nodes:
                    dir_name = server_dir_names.get(id(sn), safe_name(sn.title) or "server")
                    server_dir = servers_dir / dir_name
                    server_dir.mkdir(parents=True, exist_ok=True)
                    server_topics_dir = server_dir / TOPICS_SUBDIR
                    server_topics_dir.mkdir(parents=True, exist_ok=True)
                    server_topics_dirs[id(sn)] = server_topics_dir
                em.step_state(s_dirs, "done")

                thread_target_dirs = {}

                def add_thread_target(url, target_dir):
                    k = norm_thread_url(url)
                    if k not in thread_target_dirs:
                        thread_target_dirs[k] = []
                    if target_dir not in thread_target_dirs[k]:
                        thread_target_dirs[k].append(target_dir)

                for t in general_threads:
                    add_thread_target(t.get("url", ""), common_topics_dir)
                for sn, lst in server_data:
                    server_topics_dir = server_topics_dirs[id(sn)]
                    for t in lst:
                        add_thread_target(t.get("url", ""), server_topics_dir)
                # v1.19: привязка тем к слоям для «Истории» (общие/сервер).
                scope_map = {}
                for t in general_threads:
                    _k = norm_thread_url(t.get("url", ""))
                    if _k and "general" not in scope_map.setdefault(_k, []):
                        scope_map[_k].append("general")
                for sn, lst in server_data:
                    _stag = f"server:{sn.title}"
                    for t in lst:
                        _k = norm_thread_url(t.get("url", ""))
                        if _k and _stag not in scope_map.setdefault(_k, []):
                            scope_map[_k].append(_stag)
                all_needed = {}
                for t in general_threads:
                    k = norm_thread_url(t["url"])
                    if k not in all_needed:
                        all_needed[k] = t
                for _, lst in server_data:
                    for t in lst:
                        k = norm_thread_url(t["url"])
                        if k not in all_needed:
                            all_needed[k] = t
                thread_dates = load_thread_dates_cached(self.project.dir)
                cfg_kw = _autocheck_thread_keywords
                filtered_needed = {}
                pre_skipped = []
                for k, t in all_needed.items():
                    cached_date = thread_dates.get(k, '')
                    kw_eval = evaluate_thread_keyword_match(t.get("title", ""), cached_date, cfg_kw)
                    if kw_eval == 'uncheck':
                        pre_skipped.append(t)
                    else:
                        filtered_needed[k] = t
                if pre_skipped:
                    self._log(f"⏭ Пропущено тем (дата до фильтра): {len(pre_skipped)}", "warn")
                    for t in pre_skipped:
                        self._log(f"  • {t.get('title', '')}", "warn")
                all_needed = filtered_needed
                total = len(all_needed)
                self._log(f"📊 Итого уникальных тем для парсинга: {total}", "info")
                if total == 0:
                    self._log(
                        "⚠ Нет тем. Убедитесь что галочки стоят на "
                        "разделах которые уже прошли Этап 2.", "warn"
                    )
                    em.process_finished("warn", time.monotonic() - t0)
                    return

                s_parse = em.step_added(pid, f"Парсинг содержимого тем ({total} шт)", "⬇")
                em.step_state(s_parse, "running")
                self._log(
                    f"⬇ Парсинг содержимого {total} {plural(total, 'уникальная тема', 'уникальные темы', 'уникальных тем')} "
                    f"({self.v_workers.get()} {plural(self.v_workers.get(), 'поток', 'потока', 'потоков')})...", "head"
                )
                cache = {}
                cache_lock = threading.Lock()
                done = [0]
                done_lock = threading.Lock()
                fpo = self.v_fpo.get()
                failed_threads = []
                failed_lock = threading.Lock()
                new_topics = []
                changed_topics = []
                unchanged_count = [0]
                report_lock = threading.Lock()
                new_dates = {}
                new_dates_lock = threading.Lock()
                skipped_old_urls = []
                skipped_old_lock = threading.Lock()

                def work(item):
                    k, t = item
                    if self._abort.is_set():
                        return
                    title = t.get("title", "").strip() or "(без названия)"
                    url = t.get("url", "")
                    if "page could not be loaded" in title.lower():
                        self._log(f"    ⏩ Пропущена обработка битой ссылки из конфига: {url}", "warn")
                        with done_lock:
                            done[0] += 1
                            d = done[0]
                        em.step_progress(s_parse, d, total)
                        return
                    s_thr = em.step_added(s_parse, f"📄 {title}", "📄")
                    em.step_state(s_thr, "running")
                    self._log(f"    📄 Парсинг темы: «{title}» → {url}", "debug")
                    try:
                        s_cache = em.step_added(s_thr, "Чтение кэша", "📂")
                        em.step_state(s_cache, "running")
                        old_cache = self.project.load_thread_cache(url)
                        if old_cache:
                            em.step_state(s_cache, "cached", "кэш найден")
                        else:
                            em.step_state(s_cache, "done", "нет кэша")

                        s_content = em.step_added(s_thr, "Загрузка и парсинг контента", "🌐")
                        em.step_state(s_content, "running")
                        raw_html = self.http.get(url, emitter=em, parent_id=s_content)
                        thread_date = extract_thread_creation_date(raw_html) if raw_html else ''
                        if thread_date:
                            with new_dates_lock:
                                new_dates[norm_thread_url(url)] = thread_date
                            kw_eval = evaluate_thread_keyword_match(title, thread_date, cfg_kw)
                            if kw_eval == 'uncheck':
                                self._log(
                                    f"    ⏭ Тема пропущена (создана до указанной даты): «{title}» ({thread_date})",
                                    "warn")
                                em.step_state(s_thr, "skipped", f"дата {thread_date}")
                                em.step_state(s_content, "skipped", "дата до фильтра")
                                with skipped_old_lock:
                                    skipped_old_urls.append(norm_thread_url(url))
                                with done_lock:
                                    done[0] += 1
                                    d = done[0]
                                em.step_progress(s_parse, d, total)
                                return
                        data = parse_thread_content(self.http, t, fpo, emitter=em, parent_id=s_content, initial_html=raw_html)
                        if not data.get("ok", True):
                            em.step_state(s_content, "error", "таймаут/блок")
                            em.step_state(s_thr, "error", "не удалось загрузить")
                            self._log(
                                f"    ❌ НЕ УДАЛОСЬ загрузить «{title}»: "
                                f"сервер не ответил (timeout/блок после всех попыток). "
                                f"Тема будет ПРОПУЩЕНА в документе!", "err"
                            )
                            with failed_lock:
                                failed_threads.append(title)
                            with done_lock:
                                done[0] += 1
                                d = done[0]
                            em.step_progress(s_parse, d, total)
                            return
                        post_count = len(data.get("posts", []))
                        em.step_state(s_content, "done", f"{post_count} {plural(post_count, 'пост', 'поста', 'постов')}")

                        s_img = em.step_added(s_thr, "Обработка изображений", "🖼")
                        em.step_state(s_img, "running")
                        try:
                            data = enrich_thread_data_with_images(
                                data,
                                project_dir=self.project.dir,
                                cookies=self.http.get_cookies(),
                                base_url=data.get("url", "") or url,
                                fetcher=self._fetch_image_bytes,
                                emitter=em, parent_id=s_img,
                            )
                            em.step_state(s_img, "done")
                        except Exception as img_e:
                            em.step_state(s_img, "warn", str(img_e)[:60])
                            self._log(
                                f"    ⚠ Ошибка обработки изображений в «{title}»: {img_e}",
                                "warn"
                            )

                        s_sum = em.step_added(s_thr, "Контрольная сумма контента", "🧮")
                        em.step_state(s_sum, "running")
                        new_checksum = compute_content_checksum(data)
                        em.step_state(s_sum, "done", f"sha256: {new_checksum[:12]}...")

                        with cache_lock:
                            cache[norm_thread_url(url)] = data
                        target_dirs = thread_target_dirs.get(k, [])

                        s_cmp = em.step_added(s_thr, "Сравнение с кэшем", "📊")
                        em.step_state(s_cmp, "running")
                        if old_cache is None:
                            em.step_state(s_cmp, "new", "новая тема")
                            self.project.save_thread_cache(url, data, new_checksum, fpo=fpo)
                            self._save_topic_docx(t, data, diff_lines=None, is_new=True, target_dirs=target_dirs)
                            with report_lock:
                                new_topics.append((title, url))
                            self._log(f"    🆕 Новая тема сохранена: «{title}»", "ok")
                            em.step_state(s_thr, "new", "новая тема")
                        elif old_cache.get("fpo") is not None and old_cache.get("fpo") != fpo:
                            em.step_state(s_cmp, "done", "режим парсинга изменён")
                            self.project.save_thread_cache(url, data, new_checksum, fpo=fpo)
                            self._log(f"    ℹ Изменился режим парсинга, база обновлена без диффа: «{title}»", "info")
                            with report_lock:
                                unchanged_count[0] += 1
                            em.step_state(s_thr, "done")
                        elif old_cache.get("checksum") != new_checksum:
                            diff_lines = diff_thread_blocks(old_cache.get("data", {}), data)
                            em.step_state(s_cmp, "changed", f"{len(diff_lines)} {plural(len(diff_lines), 'изменение', 'изменения', 'изменений')}")
                            self.project.save_thread_cache(url, data, new_checksum, fpo=fpo)
                            self._save_topic_docx(t, data, diff_lines=diff_lines, is_new=False, target_dirs=target_dirs)
                            with report_lock:
                                changed_topics.append((title, url, diff_lines))
                            self._log(f"    ♻ Тема ИЗМЕНИЛАСЬ: «{title}» ({len(diff_lines)} изм.)", "warn")
                            em.step_state(s_thr, "changed", f"{len(diff_lines)} изм.")
                        else:
                            em.step_state(s_cmp, "done", "без изменений")
                            with report_lock:
                                unchanged_count[0] += 1
                            self._log(f"    ✅ Без изменений: «{title}»", "debug")
                            em.step_state(s_thr, "done")
                    except Exception as e:
                        self._log(f"    ⚠ Ошибка парсинга «{title}»: {e}", "warn")
                        get_logger().warning(f"parse {t['url']}: {e}")
                        with failed_lock:
                            failed_threads.append(title)
                        em.step_state(s_thr, "error", str(e)[:60])
                    with done_lock:
                        done[0] += 1
                        d = done[0]
                    em.step_progress(s_parse, d, total)
                    if d % 3 == 0 or d == total:
                        self._status(f"Парсинг тем: {d}/{total}")

                list(self.executor.map(work, all_needed.items()))
                if new_dates:
                    save_thread_dates(self.project.dir, new_dates)
                    self._log(f"💾 {len(new_dates)} {plural(len(new_dates), 'дата темы сохранена', 'даты тем сохранены', 'дат тем сохранено')}", "info")
                if skipped_old_urls:
                    self.after(0, lambda urls=list(skipped_old_urls): self._uncheck_threads_by_url(urls))
                em.step_state(s_parse, "done", f"{done[0]}/{total}")
                self._record_generate_history(new_topics, changed_topics, unchanged_count[0], scope_map)
                self._save_history()
                self.after(0, self._update_history_badge)
                if failed_threads:
                    self._log(
                        f"⚠ ВНИМАНИЕ: не удалось загрузить {len(failed_threads)} {plural(len(failed_threads), 'тему', 'темы', 'тем')} "
                        f"(будут отсутствовать в документе):", "warn"
                    )
                    for ft in failed_threads:
                        self._log(f"    • {ft}", "warn")

                s_lines = em.step_added(pid, "Формирование текстовых строк", "📝")
                em.step_state(s_lines, "running")
                cfg = self.project.load_config()
                self._abbreviations = cfg.get("abbreviations", {})
                self._telegraph = cfg.get("telegraph_replacements", {})
                self._topic_filters = cfg.get("topic_filters", [])
                ocr_min_global = int(cfg.get("ocr_min_chars", 0))
                general_lines = threads_to_lines(
                    general_threads, cache,
                    topic_filters=self._topic_filters,
                    abbreviations=self._abbreviations,
                    telegraph=self._telegraph,
                    ocr_min_global=ocr_min_global,
                )
                prompt_text = self._load_prompt_text()
                em.step_state(s_lines, "done", f"{len(general_lines)} {plural(len(general_lines), 'общая строка', 'общие строки', 'общих строк')}")

                s_docx = em.step_added(pid, "Формирование DOCX-файлов", "📄")
                em.step_state(s_docx, "running")
                ok_count = 0
                if prompt_text or general_lines:
                    s_common = em.step_added(s_docx, "Общие правила.docx", "📄")
                    em.step_state(s_common, "running")
                    common_path = common_dir / "Общие правила.docx"
                    common_doc = build_common_docx(prompt_text, general_lines)
                    common_doc.save(common_path)
                    self._log(f"✅ Сохранён: {common_path}", "ok")
                    em.step_state(s_common, "done")
                    ok_count += 1
                for sn, lst in server_data:
                    s_srv = em.step_added(s_docx, f"{sn.title}.docx", "📄")
                    em.step_state(s_srv, "running")
                    server_lines = threads_to_lines(
                        lst, cache,
                        topic_filters=self._topic_filters,
                        abbreviations=self._abbreviations,
                        telegraph=self._telegraph,
                        ocr_min_global=ocr_min_global,
                    )
                    dir_name = server_dir_names.get(id(sn), safe_name(sn.title) or "server")
                    server_dir = servers_dir / dir_name
                    server_dir.mkdir(parents=True, exist_ok=True)
                    doc = build_server_docx(
                        prompt_text, general_lines, sn.title, server_lines
                    )
                    fname = safe_name(sn.title) + ".docx"
                    server_path = server_dir / fname
                    doc.save(server_path)
                    self._log(f"✅ Сохранён: {server_path}", "ok")
                    em.step_state(s_srv, "done")
                    ok_count += 1
                em.step_state(s_docx, "done", f"{ok_count} {plural(ok_count, 'файл', 'файла', 'файлов')}")

                self._log(f"🏁 Готово! {plural(ok_count, 'Файл', 'Файла', 'Файлов')}: {ok_count}", "ok")
                self._status(f"✅ Готово. {plural(ok_count, 'Файл', 'Файла', 'Файлов')}: {ok_count}")
                em.process_finished("done", time.monotonic() - t0)
                self.after(800, lambda: self._open_folder(out))
            except Exception as e:
                self._log(f"❌ {e}", "err")
                get_logger().exception(e)
                if em:
                    em.process_finished("error", time.monotonic() - t0)
            finally:
                self._set_busy(False)

        threading.Thread(target=worker, daemon=True).start()

    # ── v1.17: журнал «История» ──────────────────────────────────
    def _ensure_history_loaded(self):
        try:
            dkey = str(self.project.dir)
        except Exception:
            return
        if self._history_loaded_for == dkey:
            return
        with self._history_lock:
            if self._history_loaded_for == dkey:
                return
            try:
                self._history = load_history(self.project.dir)
            except Exception:
                self._history = {"topics": {}}
            self._history_loaded_for = dkey

    def _save_history(self):
        try:
            with self._history_lock:
                data = self._history
            save_history(self.project.dir, data)
        except Exception:
            pass

    def _history_record(self, url, kind, title="", section="", checksum="", diff=None, scopes=None):
        try:
            self._ensure_history_loaded()
            with self._history_lock:
                return history_append(
                    self._history, url, kind, title=title, section=section,
                    checksum=checksum, diff=diff, first_url=url, scopes=scopes,
                )
        except Exception:
            return False

    def _record_generate_history(self, new_topics, changed_topics, unchanged_count, scope_map=None):
        n_new = 0
        for title, url in new_topics or []:
            _sc = (scope_map or {}).get(norm_thread_url(url), [])
            if self._history_record(url, "new", title=title, scopes=_sc):
                n_new += 1
        n_ch = 0
        for title, url, diff_lines in changed_topics or []:
            _sc = (scope_map or {}).get(norm_thread_url(url), [])
            if self._history_record(url, "changed", title=title, diff=diff_lines or [], scopes=_sc):
                n_ch += 1
        self._log(
            f"📜 История обновлена: {n_new} {plural(n_new, 'новая тема', 'новые темы', 'новых тем')}, "
            f"{n_ch} {plural(n_ch, 'изменённая тема', 'изменённые темы', 'изменённых тем')}, "
            f"{unchanged_count} {plural(unchanged_count, 'тема без изменений', 'темы без изменений', 'тем без изменений')}. "
            f"Подробности — кнопка «📜 История».",
            "ok" if (n_new or n_ch) else "info",
        )

    def _checked_thread_urls(self) -> set:
        """Норм. URL всех отмеченных тем (v1.23: и текущих, и исчезнувших)."""
        out = set()
        try:
            roots = self.nodes or []
        except Exception:
            return out
        for node in flatten_nodes(roots):
            try:
                checks = node.thread_checks or {}
            except Exception:
                continue
            for u, v in checks.items():
                if v and u:
                    try:
                        out.add(norm_thread_url(u))
                    except Exception:
                        pass
        return out

    def _apply_exclusions_to_checks(self) -> int:
        """Снимает галочки со всех тем под исключениями, без рескана (v1.23)."""
        try:
            excl = list(_autocheck_thread_exclude or [])
        except Exception:
            excl = []
        if not excl:
            return 0
        n = 0
        try:
            roots = self.nodes or []
        except Exception:
            roots = []
        for node in flatten_nodes(roots):
            try:
                threads = node.threads or []
            except Exception:
                continue
            for t in threads:
                try:
                    u = t.get("url", "")
                    if not u or not node.thread_checks.get(u, False):
                        continue
                    if is_excluded_by_words(t.get("title", ""), node.title, excl):
                        node.thread_checks[u] = False
                        n += 1
                except Exception:
                    continue
        try:
            for thr_iid, var in list((self._thread_checks or {}).items()):
                try:
                    niid = self.tree.parent(thr_iid)
                except Exception:
                    continue
                node = (self._nmap or {}).get(niid)
                thr = (self._thread_map or {}).get(thr_iid)
                if node is None or not thr:
                    continue
                if not node.thread_checks.get(thr.get("url", ""), False):
                    try:
                        var.set(False)
                        vals = list(self.tree.item(thr_iid, "values"))
                        vals[0] = "☐"
                        self.tree.item(thr_iid, values=vals)
                    except Exception:
                        pass
        except Exception:
            pass
        if n:
            try:
                self._save_state()
            except Exception:
                pass
            self._log(f"🚫 Исключения применены: снято {n} {plural(n, 'отметка', 'отметки', 'отметок')}", "info")
        return n

    def _record_scan_history(self, disappeared, new_found):
        # v1.23: в историю пишутся только отмеченные («Выбор») темы.
        try:
            _checked = self._checked_thread_urls()
        except Exception:
            _checked = set()
        n_f = 0
        skipped = 0
        for _section, _title, _url in new_found or []:
            try:
                _ok = norm_thread_url(_url) in _checked
            except Exception:
                _ok = False
            if not _ok:
                skipped += 1
                continue
            if self._history_record(_url, "found", title=_title, section=_section):
                n_f += 1
        n_g = 0
        for _section, _title, _url in disappeared or []:
            try:
                _ok = norm_thread_url(_url) in _checked
            except Exception:
                _ok = False
            if not _ok:
                skipped += 1
                continue
            if self._history_record(_url, "gone", title=_title, section=_section):
                n_g += 1
        if n_f or n_g:
            self._log(
                f"📜 История: {n_f} {plural(n_f, 'найденная тема', 'найденные темы', 'найденных тем')}, "
                f"{n_g} {plural(n_g, 'пропавшая тема', 'пропавшие темы', 'пропавших тем')}"
                f"{f', пропущено неотмеченных: {skipped}' if skipped else ''}. "
                f"Подробности — кнопка «📜 История» (только отмеченные темы).",
                "warn" if n_g else "ok",
            )

    def _log_state_diff(self, diff_lines, stage_label):
        if diff_lines:
            self._log(f"⚠ Изменения состояния ({stage_label}), строк: {len(diff_lines)}:", "warn")
            for ln in (diff_lines or [])[:15]:
                self._log(f"    {ln}", "warn")
        else:
            self._log(
                f"✅ Состояние галочек «Выбор» и пометок ({stage_label}) корректно — сохранено.",
                "info",
            )

    def _update_history_badge(self):
        btn = getattr(self, "btn_history", None)
        if btn is None:
            return
        try:
            self._ensure_history_loaded()
            with self._history_lock:
                topics = (self._history or {}).get("topics", {})
                n = sum(1 for t in topics.values() if (t or {}).get("events"))
            btn.config(text=f"📜 История ({n})" if n else "📜 История")
        except Exception:
            pass

    def _clear_history_topic(self, key):
        try:
            with self._history_lock:
                (self._history.get("topics", {}) or {}).pop(key, None)
            self._save_history()
            self._update_history_badge()
        except Exception:
            pass

    def _clear_history_all(self):
        try:
            with self._history_lock:
                self._history = {"topics": {}}
            self._save_history()
            self._update_history_badge()
        except Exception:
            pass

    def _show_history(self):
        try:
            self._ensure_history_loaded()
            HistoryWindow(self)
        except Exception as e:
            try:
                messagebox.showerror("История", str(e))
            except Exception:
                pass

    def _save_topic_docx(self, thread: dict, data: dict, diff_lines, is_new: bool, target_dirs=None):
        fname = safe_name(thread.get("title", "") or "thread") + ".docx"

        if not target_dirs:
            target_dirs = [self.project.out / TOPICS_SUBDIR]

        g = DocxGen()

        g.add(data.get("title", thread.get("title", "")))
        g.add(data.get("url", thread.get("url", "")))
        g.add()

        if is_new:
            g.add("=== ПЕРВЫЙ ПАРСИНГ ТЕМЫ ===")
        else:
            g.add(f"=== ТЕМА ИЗМЕНИЛАСЬ ({time.strftime('%Y-%m-%d %H:%M:%S')}) ===")

        g.add()
        g.add("--- Обнаруженные изменения ---")

        if diff_lines:
            for ln in diff_lines:
                g.add(ln)
        else:
            g.add("(контрольная сумма отличается, построчный diff не найден)")

        g.add()
        g.add("=" * 60)
        g.add()
        g.add("--- Текущее содержимое темы ---")
        g.add()

        for post in data.get("posts", []):
            g.add(f"Автор: {post.get('author', '—')}   Дата: {post.get('date', '')}")
            for ln in blocks_to_lines(post.get("blocks", [])):
                g.add(ln)
            g.add()

        saved_path = None

        for topics_dir in target_dirs:
            topics_dir.mkdir(parents=True, exist_ok=True)
            path = topics_dir / fname
            g.save(path)
            self._log(f"    💾 Тема сохранена: {path}", "ok")

            if saved_path is None:
                saved_path = path

        return saved_path


    def _collect_checked_threads(self, mark_nodes: list):
        result = []
        seen = set()

        target_node_ids = set()

        def collect_ids(node: Node):
            target_node_ids.add(id(node))
            for c in node.children:
                collect_ids(c)

        for root_node in mark_nodes:
            collect_ids(root_node)

        for thr_iid, thr_var in self._thread_checks.items():
            if not thr_var.get():
                continue

            parent_iid = self.tree.parent(thr_iid)
            node = self._nmap.get(parent_iid)

            if node is None or id(node) not in target_node_ids:
                continue

            t = self._thread_map.get(thr_iid)
            if not t:
                continue

            thr_url = t.get("url", "")
            k = norm_thread_url(thr_url)

            if k in seen:
                continue

            seen.add(k)

            clean_url = re.sub(
                r'/post-\d+$', '',
                re.sub(
                    r'/(unread|latest)$', '',
                    thr_url.split("?")[0].split("#")[0]
                )
            ).rstrip("/")

            result.append({
                "title": t.get("title", ""),
                "url": clean_url,
                "sticky": t.get("sticky", False),
            })

        self._log(
            f"    📊 _collect_checked_threads: "
            f"найдено {len(result)} {plural(len(result), 'СТРОГО выбранная тема', 'СТРОГО выбранные темы', 'СТРОГО выбранных тем')} "
            f"из {len(mark_nodes)} корневых узлов (по чекбоксу темы)",
            "info"
        )

        return result

    @staticmethod
    def _open_folder(path: Path):
        try:
            if sys.platform == "win32":
                os.startfile(path)
            elif sys.platform == "darwin":
                subprocess.run(["open", str(path)])
            else:
                subprocess.run(["xdg-open", str(path)])
        except Exception:
            pass

    def _on_close(self):
        try:
            self._abort.set()
            if self._login_browser:
                self._login_browser.close()
            with self._fetch_browser_lock:
                if self._fetch_browser:
                    self._fetch_browser.close()
                    self._fetch_browser = None
        except Exception:
            pass
        try:
            self.executor.shutdown(wait=False)
        except Exception:
            pass
        self.destroy()


if __name__ == "__main__":
    try:
        get_logger().info("Старт")
        app = App()
        app.mainloop()
    except Exception as e:
        get_logger().exception(e)
        print(f"\n❌ {e}")
    finally:
        print("\n🏁 Завершено.")