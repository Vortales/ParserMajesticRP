#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Миграция данных ultiparser1.16 -> ultiparser1.17.

Копирует рабочую папку форума (structure.json, checks.json, config.json,
cookies.json, thread_dates.json, _cache, images, *.txt с промтом) из
старой папки версии в папку этого скрипта (ultiparser1.17/).

По умолчанию НЕ копирует «Результаты/*.docx» (они перегенерируются кнопкой
«Сформировать документы»). Чтобы скопировать и их: --with-results.

Также чинит prompt_path в config.json, если он указывал на старую папку.

Запуск:
    python migrate_data.py
    python migrate_data.py --with-results
    python migrate_data.py --src C:\\путь\\к\\ultiparser1.16 --force
"""

import argparse
import json
import shutil
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

STATE_FILES = [
    "structure.json",
    "checks.json",
    "config.json",
    "cookies.json",
    "thread_dates.json",
    "history.json",
]
STATE_DIRS = ["_cache", "images", "logs", "_logs"]


def find_src(explicit: str | None) -> Path:
    """Ищет корень данных старой версии (папку, внутри которой forum.majestic-rp.ru)."""
    candidates: list[Path] = []
    if explicit:
        candidates.append(Path(explicit))
    candidates += [
        HERE.parent / "ultiparser1.16",
        Path.cwd() / "ultiparser1.16",
        HERE / "ultiparser1.16",
        HERE.parent / "ultiparser1.15",  # можно мигрировать и напрямую из старых версий
        Path.cwd() / "ultiparser1.15",
        HERE.parent / "ultiparser1.14",
        Path.cwd() / "ultiparser1.14",
        HERE.parent / "ServerParser",  # legacy-имя, если папку не переименовали
        Path.cwd() / "ServerParser",
    ]
    for c in candidates:
        try:
            c = c.resolve()
        except Exception:
            continue
        if c.name == "forum.majestic-rp.ru" and c.is_dir():
            return c.parent
        if (c / "forum.majestic-rp.ru").is_dir():
            return c
    raise SystemExit(
        "Не нашёл данные 1.16 рядом. Укажите путь явно: "
        "python migrate_data.py --src <путь к ultiparser1.16>"
    )


def copy_file(src: Path, dst: Path, force: bool) -> str:
    if dst.exists() and not force:
        return "skip (уже есть, без --force не трогаю)"
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copy2(src, dst)
    return "скопирован"


def copy_dir(src: Path, dst: Path, force: bool) -> str:
    if dst.exists() and not force:
        return "skip (уже есть, без --force не трогаю)"
    dst.parent.mkdir(parents=True, exist_ok=True)
    shutil.copytree(src, dst, dirs_exist_ok=True)
    return "скопирована"


def fix_prompt_path(cfg_path: Path) -> None:
    """Меняет сегмент старой папки в prompt_path на ultiparser1.17."""
    try:
        cfg = json.loads(cfg_path.read_text(encoding="utf-8-sig"))
    except Exception as e:
        print(f"  ⚠ Не смог прочитать {cfg_path.name}: {e}")
        return
    pp = str(cfg.get("prompt_path") or "")
    if not pp:
        print("  ℹ prompt_path в конфиге пуст — выберете промт в окне программы.")
        return
    new_pp = pp
    for old_seg in ("ServerParser", "ultiparser1.14", "ultiparser1.15", "ultiparser1.16"):
        new_pp = new_pp.replace(old_seg, "ultiparser1.17")
    if new_pp != pp:
        cfg["prompt_path"] = new_pp
        cfg_path.write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"  🔧 prompt_path обновлён:\n     было:  {pp}\n     стало: {new_pp}")
    else:
        print(f"  ℹ prompt_path без изменений: {pp}")
    if not Path(new_pp).exists():
        print("  ⚠ Файл промта по этому пути НЕ найден — проверьте путь и поправьте в программе.")


def main() -> None:
    ap = argparse.ArgumentParser(description="Миграция данных 1.16 -> 1.17")
    ap.add_argument("--src", default=None, help="Путь к папке ultiparser1.16 (если не рядом)")
    ap.add_argument("--with-results", action="store_true",
                    help="Копировать и папку «Результаты» (docx перегенерируются и без этого)")
    ap.add_argument("--force", action="store_true",
                    help="Перезаписывать уже существующие файлы в 1.17")
    args = ap.parse_args()

    src_root = find_src(args.src)
    src_data = src_root / "forum.majestic-rp.ru"
    dst_data = HERE / "forum.majestic-rp.ru"
    print(f"Источник: {src_data}")
    print(f"Назначение: {dst_data}")
    if src_data.resolve() == dst_data.resolve():
        raise SystemExit("Источник и назначение совпадают — копировать нечего.")
    dst_data.mkdir(parents=True, exist_ok=True)

    print("\n— Файлы состояния —")
    for name in STATE_FILES:
        s = src_data / name
        if not s.exists():
            print(f"  · {name}: нет в источнике, пропуск")
            continue
        print(f"  · {name}: {copy_file(s, dst_data / name, args.force)}")

    print("\n— Текстовые файлы (промт и т.п.) —")
    txts = sorted(src_data.glob("*.txt"))
    if not txts:
        print("  · *.txt не найдены")
    for s in txts:
        print(f"  · {s.name}: {copy_file(s, dst_data / s.name, args.force)}")

    print("\n— Папки кэша/логов —")
    for name in STATE_DIRS:
        s = src_data / name
        if not s.is_dir():
            print(f"  · {name}/: нет в источнике, пропуск")
            continue
        print(f"  · {name}/: {copy_dir(s, dst_data / name, args.force)}")

    print("\n— Результаты —")
    if args.with_results:
        s = src_data / "Результаты"
        if s.is_dir():
            print(f"  · Результаты/: {copy_dir(s, dst_data / 'Результаты', args.force)}")
        else:
            print("  · Результаты/: нет в источнике, пропуск")
    else:
        print("  · пропущены (без --with-results). Их пересоздаст кнопка «Сформировать документы».")

    print("\n— Конфиг —")
    if (dst_data / "config.json").exists():
        fix_prompt_path(dst_data / "config.json")
    else:
        print("  · config.json нет — будет создан при первом запуске.")

    print("\n✅ Готово. Запускайте:  python ultiparser1.17.py  (из папки ultiparser1.17)")
    print("   Галочки, пометки «Общие/Сервер», темы и кэш парсинга — на месте.")


if __name__ == "__main__":
    main()
    if sys.platform == "win32" and sys.stdin.isatty():
        input("\nНажмите Enter, чтобы закрыть...")
