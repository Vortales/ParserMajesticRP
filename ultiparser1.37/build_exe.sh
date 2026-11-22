#!/bin/bash
set -e
echo "========================================================"
echo "  Сборка ultiparser 1.37 в исполняемый файл (.exe / elf)"
echo "========================================================"
python3 -m pip install --upgrade pyinstaller requests beautifulsoup4 python-docx lxml selenium pillow cryptography || true
pyinstaller --clean -y ultiparser1.37.spec
echo "Готово! Результат в папке dist/"
