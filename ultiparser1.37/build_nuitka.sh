#!/bin/bash
set -e
echo "========================================================"
echo "  Сборка ultiparser 1.37 через Nuitka (Linux/macOS/Cross)"
echo "========================================================"
python3 -m pip install --upgrade nuitka requests beautifulsoup4 python-docx lxml selenium pillow cryptography zstandard || true

python3 -m nuitka \
    --standalone \
    --onefile \
    --enable-plugin=tk-inter \
    --include-package-data=docx \
    --include-package=requests \
    --include-package=bs4 \
    --include-package=docx \
    --include-package=lxml \
    --include-package=PIL \
    --include-package=selenium \
    --include-package=cryptography \
    --assume-yes-for-downloads \
    --output-filename=ultiparser1.37.bin \
    --output-dir=dist_nuitka \
    ultiparser1.37.py

echo "Сборка завершена! Результат в dist_nuitka/"
