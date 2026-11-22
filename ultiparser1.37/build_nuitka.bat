@echo off
chcp 65001 > nul
echo ========================================================
echo   Сборка ultiparser 1.37 в standalone папку через Nuitka
echo ========================================================
echo.

python -m pip install --upgrade pip nuitka requests beautifulsoup4 python-docx lxml selenium pillow cryptography zstandard

python -m nuitka --standalone --enable-plugin=tk-inter --include-package-data=docx --include-package=requests --include-package=bs4 --include-package=docx --include-package=lxml --include-package=PIL --include-package=selenium --include-package=cryptography --assume-yes-for-downloads --output-filename=ultiparser1.37.exe --output-dir=dist_nuitka --windows-console-mode=force ultiparser1.37.py

echo.
if exist "dist_nuitka\ultiparser1.37.dist\ultiparser1.37.exe" (
    echo ========================================================
    echo   [УСПЕХ] Папка собрана: dist_nuitka\ultiparser1.37.dist\
    echo ========================================================
) else (
    echo [ОШИБКА] Сборка Nuitka завершилась с ошибкой.
)

pause
