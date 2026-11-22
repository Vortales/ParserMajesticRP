@echo off
chcp 65001 > nul
echo ========================================================
echo   Сборка ultiparser 1.37 в .exe через Nuitka
echo ========================================================
echo.

echo [1/3] Проверка компилятора Nuitka и зависимостей...
python -m pip install --upgrade pip nuitka requests beautifulsoup4 python-docx lxml selenium pillow cryptography zstandard
echo.

echo [2/3] Запуск компиляции в C++ и сборка бинарника...
python -m nuitka ^
    --standalone ^
    --onefile ^
    --enable-plugin=tk-inter ^
    --include-package-data=docx ^
    --include-package=requests ^
    --include-package=bs4 ^
    --include-package=docx ^
    --include-package=lxml ^
    --include-package=PIL ^
    --include-package=selenium ^
    --include-package=cryptography ^
    --assume-yes-for-downloads ^
    --output-filename=ultiparser1.37.exe ^
    --output-dir=dist_nuitka ^
    --windows-console-mode=force ^
    ultiparser1.37.py

echo.
if exist "dist_nuitka\ultiparser1.37.exe" (
    echo ========================================================
    echo   [УСПЕХ] Файл собран: dist_nuitka\ultiparser1.37.exe
    echo ========================================================
) else (
    echo [ОШИБКА] Сборка Nuitka завершилась с ошибкой. Проверьте вывод выше.
)

pause
