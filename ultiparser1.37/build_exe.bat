@echo off
chcp 65001 > nul
echo ========================================================
echo   Сборка ultiparser 1.37 в исполняемый файл (.exe)
echo ========================================================
echo.

echo [1/3] Проверка и установка PyInstaller и зависимостей...
python -m pip install --upgrade pip pyinstaller requests beautifulsoup4 python-docx lxml selenium pillow cryptography
echo.

echo [2/3] Запуск сборки через PyInstaller spec...
pyinstaller --clean -y ultiparser1.37.spec
echo.

if exist "dist\ultiparser1.37.exe" (
    echo ========================================================
    echo   [УСПЕХ] Файл собран: dist\ultiparser1.37.exe
    echo ========================================================
) else (
    echo [ОШИБКА] Сборка завершилась неудачно. Проверьте вывод выше.
)

pause
