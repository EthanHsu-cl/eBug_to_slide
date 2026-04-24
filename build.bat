@echo off
REM Build eBug_to_slide.exe for Windows

call venv\Scripts\activate.bat
pip install pyinstaller --quiet
pyinstaller eBug_to_slide.spec --clean

echo.
echo Done. Output: dist\eBug_to_slide.exe
