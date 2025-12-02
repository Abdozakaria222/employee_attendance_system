@echo off
echo ========================================
echo   تحويل نظام الحضور إلى EXE
echo ========================================
echo.

echo [1/4] تثبيت PyInstaller...
pip install pyinstaller

echo.
echo [2/4] تنظيف الملفات القديمة...
if exist "build" rmdir /s /q build
if exist "dist" rmdir /s /q dist
if exist "*.spec" del /q *.spec

echo.
echo [3/4] إنشاء ملف EXE...
pyinstaller --onefile --windowed --name="نظام_الحضور" --icon=icon.ico attendance_system.py

echo.
echo [4/4] تم الانتهاء!
echo ========================================
echo ✅ ملف EXE موجود في مجلد: dist
echo ========================================
pause
