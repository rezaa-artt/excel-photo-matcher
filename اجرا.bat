@echo off
chcp 65001 >nul
cd /d "%~dp0"

if exist "افزودن عکس به اکسل.exe" (
    start "" "افزودن عکس به اکسل.exe"
    exit /b 0
)

if exist "برنامه آماده\افزودن عکس به اکسل.exe" (
    start "" "برنامه آماده\افزودن عکس به اکسل.exe"
    exit /b 0
)

if exist "dist\ExcelPhoto.exe" (
    start "" "dist\ExcelPhoto.exe"
    exit /b 0
)

echo فایل اجرایی پیدا نشد. اول build_exe.bat را اجرا کنید.
pause
