@echo off
chcp 65001 >nul
cd /d "%~dp0"

python -m pip install -r requirements.txt pyinstaller
if errorlevel 1 exit /b 1

python -m PyInstaller --noconfirm --clean ExcelPhoto.spec
if errorlevel 1 exit /b 1

if not exist "dist\ExcelPhoto.exe" (
    echo ساخت فایل اجرایی ناموفق بود.
    exit /b 1
)

if not exist "برنامه آماده" mkdir "برنامه آماده"
copy /Y "dist\ExcelPhoto.exe" "برنامه آماده\افزودن عکس به اکسل.exe" >nul
copy /Y "dist\ExcelPhoto.exe" "افزودن عکس به اکسل.exe" >nul

echo.
echo فایل اجرایی آماده شد.
echo همین فایل را برای دیگران بفرستید. پایتون لازم نیست.
