@echo off
chcp 65001 >nul
setlocal

:: ====== 請把下面兩行改成您自己電腦上的檔案路徑（只要設定一次，之後不用再改）======
set SOURCE=C:\Users\您的帳號\Desktop\momo改價\總表.xlsm
set TEMPLATE=C:\Users\您的帳號\Desktop\momo改價\momo範本.xls
:: ================================================================

if not exist "%SOURCE%" (
    echo 找不到總表檔案，請確認上面 SOURCE 路徑是否正確：
    echo %SOURCE%
    pause
    exit /b 1
)
if not exist "%TEMPLATE%" (
    echo 找不到 momo 範本檔案，請確認上面 TEMPLATE 路徑是否正確：
    echo %TEMPLATE%
    pause
    exit /b 1
)

set /p MONTH=請輸入要處理的月份（格式 YYYY-MM，例如 2026-08）：
set MONTHTAG=%MONTH:-=%

python "%~dp0generate_momo_batch.py" ^
    --source "%SOURCE%" ^
    --template "%TEMPLATE%" ^
    --month %MONTH% ^
    --output "%~dp0momo_價格異動_%MONTHTAG%.xls" ^
    --missing-output "%~dp0momo_對照不到_%MONTHTAG%.csv"

echo.
echo 執行完成，產出的檔案在：%~dp0
echo   momo_價格異動_%MONTHTAG%.xls
echo   momo_對照不到_%MONTHTAG%.csv
echo.
pause
