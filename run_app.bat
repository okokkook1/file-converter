@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo 啟動檔案轉換與合併工具...
python app.py
if errorlevel 1 (
    echo.
    echo 執行發生錯誤，請按任意鍵離開...
    pause >nul
)
