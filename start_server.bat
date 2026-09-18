@echo off
chcp 65001
setlocal

REM 财务数据看板启动脚本
REM 使用 Python 内置 HTTP 服务器

cd /d "C:\Users\GBJ-0977\Desktop\financial_management"

echo =========================================
echo     财务数据看板 - 本地服务器启动
echo =========================================
echo.
echo 正在启动服务器，请稍候...
echo.

python -m http.server 8000

pause
endlocal