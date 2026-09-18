@echo off
chcp 65001
setlocal

REM 财务数据看板启动脚本
REM 使用 Python 内置 HTTP 服务器

cd /d "%~dp0"

echo =========================================
echo     财务数据看板 - 本地服务器启动
echo =========================================
echo.

echo 正在启动服务器，请稍候...
echo.

python start_server.py

pause
endlocal