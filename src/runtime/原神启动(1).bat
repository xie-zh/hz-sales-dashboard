@echo off
chcp 65001 >nul
title Python HTTP Server

:: ========== 在这里修改端口号 ==========
set PORT=8000
:: =====================================

:: 获取本机局域网IPv4地址
for /f "tokens=2 delims=:" %%a in ('ipconfig ^| findstr "IPv4" ^| findstr /v "127.0.0.1"') do (
    set "local_ip=%%a"
    goto :ip_found
)
:ip_found
set "local_ip=%local_ip: =%"

echo ========================================
echo        Python HTTP 文件服务器
echo ========================================
echo.
echo  本地访问: http://localhost:%PORT%
echo  局域网访问: http://%local_ip%:%PORT%
echo.
echo  当前目录: %cd%
echo  端口号: %PORT%
echo ========================================
echo  按 Ctrl+C 停止服务器
echo.

python -m http.server %PORT%

pause