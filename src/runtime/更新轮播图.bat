@echo off
chcp 65001 >nul
cd /d "%~dp0"

set "PY=<HOME>\.workbuddy\binaries\python\envs\default\Scripts\python.exe"
if not exist "%PY%" set "PY=python"

echo ============================================
echo  扫描 deploy_nav\banners\ 下的板块文件夹
echo  （每个文件夹取第一张图做轮播封面）
echo ============================================
echo.

"%PY%" "_gen_banners.py"

echo.
echo 完成后把 deploy_nav 目录重新发布一次即可生效。
pause
