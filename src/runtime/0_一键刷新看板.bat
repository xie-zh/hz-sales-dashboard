@echo off
chcp 65001 >nul
title 分公司看板 · 一键刷新
cd /d "%~dp0"

echo.
echo  ============================================
echo    看板一键刷新（本目录只认这一个入口）
echo    取数 -^> 写Excel -^> 重建 -^> 发布
echo    正常约 1~2 分钟，请别关这个窗口
echo  ============================================
echo.
echo  开始时间 %time%
echo.

rem ---- 找 Python：项目内 _env 优先，其次 WorkBuddy 自带，再次系统 ----
set "PYEXE="
if exist "%~dp0_env\Scripts\python.exe" set "PYEXE=%~dp0_env\Scripts\python.exe"
if not defined PYEXE for /f "delims=" %%V in ('dir /b /o-n "%USERPROFILE%\.workbuddy\binaries\python\versions" 2^>nul') do (
  if not defined PYEXE if exist "%USERPROFILE%\.workbuddy\binaries\python\versions\%%V\python.exe" set "PYEXE=%USERPROFILE%\.workbuddy\binaries\python\versions\%%V\python.exe"
)
if not defined PYEXE for %%P in (python.exe) do set "PYEXE=%%~$PATH:P"
if not defined PYEXE for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
  if not defined PYEXE if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
)
if not defined PYEXE (
  echo  [X] 没找到 Python。请先双击「一键部署_新电脑.bat」完成环境配置。
  pause
  exit /b 1
)

"%PYEXE%" "%~dp0Hzdggbot\bi_refresh.py" --force
set RC=%ERRORLEVEL%

echo.
echo  ============================================
echo  退出码 %RC%
if "%RC%"=="0" echo   成功（或数据无变化，未重复发布）
if "%RC%"=="1" echo   本轮失败，线上保持上一次的数据
if "%RC%"=="2" echo   不在更新时段
if "%RC%"=="3" echo   上一轮还没跑完，请稍后再试
echo  ============================================
echo.
echo  ---- 本轮日志（最后 18 行）----
powershell -NoProfile -Command "[Console]::OutputEncoding=[Text.Encoding]::UTF8; $p='%~dp0Hzdggbot\_bi_refresh.log'; if (Test-Path $p) { Get-Content $p -Encoding UTF8 -Tail 18 }"
echo.

rem ---- 显示当前线上看板地址（从 迁移配置.json 读）----
"%PYEXE%" "%~dp0Hzdggbot\_migrate.py" url > "%TEMP%\hz_dash_url.txt" 2>nul
set "DASHURL="
if exist "%TEMP%\hz_dash_url.txt" set /p DASHURL=<"%TEMP%\hz_dash_url.txt"
del "%TEMP%\hz_dash_url.txt" >nul 2>&1
if defined DASHURL echo  公网看板：%DASHURL%

echo  完整日志：Hzdggbot\_bi_refresh.log
echo.
pause
