@echo off
chcp 65001 >nul
title 分公司看板 · 新电脑一键部署
cd /d "%~dp0"

echo.
echo  ============================================================
echo    分公司看板 · 新电脑一键部署
echo ------------------------------------------------------------
echo    本脚本会在*这台电脑*上做以下几件事（约 1~3 分钟）：
echo      1  建一套独立 Python 环境（放本目录 _env，不动系统环境）
echo      2  装好两个依赖库：openpyxl / pywin32
echo      3  把源表路径 / 发布参数 / BI 取数工号改成这台电脑的
echo      4  注册两个定时任务
echo           每 5 分钟唤醒刷新  +  每晚 23:59 存历史快照
echo      5  环境体检 + 首次构建本地看板
echo.
echo    中途可以按 Ctrl+C 中止；中止不会改动任何源数据文件。
echo  ============================================================
echo.
pause

set "PYEXE="

rem --- 1) WorkBuddy 自带的 Python（装了 WorkBuddy 就有）---
for /f "delims=" %%V in ('dir /b /o-n "%USERPROFILE%\.workbuddy\binaries\python\versions" 2^>nul') do (
  if not defined PYEXE if exist "%USERPROFILE%\.workbuddy\binaries\python\versions\%%V\python.exe" set "PYEXE=%USERPROFILE%\.workbuddy\binaries\python\versions\%%V\python.exe"
)
rem --- 2) 项目里已经建好的可移植环境 ---
if not defined PYEXE if exist "%~dp0_env\Scripts\python.exe" set "PYEXE=%~dp0_env\Scripts\python.exe"
rem --- 3) 系统 PATH 里的 python ---
if not defined PYEXE for %%P in (python.exe) do set "PYEXE=%%~$PATH:P"
rem --- 4) 常规安装位置 ---
if not defined PYEXE for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
  if not defined PYEXE if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
)

if not defined PYEXE (
  echo.
  echo  [X] 这台电脑上没有找到 Python，无法自动配置环境。
  echo.
  echo      两个办法，任选其一，然后重新双击本文件：
  echo        A. 安装 Python 3.10 或更高版本
  echo           https://www.python.org/downloads/
  echo           安装第一步务必勾选 "Add python.exe to PATH"
  echo        B. 安装 WorkBuddy 客户端并登录（它自带 Python 运行时）
  echo.
  pause
  exit /b 1
)

echo.
echo  使用的 Python：%PYEXE%
echo.

"%PYEXE%" "%~dp0Hzdggbot\setup_new_pc.py" --yes
set RC=%ERRORLEVEL%

echo.
echo  ------------------------------------------------------------
if "%RC%"=="0" (
  echo   部署完成：全部体检项就绪。
) else (
  echo   部署已结束，但仍有项目需要处理 —— 请对照上面的体检表。
)
echo   完整报告：%CD%\_portable\新电脑部署报告.txt
echo   日志文件：%CD%\Hzdggbot\_bi_refresh.log
echo  ------------------------------------------------------------
echo.
pause
