@echo off
chcp 65001 >nul
cd /d "%~dp0"
if not exist .venv\Scripts\python.exe (
  py -3 -m venv .venv
  if errorlevel 1 goto fail
)
.venv\Scripts\python.exe -m pip install -r requirements.txt
if errorlevel 1 goto fail
.venv\Scripts\python.exe app.py
if errorlevel 1 goto fail
exit /b 0
:fail
 echo 启动失败，请安装 Python 3.10 或更新版本并确认网络连接。
 pause
 exit /b 1
