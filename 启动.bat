@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem --- 找 Python ---
set "PY="
where py >nul 2>nul && set "PY=py -3"
if not defined PY where python >nul 2>nul && set "PY=python"
if not defined PY (
  echo [错误] 未找到 Python，请先安装 Python 3.10+：https://www.python.org/downloads/
  pause
  exit /b 1
)

rem --- 检查并安装依赖 ---
%PY% -c "import numpy, cv2, PIL, openpyxl" >nul 2>nul
if errorlevel 1 (
  echo 首次运行，正在安装依赖（numpy/opencv/pillow/openpyxl）...
  %PY% -m pip install -r requirements.txt
)

%PY% batch_analyze.py %*
if errorlevel 1 pause
