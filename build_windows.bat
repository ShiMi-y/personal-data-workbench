@echo off
chcp 65001 >nul
cd /d "%~dp0"

rem 优先用 python.org 安装目录里的解释器，避免 WindowsApps 的 0 字节应用别名
rem 其次用 py 启动器，最后才退回 PATH 上的 python
set "PYEXE="
for /d %%D in ("%LOCALAPPDATA%\Programs\Python\Python3*") do (
  if not defined PYEXE if exist "%%D\python.exe" set "PYEXE=%%D\python.exe"
)
if not defined PYEXE (
  py -c "import sys" >nul 2>nul && set "PYEXE=py"
)
if not defined PYEXE (
  python -c "import sys" >nul 2>nul && set "PYEXE=python"
)
if not defined PYEXE (
  echo 未找到可用的 Python。
  echo 请安装 Python 3.10+ 并加入 PATH，或把 %LOCALAPPDATA%\Programs\Python\Launcher 加入 PATH 后重试。
  pause
  exit /b 1
)

echo [1/3] 安装打包依赖...（解释器: %PYEXE%）
"%PYEXE%" -m pip install --upgrade pip
"%PYEXE%" -m pip install pyinstaller==6.21.0 openpyxl playwright cryptography
if errorlevel 1 goto fail

echo [2/3] PyInstaller 打包...
"%PYEXE%" -m PyInstaller --noconfirm --clean --onedir --windowed ^
  --name "快手数据工作台" ^
  --icon assets\app.ico ^
  --add-data "web;web" --add-data "config.example.json;." --add-data "assets;assets" ^
  --exclude-module PIL --exclude-module pandas --exclude-module matplotlib ^
  --exclude-module tkinter --exclude-module numpy --exclude-module lxml ^
  --exclude-module charset_normalizer web_server.py
if errorlevel 1 goto fail

echo [3/3] 生成升级 zip...
if not exist "..\release" mkdir "..\release"
powershell -NoProfile -Command "Compress-Archive -Path 'dist\快手数据工作台\快手数据工作台.exe','dist\快手数据工作台\_internal' -DestinationPath '..\release\personal-data-workbench-v1.2.9-fixed-win64.zip' -Force"
if errorlevel 1 goto fail

echo.
echo 完成：%CD%\dist\快手数据工作台
echo 升级包：%CD%\..\release\personal-data-workbench-v1.2.9-fixed-win64.zip
pause
exit /b 0

:fail
echo.
echo 构建失败，请查看上方错误。
pause
exit /b 1
