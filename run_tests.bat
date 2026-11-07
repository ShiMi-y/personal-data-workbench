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

echo 使用解释器: %PYEXE%
"%PYEXE%" -V
"%PYEXE%" -m unittest discover -s tests -v
pause
exit /b 0
