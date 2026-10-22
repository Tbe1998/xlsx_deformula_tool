@echo off
chcp 65001 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

rem Windows 控制台默认用本地代码页，中文路径可能让 Python 打印时崩溃，统一强制 UTF-8
set PYTHONUTF8=1
set PYTHONIOENCODING=utf-8

echo ============================================================
echo   Excel 公式转数值工具 —— Windows 一键打包 (PyInstaller)
echo ============================================================
echo.

rem ---------- 1. 找 Python ----------
set "PY="
py -3 -c "import sys,tkinter" >nul 2>&1 && set "PY=py -3"
if not defined PY (
    python -c "import sys,tkinter" >nul 2>&1 && set "PY=python"
)
if not defined PY (
    echo [提示] 没找到可用的 Python（需要 3.9 以上、且带 tkinter）。
    where winget >nul 2>&1
    if !errorlevel!==0 (
        choice /c YN /m "是否现在用 winget 安装 Python 3.12"
        if !errorlevel!==1 (
            echo 正在通过 winget 安装 Python 3.12 ...
            winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
            py -3 -c "import sys,tkinter" >nul 2>&1 && set "PY=py -3"
        )
    )
)
if not defined PY (
    echo.
    echo [失败] 请先安装 Python：https://www.python.org/downloads/windows/
    echo        安装时勾选 "Add python.exe to PATH"（安装包自带 tkinter，无需另装）。
    echo        装好后重新双击本脚本即可。
    pause
    exit /b 1
)
echo [1/5] 使用 Python：  %PY%

rem ---------- 2. 建独立构建环境 ----------
if not exist ".buildenv\Scripts\python.exe" (
    echo [2/5] 创建构建虚拟环境 .buildenv ...
    %PY% -m venv .buildenv
)
call ".buildenv\Scripts\activate.bat"
python -m pip install -q -U pip >nul 2>&1

echo [3/5] 安装 PyInstaller ...
python -m pip install -q pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple || python -m pip install -q pyinstaller
python -c "import PyInstaller" >nul 2>&1 || (echo [失败] PyInstaller 安装失败，请检查网络后重试 & pause & exit /b 1)

rem ---------- 3. 引擎自检（不通过就不打包） ----------
echo [4/5] 引擎自检 ...
python selftest.py
if errorlevel 1 (
    echo [失败] 自检未通过，已中止打包。
    pause
    exit /b 1
)

rem ---------- 4. 打包（内部名用 ASCII，打包后改中文名，避开编码坑） ----------
echo [5/5] 正在打包成单文件 exe（首次约 1-3 分钟）...
rmdir /s /q dist 2>nul
rmdir /s /q build 2>nul

set "ICON=--icon icon.ico"
if not exist "icon.ico" set "ICON="
pyinstaller --noconfirm --clean --onefile --windowed --name XlsxFormulaToValue !ICON! --distpath dist --workpath build --specpath build xlsx_deformula.py

if exist "dist\XlsxFormulaToValue.exe" goto built
if not defined ICON goto failed

echo [提示] 带图标打包失败，去掉图标重试 ...
rmdir /s /q dist 2>nul
rmdir /s /q build 2>nul
pyinstaller --noconfirm --clean --onefile --windowed --name XlsxFormulaToValue --distpath dist --workpath build --specpath build xlsx_deformula.py
if not exist "dist\XlsxFormulaToValue.exe" goto failed

:built
copy /y "dist\XlsxFormulaToValue.exe" "dist\Excel公式转数值.exe" >nul
echo.
echo ============================================================
echo  打包完成！
echo  可执行文件： %CD%\dist\Excel公式转数值.exe
echo.
echo  这个 exe 已内含 Python 运行库和 tkinter，拷到任何 64 位
echo  Windows 电脑上双击即可运行，目标机器无需安装 Python。
echo ============================================================
start "" "%CD%\dist"
pause
endlocal
exit /b 0

:failed
echo.
echo [失败] 打包出错，请把上面的报错信息发给管理员。
pause
endlocal
exit /b 1
