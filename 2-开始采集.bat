@echo off
chcp 936 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

set "PY_EXE="
call :find_python

if not defined PY_EXE (
    echo ============================================================
    echo   找不到 Python
    echo ============================================================
    echo.
    echo 请先双击「安装环境.bat」，它会自动装好 Python 和依赖。
    echo.
    pause
    exit /b 1
)

"!PY_EXE!" app\启动.py
if errorlevel 1 (
    echo.
    echo ============================================================
    echo   程序异常退出
    echo ============================================================
    echo.
    echo 如果是依赖缺失，先双击「安装环境.bat」补齐。
    echo 其他情况请看上面的报错信息。
    echo.
    pause
)
exit /b


rem ============================================================
rem  找 Python：py 启动器 -> PATH -> 常见安装路径
rem ============================================================
:find_python
set "PY_EXE="

rem ------------------------------------------------------------
rem  优先用 1-安装环境.bat 记下的解释器
rem  机器上装了多个 Python 时（比如 3.15 + 3.12），依赖只装在其中一个里，
rem  直接找 py -3 可能挑到没装依赖的那个，所以先认这个记录
rem ------------------------------------------------------------
if not exist "app\python_path.txt" goto :find_python_search

set "PY_RECORDED="
set /p PY_RECORDED=<"app\python_path.txt"

if not defined PY_RECORDED goto :find_python_search
if not exist "!PY_RECORDED!" goto :find_python_search

set "PY_EXE=!PY_RECORDED!"
goto :eof

:find_python_search
for /f "delims=" %%Z in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do (
    if exist "%%Z" (
        set "PY_EXE=%%Z"
        goto :eof
    )
)

for /f "delims=" %%P in ('where python 2^>nul') do (
    if exist "%%P" (
        set "PY_EXE=%%P"
        goto :eof
    )
)

for %%D in (
    "%LOCALAPPDATA%\Programs\Python\Python312"
    "%LOCALAPPDATA%\Programs\Python\Python313"
    "%LOCALAPPDATA%\Programs\Python\Python311"
    "%ProgramFiles%\Python312"
    "%ProgramFiles%\Python313"
) do (
    if exist "%%~D\python.exe" (
        set "PY_EXE=%%~D\python.exe"
        goto :eof
    )
)

goto :eof
