@echo off
chcp 936 >nul
setlocal enabledelayedexpansion
cd /d "%~dp0"

set "PY_EXE="
set "PIP_MIRROR=https://pypi.tuna.tsinghua.edu.cn/simple"
set "PROBLEMS=0"

echo ============================================================
echo   电商采集工具 - 环境安装
echo ============================================================
echo.
echo 这个脚本会自动完成：
echo   1. 安装 Python（如果没装）
echo   2. 安装 Python 依赖（走国内镜像）
echo   3. 安装 Node.js（如果没装，京东/淘宝签名要用）
echo   4. 安装 Google Chrome（如果没装）
echo   5. 环境自检
echo.
echo 全部离线可完成，不需要梯子。中途会弹权限确认，都点"是"。
echo.
choice /c YN /m "开始安装吗"
if errorlevel 2 goto cancelled
echo.


rem ============================================================
rem  第 1 步：找 Python
rem ============================================================
echo ------------------------------------------------------------
echo [1/5] 检查 Python
echo ------------------------------------------------------------

call :find_python

if not defined PY_EXE (
    echo   没有检测到 Python，用随包安装包装一个。
    call :install_bundled_python
    if defined BUNDLED_PY_EXE (
        set "PY_EXE=!BUNDLED_PY_EXE!"
        goto have_python
    )
    echo   安装失败，请手动安装 Python 后重试：
    echo     https://www.python.org/downloads/windows/
    set /a PROBLEMS+=1
    goto after_python
)

echo   已安装：!PY_EXE!
call :get_py_version

rem ------------------------------------------------------------
rem  版本检查：Python 3.15 太新，很多库还没出对应的 wheel
rem    （openpyxl/playwright 勉强能装，curl_cffi 这类带二进制的常装不上）
rem  这种情况优先换成包里带的 3.12，兼容性最好
rem ------------------------------------------------------------
if defined PY_NUM (
    if !PY_NUM! GEQ 315 (
        echo.
        echo   检测到 Python !PY_VER!，版本偏新。
        echo   新版本 Python 常常还没有常用库的预编译包，装依赖容易失败。
        echo.
        choice /c YN /m "   改用随包的 Python 3.12 吗（推荐）"
        if errorlevel 2 (
            echo   继续使用 Python !PY_VER!。
            goto have_python
        )
        echo.
        call :install_bundled_python
        if defined BUNDLED_PY_EXE (
            set "PY_EXE=!BUNDLED_PY_EXE!"
        ) else (
            echo   没装成，继续用原来的 Python !PY_VER!。
            set /a PROBLEMS+=1
        )
    )
)

:have_python

rem 把最终用的解释器记下来，2-开始采集.bat 会优先用它
rem 否则机器上装了多个 Python 时，启动脚本可能挑到没装依赖的那个
>"app\python_path.txt" echo !PY_EXE!

rem ============================================================
rem  第 2 步：装 Python 依赖
rem ============================================================
echo.
echo ------------------------------------------------------------
echo [2/5] 安装 Python 依赖
echo ------------------------------------------------------------
echo   镜像：!PIP_MIRROR!
echo.

"!PY_EXE!" -m pip install requests openpyxl Pillow playwright -i !PIP_MIRROR!
if errorlevel 1 (
    echo.
    echo   核心依赖安装失败。如果公司网络需要代理，请先设置好再重试。
    echo   也可以手动执行：
    echo     "!PY_EXE!" -m pip install requests openpyxl Pillow playwright -i !PIP_MIRROR!
    set /a PROBLEMS+=1
    goto after_python
)

echo.
echo   核心依赖装好了。接着装可选的 TLS 指纹库（失败不影响采集）...
"!PY_EXE!" -m pip install curl_cffi -i !PIP_MIRROR! --quiet
if errorlevel 1 (
    echo   curl_cffi 没装上。不影响功能，只是 TLS 指纹更像 Python，
    echo   风控识别概率略高。新版本 Python 常常还没出对应的 wheel，
    echo   装不上就先用着，功能不受影响。
) else (
    echo   curl_cffi 装好了（TLS 指纹像 Chrome）。
)
echo.
echo   依赖安装完成。
goto after_python

:after_python


rem ============================================================
rem  第 3 步：Node.js
rem ============================================================
echo.
echo ------------------------------------------------------------
echo [3/5] 检查 Node.js
echo ------------------------------------------------------------

where node >nul 2>nul
if not errorlevel 1 (
    for /f "delims=" %%V in ('node --version 2^>nul') do set "NODE_VER=%%V"
    echo   已安装：!NODE_VER!
    goto after_node
)

echo   没有检测到 Node.js（京东、淘宝的签名需要它），开始安装...

set "NODE_INSTALLER="
for %%F in ("vendor\node\node-*-x64.msi") do set "NODE_INSTALLER=%%~fF"

if not defined NODE_INSTALLER (
    echo   包里没有 Node 安装包。
    echo   可以手动安装：https://npmmirror.com/mirrors/node/
    echo   （没有 Node 也能跑闲鱼，只是京东和淘宝用不了）
    goto after_node
)

echo   使用随包安装包：!NODE_INSTALLER!
msiexec /i "!NODE_INSTALLER!" /qb
if errorlevel 1 (
    echo   安装可能没完成（旧版 MSI 装到默认路径后需要重开命令行）。
)

where node >nul 2>nul
if not errorlevel 1 (
    echo   Node.js 安装完成。
) else (
    echo   Node.js 装了但当前窗口还认不到，重开本脚本或重启即可。
)

:after_node


rem ============================================================
rem  第 4 步：Chrome
rem ============================================================
echo.
echo ------------------------------------------------------------
echo [4/5] 检查 Google Chrome
echo ------------------------------------------------------------

set "CHROME_OK="
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" set "CHROME_OK=1"
if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" set "CHROME_OK=1"
if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" set "CHROME_OK=1"

if defined CHROME_OK (
    echo   已安装。
    goto after_chrome
)

echo   没有检测到 Chrome，开始安装...

if not exist "vendor\browser\ChromeStandaloneSetup64.exe" (
    echo   包里没有 Chrome 安装包。
    echo   可以手动安装：https://www.google.cn/chrome/
    echo   （不装也能跑，程序会退回 playwright 自带内核，但更容易被风控）
    goto after_chrome
)

echo   使用随包安装包（约 160 MB，安装需要几分钟）...
start /wait "" "vendor\browser\ChromeStandaloneSetup64.exe"

if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    echo   Chrome 安装完成。
) else if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    echo   Chrome 安装完成。
) else if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" (
    echo   Chrome 安装完成。
) else (
    echo   没检测到安装结果。可能你取消了安装，或装到了其他用户目录。
    echo   程序仍可运行，会退回 playwright 自带内核。
)

:after_chrome


rem ============================================================
rem  第 5 步：自检
rem ============================================================
echo.
echo ------------------------------------------------------------
echo [5/5] 环境自检
echo ------------------------------------------------------------
echo.

if not defined PY_EXE (
    echo   跳过自检：没有可用的 Python。
    goto end
)

"!PY_EXE!" app\启动.py --check
set "CHECK_CODE=!errorlevel!"

echo.
echo ============================================================
if "!CHECK_CODE!"=="0" (
    echo   安装完成，环境就绪！
    echo.
    echo   现在可以双击「2-开始采集.bat」开始采集。
) else (
    echo   安装基本完成，但自检发现还有问题（见上）。
    echo   常见原因：需要重启命令行让 PATH 生效，或者依赖没装上。
)
echo ============================================================
goto end


:cancelled
echo.
echo 已取消。
goto end

:end
echo.
echo 按任意键关闭
pause >nul
exit /b


rem ============================================================
rem  子程序：找 Python
rem  依次尝试：py 启动器 / PATH 里的 python / 常见安装路径
rem ============================================================
:find_python
set "PY_EXE="

rem py 启动器：解析成真实的 python.exe 路径（"py -3" 带空格不能直接加引号用）
for /f "delims=" %%Z in ('py -3 -c "import sys;print(sys.executable)" 2^>nul') do (
    if exist "%%Z" (
        set "PY_EXE=%%Z"
        goto :eof
    )
)

rem PATH 里的 python
for /f "delims=" %%P in ('where python 2^>nul') do (
    if exist "%%P" (
        set "PY_EXE=%%P"
        goto :eof
    )
)

rem 常见安装位置（含随包静默安装的位置）
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

rem ============================================================
rem  子程序：取 Python 版本号
rem  设置 PY_VER（如 3.12.10）和 PY_NUM（如 312，便于比较大小）
rem ============================================================
:get_py_version
set "PY_VER="
set "PY_NUM="
if not defined PY_EXE goto :eof

rem 用临时文件读版本号：for /f 套多层引号在 cmd 里不可靠
set "PYVER_FILE=%TEMP%\ec_pyver_%RANDOM%%RANDOM%.txt"
"!PY_EXE!" -c "import sys;print(sys.version.split()[0])" > "!PYVER_FILE!" 2>nul

if not exist "!PYVER_FILE!" goto :eof
set /p PY_VER=<"!PYVER_FILE!"
del "!PYVER_FILE!" >nul 2>nul

if not defined PY_VER goto :eof
echo  版本：!PY_VER!

rem 3.15 -> 315，3.9 -> 309，方便按数字比较大小
for /f "tokens=1,2 delims=." %%A in ("!PY_VER!") do set /a PY_NUM=%%A*100+%%B
goto :eof


rem ============================================================
rem  子程序：用随包安装包装 Python 3.12
rem  成功后设置 BUNDLED_PY_EXE，不碰 PY_EXE —— 装失败时调用方还能用原来那个
rem  装完按已知路径取，不走 find_python（否则又被新版本抢占）
rem ============================================================
:install_bundled_python
set "BUNDLED_PY_EXE="

set "PY_INSTALLER="
for %%F in ("vendor\python\python-*-amd64.exe") do set "PY_INSTALLER=%%~fF"

if not defined PY_INSTALLER (
    echo   包里没有 Python 安装包。
    goto :eof
)

echo   使用随包安装包：!PY_INSTALLER!
echo   正在安装（当前用户，不需要管理员权限，约 1~2 分钟）...
"!PY_INSTALLER!" /quiet InstallAllUsers=0 PrependPath=1 Include_pip=1 Include_test=0 Include_launcher=1

rem 安装位置是固定的，直接找，避免又被旧版本抢占
for %%D in (
    "%LOCALAPPDATA%\Programs\Python\Python312"
    "%ProgramFiles%\Python312"
) do (
    if exist "%%~D\python.exe" (
        set "BUNDLED_PY_EXE=%%~D\python.exe"
        echo   安装完成：!BUNDLED_PY_EXE!
        goto :eof
    )
)

echo   安装后没找到 Python 3.12，可能被杀毒软件拦了，或者需要重启。
goto :eof