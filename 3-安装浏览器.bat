@echo off
chcp 936 >nul
setlocal
cd /d "%~dp0"

echo ============================================================
echo   安装浏览器
echo ============================================================
echo.

rem ---------- 情况 1：包里带了 Chrome 安装包 ----------
if exist "vendor\browser\ChromeStandaloneSetup64.exe" (
    echo 找到随包附带的 Chrome 离线安装包：
    echo   vendor\browser\ChromeStandaloneSetup64.exe
    echo.
    echo 如果不装 Google Chrome，程序会退而使用 playwright 自带内核，
    echo 虽然也能跑，但指纹更容易被电商风控识别。
    echo.
    choice /c YN /m "现在安装 Google Chrome 吗"
    if errorlevel 2 goto bundled_chromium

    echo.
    echo 正在启动安装程序，请在弹出的窗口里点"下一步"完成安装...
    echo 安装完成后可以关闭那个窗口，回到这里按任意键继续。
    echo.
    start /wait "" "vendor\browser\ChromeStandaloneSetup64.exe"
    goto verify
)

rem ---------- 情况 2：包里带企业版 MSI ----------
if exist "vendor\browser\googlechromestandaloneenterprise64.msi" (
    echo 找到随包附带的 Chrome 企业版安装包（MSI）。
    echo.
    choice /c YN /m "现在安装 Google Chrome 吗"
    if errorlevel 2 goto bundled_chromium

    echo.
    echo 正在安装（会弹出管理员确认）...
    msiexec /i "vendor\browser\googlechromestandaloneenterprise64.msi" /qb
    goto verify
)

rem ---------- 情况 3：包里没有安装包 ----------
echo 没有找到随包附带的 Chrome 安装包。
echo.
echo 如果你的电脑能上网，可以自己下载官方安装程序：
echo   https://www.google.cn/chrome/
echo.
echo 或者改用 playwright 自带内核（约 450 MB，需要能连国外服务器）：
echo   python -m playwright install chromium
echo.
goto try_pw

:bundled_chromium
echo.
echo 那就用 playwright 自带内核。

:try_pw
echo.
echo 正在执行：python -m playwright install chromium
echo （如果这一步很慢或失败，说明要连的服务器访问不了，
echo   那就只能用随包的 Chrome 安装包，或者手动装 Chrome）
echo.
python -m playwright install chromium
if errorlevel 1 (
    echo.
    echo 安装失败。
) else (
    echo.
    echo 完成。
)
goto end

:verify
echo.
echo 检查安装结果...
if exist "%ProgramFiles%\Google\Chrome\Application\chrome.exe" (
    echo   已安装：%ProgramFiles%\Google\Chrome\Application\chrome.exe
    goto end
)
if exist "%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe" (
    echo   已安装：%ProgramFiles(x86)%\Google\Chrome\Application\chrome.exe
    goto end
)
if exist "%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe" (
    echo   已安装（当前用户）：%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe
    goto end
)
echo   没检测到 Chrome，可能安装没完成，或者装到了别的用户目录。
echo   可以先运行"双击启动.bat"，用"环境自检"看程序认到的是哪个浏览器。

:end
echo.
echo ============================================================
echo 按任意键关闭
echo ============================================================
pause >nul
