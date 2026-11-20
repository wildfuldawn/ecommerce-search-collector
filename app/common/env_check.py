"""环境自检：依赖、签名文件、浏览器、登录态。

分享给别人时最先跑这个 —— 缺什么一次性说清楚，而不是跑到一半报错。
"""

import importlib
import os
import shutil

from . import paths

BASE_DIR = paths.ROOT
VENDOR = paths.VENDOR
PROFILES = paths.PROFILES
OUTPUT = paths.OUTPUT

# 平台 → (vendor 子目录, 主脚本, 额外必需文件)
PLATFORM_FILES = {
    "jingdong": ("jingdong", "jd_sku.py", []),
    "taobao": ("taobao", "xuanzeshuliang.py", []),
    "xianyu": ("xianyu", "xianyu.py", []),
}

# 平台 → 判定"已登录"的身份 cookie 名
LOGIN_MARKERS = {
    "jingdong": ("pt_key", "pt_pin", "thor", "pin", "pinId"),
    "taobao": ("unb", "cookie2", "_tb_token_"),
    "xianyu": ("unb", "tracknick"),
}

PLATFORM_NAMES = {
    "jingdong": "京东",
    "taobao": "淘宝",
    "xianyu": "闲鱼",
}


def ensure_dirs():
    paths.ensure_dirs()


def module_available(name):
    try:
        importlib.import_module(name)
        return True
    except ImportError:
        return False


def chrome_path():
    """找系统 Chrome。

    按可信度依次尝试：
      1) 注册表 App Paths（Chrome 装哪都能查到）
      2) 常见安装目录
      3) PATH
      4) 各盘符的 Program Files（虚拟机里常装在非 C 盘）
    """
    # ---- 1) 注册表 ----
    try:
        import winreg
        for hive, key in (
            (winreg.HKEY_CURRENT_USER,
             r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"),
            (winreg.HKEY_LOCAL_MACHINE,
             r"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\chrome.exe"),
            (winreg.HKEY_LOCAL_MACHINE,
             r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion"
             r"\App Paths\chrome.exe"),
        ):
            try:
                with winreg.OpenKey(hive, key) as handle:
                    value, _ = winreg.QueryValueEx(handle, "")
                value = str(value).strip('"')
                if os.path.isfile(value):
                    return value
            except OSError:
                continue
    except ImportError:
        pass

    # ---- 2) 常见目录 ----
    candidates = [
        os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%PROGRAMFILES(X86)%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%LOCALAPPDATA%\Google\Chrome\Application\chrome.exe"),
        os.path.expandvars(r"%PROGRAMFILES%\Google\Chrome Beta\Application\chrome.exe"),
        r"C:\Program Files\Google\Chrome\Application\chrome.exe",
        r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe",
    ]
    for path in candidates:
        if path and os.path.isfile(path):
            return path

    # ---- 3) PATH ----
    for name in ("chrome", "chrome.exe", "google-chrome"):
        found = shutil.which(name)
        if found:
            return found

    # ---- 4) 各盘符扫一遍（只查两层，很快）----
    import string
    for letter in string.ascii_uppercase:
        drive = f"{letter}:\\"
        if not os.path.isdir(drive):
            continue
        for base in ("Program Files", "Program Files (x86)", "Google", "Apps"):
            candidate = os.path.join(
                drive, base, "Google", "Chrome", "Application", "chrome.exe")
            if os.path.isfile(candidate):
                return candidate
            candidate = os.path.join(
                drive, base, "Chrome", "Application", "chrome.exe")
            if os.path.isfile(candidate):
                return candidate

    return ""


def playwright_chromium_path():
    """找 playwright 自带的内核可执行文件；没有返回空。"""
    roots = []
    custom = os.environ.get("PLAYWRIGHT_BROWSERS_PATH")
    if custom:
        roots.append(custom)
    local = os.environ.get("LOCALAPPDATA")
    if local:
        roots.append(os.path.join(local, "ms-playwright"))
    roots.append(os.path.join(os.path.expanduser("~"), ".cache", "ms-playwright"))

    relative = [
        ("chrome-win64", "chrome.exe"),
        ("chrome-win", "chrome.exe"),
        ("chrome-linux64", "chrome"),
        ("chrome-linux", "chrome"),
        ("chrome-mac-arm64", "Chromium.app", "Contents", "MacOS", "Chromium"),
        ("chrome-mac", "Chromium.app", "Contents", "MacOS", "Chromium"),
    ]

    for root in roots:
        if not os.path.isdir(root):
            continue
        for name in sorted(os.listdir(root), reverse=True):
            if not name.startswith("chromium-"):
                continue
            base = os.path.join(root, name)
            for parts in relative:
                candidate = os.path.join(base, *parts)
                if os.path.isfile(candidate):
                    return candidate
    return ""


def bundled_installer():
    """包里自带的 Chrome 离线安装包；没有返回空。"""
    for name in ("ChromeStandaloneSetup64.exe",
                 "googlechromestandaloneenterprise64.msi"):
        path = os.path.join(BASE_DIR, "vendor", "browser", name)
        if os.path.isfile(path):
            return path
    return ""


def browser_choice():
    """决定用哪个浏览器。

    返回 (kind, value, warning)：
        ("chrome", "", "")                     —— 系统 Chrome，最好
        ("chromium", "<路径>", "<提示>")        —— playwright 内核，能用但更易被识别
        ("", "", "<该怎么办>")                  —— 都没有
    """
    chrome = chrome_path()
    if chrome:
        return "chrome", "", ""

    bundled = playwright_chromium_path()
    if bundled:
        return "chromium", bundled, (
            "没找到系统 Chrome，将改用 playwright 自带内核。\n"
            "  能用，但它不是 Google Chrome，指纹上更容易被风控识别。\n"
            "  建议装一个 Chrome：直接搜 \"Chrome 下载\" 装官方版即可。"
        )

    installer = bundled_installer()
    if installer:
        return "", "", (
            "没有可用的浏览器。包里自带 Chrome 离线安装包：\n"
            f"  {os.path.basename(installer)}\n"
            "  双击一级目录的「3-安装浏览器.bat」，或直接双击那个安装包。\n"
            "  装完重新启动本程序即可。"
        )

    return "", "", (
        "没有可用的浏览器。二选一：\n"
        "  1) 装 Google Chrome（官网 google.cn/chrome 下载）\n"
        "  2) 在命令行执行：python -m playwright install chromium （约 450 MB）"
    )


def browser_env():
    """给子进程用的浏览器环境变量。

    找到浏览器就显式给路径 —— 比让 playwright 按 channel 自己找更可靠，
    Chrome 装在非标准位置（虚拟机、绿色版）时也能用。
    """
    chrome = chrome_path()
    if chrome:
        return {"EC_EXECUTABLE": chrome}

    kind, value, _ = browser_choice()
    if kind == "chromium" and value:
        return {"EC_EXECUTABLE": value}
    return {}


def check_common():
    """返回 (必需缺失, 可选缺失)。"""
    required, optional = [], []

    for name, hint in (
        ("requests", "python -m pip install requests"),
        ("openpyxl", "python -m pip install openpyxl"),
        ("PIL", "python -m pip install Pillow"),
        ("playwright", "python -m pip install playwright"),
    ):
        if not module_available(name):
            label = {"PIL": "Pillow"}.get(name, name)
            required.append(f"{label} 未安装 · {hint}")

    kind, _, warning = browser_choice()
    if not kind:
        required.append(warning)
    elif kind == "chromium":
        optional.append(warning)

    if not module_available("curl_cffi"):
        optional.append(
            "curl_cffi 未安装 · python -m pip install curl_cffi （装了 TLS 指纹才像 Chrome）")

    return required, optional


def check_platform(platform):
    """返回该平台缺失的文件列表。"""
    missing = []
    sub, script, extras = PLATFORM_FILES[platform]
    folder = os.path.join(VENDOR, sub)

    if not os.path.isfile(os.path.join(folder, "config.py")):
        missing.append(f"vendor/{sub}/config.py")

    if not os.path.isfile(os.path.join(folder, script)):
        missing.append(f"vendor/{sub}/{script}")

    for extra in extras:
        if not os.path.isfile(os.path.join(folder, extra)):
            missing.append(f"vendor/{sub}/{extra}")

    return missing


def profile_dir(platform):
    return os.path.join(PROFILES, platform)


def login_state(platform, harvest=None):
    """检查登录态。

    harvest 是一个可调用对象；传了就实际启动一次浏览器读 cookie。
    返回 (是否登录, cookie 数量, 命中的身份 cookie 名)。
    """
    markers = LOGIN_MARKERS.get(platform, ())
    if harvest is None:
        return False, 0, []

    try:
        cookies = harvest(platform)
    except Exception:
        return False, 0, []

    names = {c.get("name") for c in cookies if c.get("name")}
    hit = [m for m in markers if m in names]
    if platform in ("taobao", "jingdong"):
        from manual_login import has_identity
        cookie_text = "; ".join(f"{c.get('name')}={c.get('value')}" for c in cookies
                                if c.get("name") and c.get("value"))
        logged = has_identity(platform, cookie_text)
    else:
        logged = bool(hit)
    return logged, len(cookies), hit
