"""电商采集工具 —— 统一入口。

用法：
    python app\启动.py          （正常情况双击 2-开始采集.bat 即可）

菜单里选平台 → 填关键词和数量 → 自动打开 Chrome（没登录的话扫码一次）
→ 结果存到 output/ 目录。
"""

import os
import sys

APP_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, APP_DIR)

from common import env_check, net, paths, settings, ui
from common.env_check import PLATFORM_NAMES
import platforms

BASE_DIR = paths.ROOT

PLATFORM_HINTS = {
    "jingdong": "只采集搜索列表，按目标数量自动翻页，含主图",
    "taobao": "只采集搜索列表，按目标数量自动翻页，含主图",
    "xianyu": "只采集搜索列表，按目标数量自动翻页，含主图",
}


# ------------------------------------------------------------
# 登录态检查
# ------------------------------------------------------------

def harvest(platform):
    """启动一次浏览器（不开网页）读 cookie。"""
    import importlib.util

    sub, script = platforms.SCRIPTS[platform]
    path = os.path.join(platforms.VENDOR, sub, script)
    if not os.path.isfile(path):
        return []

    # 只读 cookie，不需要跑脚本，直接用 playwright
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return []

    profile = env_check.profile_dir(platform)
    os.makedirs(profile, exist_ok=True)

    kind, executable, _ = env_check.browser_choice()
    options = {"headless": True,
               "args": ["--disable-blink-features=AutomationControlled",
                        "--mute-audio"]}
    if kind == "chromium" and executable:
        options["executable_path"] = executable
    else:
        options["channel"] = "chrome"

    try:
        with sync_playwright() as pw:
            context = pw.chromium.launch_persistent_context(profile, **options)
            try:
                return context.cookies()
            finally:
                try:
                    context.close()
                except Exception:
                    pass
    except Exception:
        return []


def check_login(platform, verbose=True):
    logged, total, hit = env_check.login_state(platform, harvest)
    if verbose:
        name = PLATFORM_NAMES[platform]
        if logged:
            print(f"登录态：已保存（{'、'.join(hit)}）")
        else:
            print(f"登录态：未检测到")
            print(f"  本次运行会打开 Chrome，请在里面登录 {name}，")
            print(f"  登录后不用管，程序会接着跑。登录态会保存到")
            print(f"  profiles/{platform}/，以后不用再登。")
    return logged


# ------------------------------------------------------------
# 跑一个平台
# ------------------------------------------------------------

def _snapshot_output():
    """记录 output 目录当前的文件集合，用来判断这次到底有没有产出。"""
    try:
        return set(os.listdir(env_check.OUTPUT))
    except OSError:
        return set()


def run_platform(platform, app_settings):
    name = PLATFORM_NAMES[platform]

    ui.title(f"{name} 采集")

    # 平台文件和全局依赖一起查：缺浏览器/缺库就不要假装能跑
    blocking = list(env_check.check_platform(platform))
    required, optional = env_check.check_common()
    blocking.extend(required)

    if optional:
        print("提示：")
        for item in optional:
            for line_no, line in enumerate(item.splitlines()):
                print(f"  {line if line_no else '· ' + line}")
        print()

    if blocking:
        print("这个平台还不能跑，缺：")
        for item in blocking:
            for line_no, line in enumerate(item.splitlines()):
                print(f"  {line if line_no else '· ' + line}")
        print()
        print("先双击 1-安装环境.bat 补齐，再回来试。")
        ui.pause()
        return

    saved = settings.platform_settings(app_settings, platform)

    keyword = ui.ask("搜索关键词", saved.get("keyword", ""))
    if not keyword:
        print("关键词不能为空")
        ui.pause()
        return

    count = ui.ask_int("要采集多少个商品（自动翻页补足）", max(1, int(saved.get("count", 10))), 1)

    print(f"登录目录：{env_check.profile_dir(platform)}")
    print("将在采集浏览器中直接复用并检查会话。")

    proxy = (app_settings.get("proxy") or "").strip()
    if proxy:
        print(f"代理：{net.mask_proxy(proxy)}")

    settings.update_platform(app_settings, platform,
                             keyword=keyword, count=count)
    app_settings["last_platform"] = platform
    settings.save(app_settings)

    print()
    print(f"关键词：{keyword}　目标数量：{count}（自动翻页，最多 100 页）")
    print()

    before = _snapshot_output()

    code = platforms.run(
        platform,
        keyword=keyword,
        limit=count,
        proxy=proxy,
        profile_dir=env_check.profile_dir(platform),
        output_dir=env_check.OUTPUT,
    )

    after = _snapshot_output()
    new_files = sorted(after - before)

    print()
    print("=" * 72)

    if new_files:
        print(f"{name} {'采集完成' if code == 0 else '部分采集或提前停止'}，生成 {len(new_files)} 个文件：")
        for filename in new_files:
            full = os.path.join(env_check.OUTPUT, filename)
            try:
                size_kb = os.path.getsize(full) / 1024
                print(f"  {full}  （{size_kb:.1f} KB）")
            except OSError:
                print(f"  {full}")
        print()
        print(f"目录：{env_check.OUTPUT}")
    elif code != 0:
        print(f"{name} 采集失败（退出代码 {code}）。")
        print("原因见上面的输出。")
    else:
        print(f"{name} 这次没有生成文件 —— 数据没采集到。")
        print()
        print("常见原因（按可能性排序）：")
        print("  1. 没有登录态：第一次跑要在弹出的 Chrome 里登录一次")
        print("  2. 搜索没有结果：关键词太生僻，或者被限流")
        print("  3. 浏览器启动失败：确认 Chrome 能用，或先跑 1-安装环境.bat")
        print()
        print("往上翻，找红色报错或带「失败 / 未检测 / 被拦」的那几行。")

    print("=" * 72)

    ui.pause()


# ------------------------------------------------------------
# 设置
# ------------------------------------------------------------

def show_settings(app_settings):
    ui.title("当前设置")

    proxy = (app_settings.get("proxy") or "").strip()
    print(f"代理：{net.mask_proxy(proxy)}")
    if proxy:
        net.print_exit_report(proxy)
    else:
        print("  （直连本机网络，推荐）")

    print()
    print("各平台上次用的参数：")
    for platform, label in PLATFORM_NAMES.items():
        item = settings.platform_settings(app_settings, platform)
        print(f"  {label}：关键词 {item.get('keyword', '-')}"
              f"，目标数量 {item.get('count', '-')}，在结果页点击下一页")

    print()
    print(f"登录目录：{env_check.PROFILES}")
    print(f"结果目录：{env_check.OUTPUT}")
    ui.pause()


def edit_settings(app_settings):
    ui.title("修改设置")

    proxy = ui.ask("代理地址（留空 = 直连；不要填海外机房）",
                   app_settings.get("proxy") or "")
    app_settings["proxy"] = proxy
    settings.save(app_settings)

    if proxy:
        print()
        net.print_exit_report(proxy)
    else:
        print("已设为直连")

    ui.pause()


def run_env_check(interactive=True):
    ui.title("环境自检")

    required, optional = env_check.check_common()

    kind, executable, warning = env_check.browser_choice()
    if kind == "chrome":
        print(f"浏览器：系统 Chrome（{env_check.chrome_path()}）")
    elif kind == "chromium":
        print(f"浏览器：playwright 自带内核（{executable}）")
        print("  ⚠ 这不是 Google Chrome，指纹上更容易被识别，建议装官方 Chrome")
    else:
        print("浏览器：未找到")
        if warning:
            for line in warning.splitlines():
                print(f"  {line}")
    print(f"Python：{sys.version.split()[0]}")
    print()
    print("三个平台的依赖：")
    for platform, label in PLATFORM_NAMES.items():
        missing = env_check.check_platform(platform)
        status = "就绪" if not missing else "缺 " + "、".join(missing)
        print(f"  {label}：{status}")

    print()
    if required:
        print("必须先解决：")
        for item in required:
            print(f"  · {item}")
    else:
        print("必需依赖齐全。")

    if optional:
        print()
        print("建议安装（降低风控）：")
        for item in optional:
            print(f"  · {item}")

    if interactive:
        ui.pause()


def check_only():
    """给 1-安装环境.bat 用：只做自检，不做任何交互。"""
    env_check.ensure_dirs()
    run_env_check(interactive=False)
    required, _ = env_check.check_common()
    missing_files = {}
    for platform in ("jingdong", "taobao", "xianyu"):
        broken = env_check.check_platform(platform)
        if broken:
            missing_files[platform] = broken

    print()
    if required or missing_files:
        print("还有问题需要处理（见上）。")
        return 1
    print("环境就绪，可以运行 2-开始采集.bat 开始采集。")
    return 0


def relogin(app_settings):
    """清掉某个平台的登录态，让它下次重新登录。"""
    ui.title("重新登录")

    options = [(key, PLATFORM_NAMES[key], "") for key in PLATFORM_NAMES]
    platform = ui.choose("要清掉哪个平台的登录态？", options)
    name = PLATFORM_NAMES[platform]

    folder = env_check.profile_dir(platform)
    if not os.path.isdir(folder):
        print(f"{name} 还没有登录目录，直接运行即可")
        ui.pause()
        return

    if not ui.ask_yes_no(f"确认清空 {name} 的登录态？", default=False):
        print("已取消")
        ui.pause()
        return

    import shutil
    shutil.rmtree(folder, ignore_errors=True)
    if os.path.isdir(folder):
        print("删除失败，可能有 Chrome 进程占用，请先关闭所有 Chrome 窗口")
    else:
        print(f"{name} 的登录态已清空，下次运行请重新登录")
    ui.pause()


# ------------------------------------------------------------
# 主菜单
# ------------------------------------------------------------

def main():
    ui.setup_console()
    env_check.ensure_dirs()

    app_settings = settings.load()
    settings.save(app_settings)          # 首次运行就把 settings.json 建出来
    last = app_settings.get("last_platform", "xianyu")

    while True:
        ui.title("电商采集工具")

        options = [(platform, PLATFORM_NAMES[platform], PLATFORM_HINTS[platform])
                   for platform in ("jingdong", "taobao", "xianyu")]
        options += [
            ("__settings__", "查看 / 修改设置", "代理、上次参数、目录位置"),
            ("__env__", "环境自检", "检查依赖和 Chrome"),
            ("__relogin__", "重新登录", "清掉某个平台的登录态"),
            ("__quit__", "退出", ""),
        ]

        default_index = list(PLATFORM_NAMES).index(last) if last in PLATFORM_NAMES else 0
        choice = ui.choose("请选择要采集的平台：", options, default_index)

        if choice == "__quit__":
            print("再见")
            return

        if choice == "__settings__":
            show_settings(app_settings)
            continue

        if choice == "__env__":
            run_env_check()
            continue

        if choice == "__relogin__":
            relogin(app_settings)
            continue

        run_platform(choice, app_settings)


if __name__ == "__main__":
    if "--check" in sys.argv:
        ui.setup_console()
        raise SystemExit(check_only())

    try:
        main()
    except KeyboardInterrupt:
        print("\n已退出")
