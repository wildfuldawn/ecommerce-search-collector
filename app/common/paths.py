"""统一的路径解析。

目录结构（一级目录只放要双击的 bat）：

    电商采集/
    ├── 1-安装环境.bat
    ├── 2-开始采集.bat
    ├── 3-安装浏览器.bat
    ├── app/          ← 代码（启动.py + common + platforms）
    ├── docs/         ← 使用说明
    ├── tools/        ← 打包时用的下载脚本
    ├── vendor/       ← 采集脚本 + 离线安装包
    ├── profiles/     ← 各平台登录态
    └── output/       ← 采集结果

这里不写死层级，而是向上找带 app/ 和 vendor/ 的那一层，
以后调整目录也不用改代码。
"""

import os


def find_root(start=None):
    """从 start 往上找项目根目录（同时含 app/ 和 vendor/ 的那层）。"""
    current = os.path.abspath(start or os.path.dirname(os.path.abspath(__file__)))

    while True:
        has_app = os.path.isdir(os.path.join(current, "app"))
        has_vendor = os.path.isdir(os.path.join(current, "vendor"))
        if has_app and has_vendor:
            return current

        parent = os.path.dirname(current)
        if parent == current:
            # 到顶都没找到，退回本文件所在目录的上两级
            return os.path.abspath(
                os.path.join(os.path.dirname(os.path.abspath(__file__)), "..", ".."))
        current = parent


ROOT = find_root()
APP_DIR = os.path.join(ROOT, "app")
VENDOR = os.path.join(ROOT, "vendor")
PROFILES = os.path.join(ROOT, "profiles")
OUTPUT = os.path.join(ROOT, "output")
DOCS = os.path.join(ROOT, "docs")
TOOLS = os.path.join(ROOT, "tools")
SETTINGS_PATH = os.path.join(APP_DIR, "settings.json")


def ensure_dirs():
    for path in (PROFILES, OUTPUT):
        os.makedirs(path, exist_ok=True)
