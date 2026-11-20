"""三个平台的调用封装。

每个平台都是「子进程 + 环境变量」的方式调起 vendor 里的原始脚本：
    EC_SEARCH  关键词
    EC_LIMIT   商品数量（0 = 全部）
    EC_PROXY   代理
    EC_PROFILE 登录目录
    EC_OUTPUT  Excel 输出目录

这样原脚本既保留了单独运行的能力，又不用为统一入口改写核心逻辑。
"""

import os
import subprocess
import sys

import sys as _sys, os as _os
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))

from common import paths

BASE_DIR = paths.ROOT
VENDOR = paths.VENDOR

# 平台 → (vendor 子目录, 主脚本, 支持的关键词/数量级参数)
SCRIPTS = {
    "jingdong": ("jingdong", "jd_sku.py"),
    "taobao": ("taobao", "xuanzeshuliang.py"),
    "xianyu": ("xianyu", "xianyu.py"),
}


def script_path(platform):
    sub, script = SCRIPTS[platform]
    return os.path.join(VENDOR, sub, script)


def build_env(platform, keyword=None, limit=None, proxy=None,
              profile_dir=None, output_dir=None, pages=None):
    env = dict(os.environ)

    # 系统没装 Chrome 时改用 playwright 自带内核
    try:
        from common import env_check
        env.update(env_check.browser_env())
    except Exception:
        pass
    env["PYTHONIOENCODING"] = "utf-8"
    env["PYTHONUTF8"] = "1"

    if keyword:
        env["EC_SEARCH"] = keyword
    if limit is not None:
        env["EC_LIMIT"] = str(limit)
    if proxy is not None:
        env["EC_PROXY"] = proxy
    if profile_dir:
        env["EC_PROFILE"] = profile_dir
    if output_dir:
        env["EC_OUTPUT"] = output_dir
    if pages is not None:
        env["EC_PAGES"] = str(pages)

    return env


def run(platform, keyword=None, limit=None, proxy=None,
        profile_dir=None, output_dir=None, pages=None, extra_args=None):
    """在 vendor 目录里跑对应脚本，把控制台交还给子进程（登录要交互）。

    返回子进程退出码。
    """
    path = script_path(platform)
    if not os.path.isfile(path):
        print(f"找不到脚本：{path}")
        return 1

    env = build_env(platform, keyword=keyword, limit=limit, proxy=proxy,
                    profile_dir=profile_dir, output_dir=output_dir, pages=pages)

    command = [sys.executable, os.path.basename(path)]
    command.extend(extra_args or [])

    try:
        completed = subprocess.run(
            command,
            cwd=os.path.dirname(path),
            env=env,
            check=False,
        )
        return completed.returncode
    except KeyboardInterrupt:
        print("\n已中断")
        return 130
    except OSError as error:
        print(f"启动失败：{error}")
        return 1
