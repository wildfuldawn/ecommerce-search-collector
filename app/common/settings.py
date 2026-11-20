"""记住上次用的参数（关键词、数量、代理），存在 settings.json。

分享给别人时这个文件不存在也没关系，程序会用默认值并重新生成。
"""

import json
import os

from . import paths

BASE_DIR = paths.ROOT
SETTINGS_PATH = paths.SETTINGS_PATH

DEFAULTS = {
    "proxy": "",
    "last_platform": "xianyu",
    "platforms": {
        "jingdong": {"keyword": "纸巾", "count": 10, "pages": 1},
        "taobao": {"keyword": "纸巾", "count": 10},
        "xianyu": {"keyword": "机械键盘", "count": 10, "pages": 1},
    },
}


def _deep_merge(base, extra):
    result = dict(base)
    for key, value in (extra or {}).items():
        if isinstance(value, dict) and isinstance(result.get(key), dict):
            result[key] = _deep_merge(result[key], value)
        else:
            result[key] = value
    return result


def load():
    """读配置；文件损坏或不存在都回到默认值。"""
    data = {}
    if os.path.isfile(SETTINGS_PATH):
        try:
            with open(SETTINGS_PATH, "r", encoding="utf-8") as handle:
                data = json.load(handle)
        except (OSError, json.JSONDecodeError, ValueError):
            data = {}
    return _deep_merge(DEFAULTS, data if isinstance(data, dict) else {})


def _clean(value):
    """把非法字符（比如粘贴乱码产生的代理项）替换掉，避免写文件时崩。"""
    if isinstance(value, str):
        return value.encode("utf-8", "replace").decode("utf-8")
    if isinstance(value, dict):
        return {k: _clean(v) for k, v in value.items()}
    if isinstance(value, list):
        return [_clean(v) for v in value]
    return value


def save(settings):
    try:
        with open(SETTINGS_PATH, "w", encoding="utf-8", errors="replace") as handle:
            json.dump(_clean(settings), handle, ensure_ascii=False, indent=2)
        return True
    except (OSError, TypeError, ValueError):
        return False


def platform_settings(settings, platform):
    return dict((settings.get("platforms") or {}).get(platform) or {})


def update_platform(settings, platform, **values):
    settings.setdefault("platforms", {})[platform] = {
        **platform_settings(settings, platform), **values
    }
    return settings
