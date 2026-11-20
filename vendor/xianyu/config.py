"""闲鱼采集配置（统一入口会通过环境变量覆盖这些值）。

闲鱼不需要填 Cookie：登录态存在 profiles/xianyu，程序直接复用。
"""

# 登录目录。统一入口会通过 EC_PROFILE 覆盖成 <项目>/profiles/xianyu；
# 单独运行本脚本时用下面这个（指向项目根的 profiles，不污染 vendor 目录）。
import os as _os
_RUN_ROOT = _os.path.dirname(_os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
USER_DATA_DIR = _os.path.join(_RUN_ROOT, "profiles", "xianyu")

SEARCH = "机械键盘"
PAGES = 1
LIMIT = 0

# 闲鱼只有主图相册，没有详情图
ENRICH_DETAIL = False
EMBED_IMAGES = True
IMAGE_SIZE = 400
MAX_MAIN_IMAGES = 0
DOWNLOAD_IMAGES = False

PROXY = ""

MIN_INTERVAL = 3.0
MAX_INTERVAL = 6.0
COOLDOWN_SECONDS = 180

DEBUG_DUMP = False
FORCE_LOGIN = False
USE_API = False
API_FALLBACK_BROWSER = True
TLS_IMPERSONATE = "chrome"


def validate():
    problems = []
    if not SEARCH.strip():
        problems.append("SEARCH 是空的")
    if PAGES < 1:
        problems.append("PAGES 必须 >= 1")
    if LIMIT < 0:
        problems.append("LIMIT 不能是负数")
    if MIN_INTERVAL > MAX_INTERVAL:
        problems.append("MIN_INTERVAL 不能大于 MAX_INTERVAL")
    if DOWNLOAD_IMAGES and not ENRICH_DETAIL:
        problems.append("DOWNLOAD_IMAGES=True 需要 ENRICH_DETAIL=True")
    return problems
