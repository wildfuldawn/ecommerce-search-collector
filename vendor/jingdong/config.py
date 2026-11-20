"""京东采集配置（统一入口会通过环境变量覆盖这些值）。

Cookie 不用手工填：程序会从 profiles/jingdong 里的 Chrome 登录态自动读取。
也可以把整段 Cookie 填到 COOKIE 里强制指定。
"""

SEARCH = "纸巾"
PAGE_COUNT = 1
LIMIT = 0

# 留空 = 从登录目录自动取；填了就优先用这里的
COOKIE = r"""
"""

# 代理：留空直连。不要填海外机房/VPN。
PROXY = ""

SEARCH_MIN_INTERVAL = 3.0
SEARCH_MAX_INTERVAL = 6.0
DETAIL_MIN_INTERVAL = 4.0
DETAIL_MAX_INTERVAL = 8.0
COOLDOWN_SECONDS = 180


def validate():
    problems = []
    if not SEARCH.strip():
        problems.append("SEARCH 是空的")
    if PAGE_COUNT < 1:
        problems.append("PAGE_COUNT 必须 >= 1")
    if LIMIT < 0:
        problems.append("LIMIT 不能是负数")
    if SEARCH_MIN_INTERVAL > SEARCH_MAX_INTERVAL:
        problems.append("SEARCH_MIN_INTERVAL 不能大于 SEARCH_MAX_INTERVAL")
    if DETAIL_MIN_INTERVAL > DETAIL_MAX_INTERVAL:
        problems.append("DETAIL_MIN_INTERVAL 不能大于 DETAIL_MAX_INTERVAL")
    return problems
