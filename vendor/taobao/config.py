"""淘宝采集配置（统一入口会通过环境变量覆盖这些值）。

Cookie 不用手工填：程序会从 profiles/taobao 里的 Chrome 登录态自动读取。
"""

SEARCH = "纸巾"

# 搜索页数（每页 10 个商品）
PAGES = 1

COOKIE = r"""
"""

PROXY = ""

MIN_INTERVAL = 4.0
MAX_INTERVAL = 8.0
COOLDOWN_SECONDS = 180


def validate():
    problems = []
    if not SEARCH.strip():
        problems.append("SEARCH 是空的")
    if PAGES < 1:
        problems.append("PAGES 必须 >= 1")
    if MIN_INTERVAL > MAX_INTERVAL:
        problems.append("MIN_INTERVAL 不能大于 MAX_INTERVAL")
    return problems
