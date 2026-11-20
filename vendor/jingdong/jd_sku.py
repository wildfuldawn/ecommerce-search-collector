"""京东商品 + SKU 信息采集。

搜索：走 api.m.jd.com 的 pc_search_searchWare 接口，h5st 由同目录 test_fixed.js 生成。
详情：不直接调详情接口（会被判 605），改用真实 Chrome 打开商品页，
      把页面自己发的请求接住再解析。两级策略：
        1) PC 详情页 → 截获 pc_detailpage_wareBusiness，能拿到 SKU + 价格 + 库存
        2) PC 被频控 / 拿不到时 → 移动页 item.m.jd.com，解析内嵌 JSON，能拿到 SKU 列表
      签名和风控都交给浏览器，不用自己算 h5st。

Cookie 与搜索词在同目录 config.py 里改，本文件不用动。
运行：python jd_sku.py
"""

import argparse
import html as html_lib
import json
import logging
import os
import random
import re
import secrets
import subprocess
import sys
import time
from urllib.parse import quote, unquote, urlparse

try:
    import requests
except ImportError:
    raise SystemExit(
        "缺少 requests 库（不是标准库，需要单独安装）\n"
        "  处理：双击 1-安装环境.bat，或执行\n"
        "        python -m pip install requests"
    )

try:
    from curl_cffi import requests as curl_requests
    CURL_CFFI_OK = True
except ImportError:
    curl_requests = None
    CURL_CFFI_OK = False

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("缺少依赖，请先运行：python -m pip install openpyxl")

try:
    from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError
except ImportError:
    raise SystemExit("缺少依赖，请先运行：python -m pip install playwright")


# ============================================================
# 配置
# ============================================================

try:
    from config import (COOKIE, SEARCH, PAGE_COUNT, LIMIT,
                        PROXY, SEARCH_MIN_INTERVAL, SEARCH_MAX_INTERVAL,
                        DETAIL_MIN_INTERVAL, DETAIL_MAX_INTERVAL, COOLDOWN_SECONDS,
                        validate as validate_config)
except ImportError as error:
    raise SystemExit(
        f"缺少同目录 config.py，或其中缺少 COOKIE / SEARCH：{error}\n"
        "请确保 config.py 和本文件放在同一个文件夹里。"
    ) from None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# 浏览器持久化目录：存京东登录态，不是采集结果
USER_DATA_DIR = os.path.join(BASE_DIR, "jd_browser")

# 采集结果直接存到脚本所在目录
EXCEL_DIR = BASE_DIR
EXCEL_NAME = "jd_result.xlsx"

SEARCH_SIGNER = os.path.join(BASE_DIR, "test_fixed.js")
SEARCH_APP_ID = "f06cc"

# 两次详情请求之间的随机间隔（秒）；太快会触发 PC 频控
MIN_INTERVAL = DETAIL_MIN_INTERVAL
MAX_INTERVAL = DETAIL_MAX_INTERVAL

PC_UA = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36"
)
EXCEL_HEADERS = [
    "商品ID", "商品标题", "商品销量", "SPU ID", "店铺名称", "卖家ID", "店铺ID",
    "商品链接", "SKU优惠前价格", "SKU价格", "SKU ID", "SKU规格", "SKU属性",
    "SKU库存", "SKU状态", "数据来源",
]
EXCEL_WIDTHS = [15, 45, 12, 12, 22, 14, 14, 42, 15, 11, 15, 40, 12, 10, 12, 14]


# ============================================================
# 统一入口（电商采集/启动.py）用的覆盖层
# 不改这里也能单独运行本脚本；环境变量存在时优先级更高。
# ============================================================
def _env(name, default=None):
    value = os.environ.get(name)
    return value if value not in (None, "") else default


SEARCH = _env("EC_SEARCH", SEARCH)
PROXY = _env("EC_PROXY", PROXY)
USER_DATA_DIR = _env("EC_PROFILE", USER_DATA_DIR)
EXCEL_DIR = _env("EC_OUTPUT", EXCEL_DIR)
try:
    _limit_env = _env("EC_LIMIT")
    if _limit_env is not None:
        LIMIT = int(_limit_env)
except (TypeError, ValueError):
    pass

try:
    _pages_env = _env("EC_PAGES")
    if _pages_env is not None:
        PAGE_COUNT = max(1, int(_pages_env))
except (TypeError, ValueError):
    pass

LOG = logging.getLogger("jd_sku")


# ============================================================
# 网络出口（代理 / TLS 指纹）
# ============================================================

def parse_proxy(text=None):
    """把 PROXY 拆成 Playwright 需要的结构；空串返回 None。"""
    text = (PROXY if text is None else text or "").strip()
    if not text:
        return None
    if "://" not in text:
        text = "http://" + text

    parsed = urlparse(text)
    if not parsed.hostname:
        return None

    server = f"{parsed.scheme}://{parsed.hostname}"
    if parsed.port:
        server += f":{parsed.port}"

    config = {"server": server}
    if parsed.username:
        config["username"] = unquote(parsed.username)
    if parsed.password:
        config["password"] = unquote(parsed.password)
    return config


def proxy_urls(text=None):
    text = (PROXY if text is None else text or "").strip()
    if not text:
        return None
    if "://" not in text:
        text = "http://" + text
    return {"http": text, "https": text}


def mask_proxy(text=None):
    text = (PROXY if text is None else text or "").strip()
    if not text:
        return "(直连)"
    if "://" not in text:
        text = "http://" + text
    parsed = urlparse(text)
    host = parsed.hostname or "?"
    port = f":{parsed.port}" if parsed.port else ""
    auth = "带认证@" if parsed.username else ""
    return f"{parsed.scheme}://{auth}{host}{port}"


def new_session(impersonate="chrome"):
    """带 TLS 指纹伪装和代理的 HTTP 会话。

    普通 requests 的 JA3 指纹是 Python 特征，京东边缘 WAF 很容易识别；
    curl_cffi 让握手像 Chrome。没装就退回 requests（功能不变，只是更易被拦）。
    """
    proxies = proxy_urls()

    if CURL_CFFI_OK:
        try:
            session = curl_requests.Session(impersonate=impersonate, proxies=proxies)
            session._tls_profile = f"curl_cffi({impersonate})"
            return session
        except TypeError:
            session = curl_requests.Session(impersonate=impersonate)
            if proxies:
                session.proxies.update(proxies)
            session._tls_profile = f"curl_cffi({impersonate})"
            return session
        except Exception:
            pass

    session = requests.Session()
    if proxies:
        session.proxies.update(proxies)
    session._tls_profile = "requests 默认（Python 特征）"
    return session


def tls_profile_of(session):
    return getattr(session, "_tls_profile", "未知")


def probe_exit_ip(timeout=20):
    """查出口 IP 详情，判断这个 IP 适不适合跑京东。"""
    session = new_session()
    url = ("http://ip-api.com/json/?fields=query,country,regionName,"
           "city,isp,org,hosting")
    try:
        response = session.get(url, timeout=timeout, verify=False)
        return response.json()
    except Exception:
        return None


def check_exit_ip(timeout=20):
    info = probe_exit_ip(timeout)
    if not info or not info.get("query"):
        return "(查询失败，代理可能不通)"
    return (f"{info.get('query')}  {info.get('country')} / "
            f"{info.get('isp') or info.get('org') or ''}").strip()


def warn_if_risky_ip(info):
    if not info or not info.get("query"):
        return

    problems = []
    if info.get("hosting"):
        problems.append("这是云服务器/机房 IP（hosting=True）")
    if str(info.get("country") or "").lower() != "china":
        problems.append(f"出口在海外（{info.get('country')}）")

    if not problems:
        return

    print()
    print("  ⚠ 这个出口 IP 不适合跑京东：")
    for item in problems:
        print(f"      · {item}")
    print("     机房 IP 和海外异地登录是京东风控的重点，")
    print("     用它们只会更容易 605 / 403。建议 PROXY 留空，")
    print("     或者用手机热点（运营商 IP 最接近真人）。")
    print()



def harvest_cookies(profile_dir, headless=True):
    """从 Chrome 登录目录读出 cookie，不打开任何网页（零网络请求）。"""
    if not os.path.isdir(profile_dir):
        return []

    if not CURL_CFFI_OK:
        pass  # 与 TLS 无关，这里只是保留结构

    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return []

    try:
        with sync_playwright() as pw:
            options = _apply_browser_env(dict(
                channel="chrome",
                headless=headless,
                args=["--disable-blink-features=AutomationControlled",
                      "--mute-audio"],
            ))
            context = pw.chromium.launch_persistent_context(profile_dir, **options)
            try:
                return context.cookies()
            finally:
                try:
                    context.close()
                except Exception:
                    pass
    except Exception:
        return []


def cookies_to_text(cookies):
    return "; ".join(
        f"{c.get('name')}={c.get('value')}"
        for c in cookies
        if c.get("name") and c.get("value")
    )


# 搜索接口共用一个会话：TLS 指纹像 Chrome，并按需走代理
HTTP = new_session()


class JdBlocked(RuntimeError):
    """京东要求验证或明确拒绝；继续只会扩大风控范围。"""


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def clean_text(value):
    if value is None:
        return ""
    text = str(value)
    if re.search(r"</?[A-Za-z][^>]*>", text):
        text = re.sub(r"<[^>]+>", "", text)
    return re.sub(r"\s+", " ", html_lib.unescape(text)).strip()


def number(value, positive=False):
    if isinstance(value, bool) or value is None:
        return None
    text = str(value).strip()
    if not re.fullmatch(r"\d+(?:\.\d+)?", text):
        return None
    result = float(text)
    if result < 0 or (positive and result <= 0):
        return None
    return result


# ============================================================
# 搜索接口（h5st 签名 + API）
# ============================================================

def sign(params, page_url, cookie_text):
    if not os.path.isfile(SEARCH_SIGNER):
        raise SystemExit(f"找不到签名器：{SEARCH_SIGNER}")

    payload = {
        "params": params,
        "appId": SEARCH_APP_ID,
        "url": page_url,
        "userAgent": PC_UA,
        "languages": ["zh-CN", "zh"],
        "cookie": cookie_text,
        "browserProfile": {
            "canvas": "d292170ca009238e7f6c027ac97f13fa",
            "webgl": "41c7496853e79e15786704fb6145732c",
            "pFlag": "mVVVAmxELq",
            "behavior": "F2",
        },
    }

    completed = subprocess.run(
        ["node", SEARCH_SIGNER],
        input=json.dumps(payload, ensure_ascii=False),
        capture_output=True, text=True, encoding="utf-8", timeout=60, check=False,
    )

    if completed.returncode:
        lines = (completed.stderr or "").strip().splitlines()
        raise SystemExit(
            "签名失败（确认 Node.js 在 PATH 中）：" + (lines[-1] if lines else "")
        )

    try:
        signed = json.loads(completed.stdout)
    except ValueError:
        raise SystemExit("签名器输出不是 JSON") from None

    h5st = signed.get("h5st", "")
    pieces = h5st.split(";")

    if len(pieces) != 10 or pieces[2] != SEARCH_APP_ID:
        raise SystemExit(
            f"签名结构异常：期望 10 段且第 3 段为 {SEARCH_APP_ID}，实际 {len(pieces)} 段"
        )

    params["h5st"] = h5st
    return params


def explain_failure(data, api_name):
    code = str(data.get("code", "0"))

    if code == "605":
        disposal = data.get("disposal") or {}
        ev_api = disposal.get("evApi")
        if not ev_api:
            try:
                ev_api = (json.loads(disposal.get("evContent") or "{}")).get("evApi")
            except ValueError:
                ev_api = None
        return (
            f"{api_name}：code=605，京东要求先验证身份。\n"
            f"  风控接口：{ev_api or '未知'}\n"
            "  程序会打开京东页面等待你手动登录/验证，再重试当前搜索页。"
        )

    if code != "0":
        return f"{api_name}：code={code}，{data.get('echo') or data.get('msg') or ''}"

    return None


def search_page(keyword, page_index, cookie_text, cookie_map):
    pvid = secrets.token_hex(16)
    referer = (
        f"https://search.jd.com/Search?keyword={quote(keyword)}&enc=utf-8"
        f"&pvid={pvid}&sourcePage=myjd&from=myjd"
    )

    body = json.dumps({
        "enc": "utf-8", "pvid": pvid, "from": "myjd", "page": page_index,
        "mode": "", "concise": False, "hoverPictures": True, "newAdvRepeat": False,
        "mixerParam": False, "new_interval": True, "s": 1, "pageSize": 30,
    }, ensure_ascii=False, separators=(",", ":"))

    params = {
        "keyword": keyword,
        "appid": "search-pc-java",
        "client": "pc",
        "clientVersion": "1.0.0",
        "cthr": "1",
        "uuid": cookie_map.get("__jdu") or cookie_map.get("mba_muid") or "",
        "loginType": "3",
        "functionId": "pc_search_searchWare",
        "body": body,
        "t": str(time.time_ns() // 1_000_000),
    }
    sign(params, referer, cookie_text)

    headers = {
        "Accept": "*/*",
        "Accept-Language": "zh-CN,zh;q=0.9",
        "Origin": "https://search.jd.com",
        "Referer": referer,
        "Sec-Fetch-Dest": "empty",
        "Sec-Fetch-Mode": "cors",
        "Sec-Fetch-Site": "same-site",
        "User-Agent": PC_UA,
        "X-Referer-Page": "https://search.jd.com/Search",
        "x-api-eid-token": cookie_map.get("3AB9D23F7A4B3CSS", ""),
    }

    # 用带 TLS 指纹伪装（+ 可选代理）的 session，而不是裸 requests
    response = HTTP.get(
        "https://api.m.jd.com/api", params=params, headers=headers,
        cookies=cookie_map, timeout=30, verify=False,
    )

    if response.status_code == 403 and not response.content:
        raise JdBlocked(
            "搜索接口 HTTP 403 且响应 0 字节：被京东边缘 WAF 拦下。\n"
            "  两种情况：\n"
            "    1) 请求太频繁被临时拦（等 3~5 分钟，或调大间隔）\n"
            "    2) Cookie 不完整（检查是否含 __jdu 和 3AB9D23F7A4B3CSS）\n"
            "  另外：海外机房 IP 会显著提高 403 概率。"
        )

    if response.status_code != 200:
        raise JdBlocked(f"搜索接口 HTTP {response.status_code}")

    try:
        data = response.json()
    except ValueError:
        raise JdBlocked("搜索接口返回的不是 JSON") from None

    problem = explain_failure(data, f"搜索第 {page_index} 页")

    if problem:
        raise JdBlocked(problem)

    return ((data.get("data") or {}).get("wareList")) or []


# ============================================================
# 详情：浏览器自动化
# ============================================================

def find_embedded_json(page_html, marker, must_have=()):
    """在 HTML 里找包含 marker 的 JSON 对象。

    京东把商品数据直接以 JSON 写进页面，同一个 marker 可能命中多个嵌套层级，
    这里从所有可解析的候选里挑字段最全的那个（优先含 must_have 中的键）。
    """
    decoder = json.JSONDecoder()
    needle = marker.strip('"')
    position = page_html.find(marker)

    if position < 0:
        return None

    candidates = []

    for start in range(position, max(0, position - 20000), -1):
        if page_html[start] != "{":
            continue
        try:
            obj, _ = decoder.raw_decode(page_html, start)
        except ValueError:
            continue
        if not isinstance(obj, dict) or needle not in obj:
            continue
        bonus = sum(10 for key in must_have if key in obj)
        candidates.append((bonus + len(obj), obj))

    if not candidates:
        return None

    candidates.sort(key=lambda pair: pair[0])
    return candidates[-1][1]


def normalize_from_ware_business(data, sku):
    """把 PC 详情接口响应整理成统一结构。"""
    config = data.get("pageConfigVO") or {}
    shop = data.get("itemShopInfo") or {}
    head = data.get("skuHeadVO") or {}
    options = data.get("colorSizeVO") or {}
    stock_map = options.get("stockInfo") or {}
    price = data.get("price") or {}
    stock = data.get("stockInfo") or {}

    found = {}
    for group in options.get("colorSizeList", []):
        title = clean_text(group.get("title")) or "规格"
        for button in group.get("buttons", []):
            for value in button.get("skuList") or [button.get("skuId")]:
                child = str(value or "")
                if not (child.isascii() and child.isdigit()):
                    continue
                row = found.setdefault(child, {"skuId": child, "attributes": {}})
                choices = row["attributes"].setdefault(title, [])
                label = clean_text(button.get("text"))
                if label and label not in choices:
                    choices.append(label)

    if sku not in found:
        found[sku] = {"skuId": sku, "attributes": {}}

    for child, row in found.items():
        row["jdPrice"] = None
        row["finalPrice"] = None
        row["availableNum"] = None
        row["stockState"] = (stock_map.get(child) or {}).get("stockState")

    if str(price.get("id", "")) == sku:
        found[sku]["jdPrice"] = price.get("p")
        found[sku]["finalPrice"] = (price.get("finalPrice") or {}).get("price")
    found[sku]["availableNum"] = (stock.get("stockInfo") or {}).get("availableNum")
    found[sku]["stockState"] = stock.get("stockState", found[sku]["stockState"])

    for child, row in found.items():
        row["spec"] = "；".join(" / ".join(values) for values in row["attributes"].values())

    return {
        "skuId": sku,
        "title": clean_text(head.get("skuTitle")),
        "venderId": clean_text(config.get("venderId")),
        "shopId": clean_text(shop.get("shopId") or config.get("shopId")),
        "shopName": clean_text(shop.get("shopName")),
        "brand": clean_text(price.get("brand") or (data.get("brandInfo") or {}).get("brandName")),
        "skus": list(found.values()),
        "source": "PC详情",
    }


def normalize_from_mobile(page_html, sku):
    """解析移动页内嵌的 newColorSize，得到 SKU 列表（价格需要登录，可能为空）。"""
    data = find_embedded_json(
        page_html, '"newColorSize"',
        must_have=("skuName", "saleProp", "salePropSeq"),
    )

    if not data:
        raise RuntimeError("移动页里没找到内嵌的 SKU 数据（newColorSize）")

    prop_names = data.get("saleProp") or {}
    skus = []

    for item in data.get("newColorSize") or []:
        child = str(item.get("skuId") or "")
        if not (child.isascii() and child.isdigit()):
            continue

        attributes = {}
        for seq, values in (data.get("salePropSeq") or {}).items():
            label = clean_text(prop_names.get(str(seq))) or f"规格{seq}"
            text = clean_text(item.get("color") if seq == "1" else item.get("size"))
            if not text:
                text = clean_text(item.get("1") if seq == "1" else "")
            if not text:
                text = clean_text((values or [""])[0])
            if text:
                attributes[label] = [text]

        if not attributes:
            single = clean_text(item.get("1") or item.get("color") or item.get("specName"))
            if single:
                attributes["规格"] = [single]

        skus.append({
            "skuId": child,
            "attributes": attributes,
            "jdPrice": None,
            "finalPrice": None,
            "availableNum": None,
            "stockState": clean_text(item.get("skuMark") or ""),
            "spec": "；".join(" / ".join(v) for v in attributes.values()),
        })

    if not skus:
        raise RuntimeError("移动页拿到了 SKU 结构但没解析出任何规格")

    return {
        "skuId": sku,
        "title": clean_text(data.get("skuName")),
        "venderId": clean_text(data.get("venderID")),
        "shopId": "",
        "shopName": "",
        "brand": clean_text(data.get("brandName")),
        "skus": skus,
        "source": "移动页SKU",
    }


class JdDetailBrowser:
    """真实 Chrome 打开商品页，截获页面自己的详情请求。"""

    FREQUENT_HOST = "pc-frequent-pro.pf.jd.com"
    LOGIN_HOSTS = ("passport.jd.com", "risk_handler", "cfe.m.jd.com")

    def __init__(self, playwright, headless=False):
        ensure_dir(USER_DATA_DIR)

        print()
        print("=" * 80)
        print("正在启动 Chrome")
        print("=" * 80)
        print("首次运行请在打开的窗口里登录京东，登录态会保存在 jd_browser 目录。")
        print("登录后能拿到 SKU 价格；不登录也能拿到 SKU 列表，只是没有价格。")

        # 只开一个持久化上下文：同一个 profile 不能同时被两个 Chrome 实例占用。
        # 移动页 item.m.jd.com 用 PC UA 访问同样返回内嵌 SKU 数据，所以共用一个上下文即可。
        launch_options = _apply_browser_env(dict(
            channel="chrome",
            headless=headless,
            user_agent=PC_UA,
            viewport={"width": 1440, "height": 900},
            locale="zh-CN",
            timezone_id="Asia/Shanghai",
            args=["--disable-blink-features=AutomationControlled", "--mute-audio"],
        ))
        proxy = parse_proxy()
        if proxy:
            launch_options["proxy"] = proxy
            print(f"网络出口：{mask_proxy()}")

        self.context = playwright.chromium.launch_persistent_context(
            USER_DATA_DIR, **launch_options)

        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        self.mobile_page = self.context.new_page()
        self.logged_in = False

    def close(self):
        try:
            self.context.close()
        except Exception:
            pass

    # --------------------------------------------------------
    # 状态判断
    # --------------------------------------------------------

    @staticmethod
    def build_pc_url(sku):
        return f"https://item.jd.com/{sku}.html"

    @staticmethod
    def build_mobile_url(sku):
        return f"https://item.m.jd.com/product/{sku}.html"

    def is_rate_limited(self):
        return self.FREQUENT_HOST in (self.page.url or "")

    def is_login_page(self):
        url = (self.page.url or "").lower()
        return any(host in url for host in self.LOGIN_HOSTS)

    # --------------------------------------------------------
    # 登录
    # --------------------------------------------------------

    def ensure_login(self, sku):
        if not self.is_login_page() and not self.is_rate_limited():
            return True

        print()
        print("=" * 80)
        print("检测到京东登录 / 验证页面")
        print("=" * 80)
        print()
        print("请在当前打开的 Chrome 里完成：")
        print("  1. 登录京东账号")
        print("  2. 如果弹出滑块/短信验证，也在这里过掉")
        print("  3. 完成后不要关闭 Chrome")

        input("\n完成后回到终端按 Enter 继续...")

        print("\n正在重新打开商品页...")
        try:
            self.page.goto(self.build_pc_url(sku), wait_until="domcontentloaded", timeout=60000)
        except Exception as error:
            print(f"重新打开出现异常：{type(error).__name__}")

        time.sleep(3)

        if self.is_login_page() or self.is_rate_limited():
            print("\n仍然卡在登录 / 验证 / 频控页，本次商品改用移动页兜底。")
            return False

        self.logged_in = True
        return True

    # --------------------------------------------------------
    # PC 详情：截获 wareBusiness
    # --------------------------------------------------------

    @staticmethod
    def _is_detail_api(response):
        return "api.m.jd.com" in response.url and "pc_detailpage_wareBusiness" in response.url

    def fetch_pc(self, sku, timeout=45000):
        """返回 wareBusiness 的 JSON；拿不到返回 None（由调用方降级）。"""
        url = self.build_pc_url(sku)
        response = None

        try:
            with self.page.expect_response(self._is_detail_api, timeout=timeout) as info:
                try:
                    self.page.goto(url, wait_until="domcontentloaded", timeout=60000)
                except PlaywrightTimeoutError:
                    pass
            response = info.value
        except PlaywrightTimeoutError:
            response = None

        if response is None:
            return None

        if response.status != 200:
            return None

        try:
            data = response.json()
        except Exception:
            return None

        code = str(data.get("code", "0"))

        if code == "605":
            if self.ensure_login(sku):
                return self.fetch_pc(sku, timeout)
            raise JdBlocked("详情接口 code=605：登录态失效，请在 Chrome 里重新登录京东")

        if code != "0":
            return None

        if not str((data.get("pageConfigVO") or {}).get("skuid", "")):
            return None

        return data

    # --------------------------------------------------------
    # 移动页兜底
    # --------------------------------------------------------

    def fetch_mobile(self, sku):
        url = self.build_mobile_url(sku)
        try:
            self.mobile_page.goto(url, wait_until="domcontentloaded", timeout=60000)
        except PlaywrightTimeoutError:
            pass
        time.sleep(2)
        return self.mobile_page.content()

    # --------------------------------------------------------
    # 统一入口
    # --------------------------------------------------------

    def fetch(self, sku):
        """返回 (整理后的详情 dict, 备注)。"""
        data = self.fetch_pc(sku)

        if data is not None:
            return normalize_from_ware_business(data, sku), ""

        # PC 拿不到（频控 / 结构异常）→ 移动页兜底
        note = "PC 详情未拿到（频控或风控），已改用移动页（无价格）"
        page_html = self.fetch_mobile(sku)
        return normalize_from_mobile(page_html, sku), note


# ============================================================
# 组装 Excel 行
# ============================================================

def build_rows(parent, detail):
    sku = str(detail["skuId"])
    rows = []

    for variant in detail["skus"]:
        jd_price = number(variant.get("jdPrice"), positive=True)
        final_price = number(variant.get("finalPrice"), positive=True)

        rows.append([
            sku,
            detail.get("title") or clean_text(parent.get("wareName")),
            clean_text(parent.get("totalSales")),
            clean_text(parent.get("spuId")),
            detail.get("shopName") or clean_text(parent.get("shopName")),
            detail.get("venderId") or clean_text(parent.get("venderId")),
            detail.get("shopId") or clean_text(parent.get("shopId")),
            f"https://item.jd.com/{sku}.html",
            jd_price if final_price is not None else None,
            final_price or jd_price,
            variant["skuId"],
            clean_text(variant.get("spec")),
            "；".join(variant.get("attributes", {})),
            variant.get("availableNum"),
            clean_text(variant.get("stockState")),
            detail.get("source", ""),
        ])

    return rows


# ============================================================
# 选择数量 / Excel
# ============================================================

def choose_product_count(total_count):
    if LIMIT:
        chosen = min(LIMIT, total_count)
        print(f"\n按 config.py 的 LIMIT={LIMIT}，本次抓取前 {chosen} 个商品")
        return chosen

    print(f"\n共 {total_count} 个商品")

    while True:
        value = input("抓取前几个？输入 0 或直接回车 = 全部：").strip()

        if not value:
            return total_count

        if value.isascii() and value.isdigit():
            chosen = int(value)
            if chosen == 0:
                return total_count
            if chosen <= total_count:
                return chosen
            print(f"最多 {total_count} 个，请重新输入。")
        else:
            print("请输入非负整数。")


def get_unique_excel_path():
    ensure_dir(EXCEL_DIR)
    base = os.path.join(EXCEL_DIR, EXCEL_NAME)

    if not os.path.exists(base):
        return base

    index = 1
    while True:
        path = os.path.join(EXCEL_DIR, f"jd_result_{index}.xlsx")
        if not os.path.exists(path):
            return path
        index += 1


def save_excel(rows):
    if not rows:
        print("\n没有可以写入 Excel 的数据。")
        return None

    path = get_unique_excel_path()

    print()
    print("=" * 80)
    print("正在写入 Excel...")
    print("=" * 80)

    wb = Workbook()
    ws = wb.active
    ws.title = "商品SKU"

    head_fill = PatternFill("solid", fgColor="C8102E")
    head_font = Font(name="Microsoft YaHei", size=10, bold=True, color="FFFFFF")

    for col, title_text in enumerate(EXCEL_HEADERS, 1):
        cell = ws.cell(1, col, title_text)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    ws.row_dimensions[1].height = 30

    for row_index, values in enumerate(rows, 2):
        for col, value in enumerate(values, 1):
            cell = ws.cell(row_index, col, value)
            cell.font = Font(name="Microsoft YaHei", size=10)
            cell.alignment = Alignment(vertical="center", wrap_text=True)
            if col in (1, 4, 6, 7, 11):
                cell.number_format = "@"
            if col in (9, 10):
                cell.number_format = "0.00"

    for index, width in enumerate(EXCEL_WIDTHS, 1):
        ws.column_dimensions[get_column_letter(index)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(EXCEL_HEADERS))}{len(rows) + 1}"
    wb.save(path)
    wb.close()

    return path


# ============================================================
# 主流程
# ============================================================


def _apply_browser_env(options):
    """统一入口可能要求改用 playwright 自带内核（系统没装 Chrome 时）。"""
    executable = os.environ.get("EC_EXECUTABLE")
    if executable:
        options.pop("channel", None)
        options["executable_path"] = executable
    return options


LOGIN_COOKIE_MARKERS = ("pt_key", "__jdu")


def login_via_browser(profile_dir, timeout=300, warm_url="https://www.jd.com"):
    """打开有头 Chrome 让用户登录京东，登录后读回 cookie。

    和淘宝一样：以前提示"去浏览器登录"，但代码不会打开浏览器，
    用户就卡在拿不到 __jdu / 3AB9D23F7A4B3CSS 上。
    """
    from playwright.sync_api import sync_playwright

    deadline = time.time() + timeout
    cookie_text = ""

    print()
    print("=" * 80)
    print("需要登录京东")
    print("=" * 80)
    print()
    print("  接下来会打开一个 Chrome 窗口，请在里面登录京东。")
    print(f"  登录成功后程序会自动继续，最多等 {timeout // 60} 分钟。")
    print("  登录态会保存下来，以后不用再登。")
    print()

    with sync_playwright() as pw:
        options = _apply_browser_env(dict(
            channel="chrome",
            headless=False,
            user_agent=PC_UA,
            viewport={"width": 1440, "height": 900},
            locale="zh-CN",
            timezone_id="Asia/Shanghai",
            args=["--disable-blink-features=AutomationControlled", "--mute-audio"],
        ))
        proxy = parse_proxy()
        if proxy:
            options["proxy"] = proxy

        context = pw.chromium.launch_persistent_context(profile_dir, **options)
        try:
            page = context.pages[0] if context.pages else context.new_page()
            try:
                page.goto(warm_url, wait_until="domcontentloaded", timeout=60000)
            except Exception:
                pass

            detected = False
            while time.time() < deadline:
                names = {c.get("name") for c in context.cookies() if c.get("name")}
                if all(marker in names for marker in LOGIN_COOKIE_MARKERS):
                    detected = True
                    break
                time.sleep(2)

            print("  检测到登录成功" if detected else "  等待登录超时")

            # 再刷一次，确保设备 cookie（3AB9D23F7A4B3CSS）也下发
            try:
                page.goto(warm_url, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(3000)
            except Exception:
                pass

            cookies = context.cookies()
            cookie_text = "; ".join(
                f"{c.get('name')}={c.get('value')}"
                for c in cookies if c.get("name") and c.get("value"))
        finally:
            try:
                context.close()
            except Exception:
                pass

    return cookie_text


JD_COOKIE_KEYS = ("__jdu", "3AB9D23F7A4B3CSS")


def ensure_cookie():
    """优先用 config.COOKIE；没有就从登录目录自动读。"""
    text = "".join(line.strip() for line in (COOKIE or "").splitlines())
    cookie_map = parse_cookie(text)

    if all(key in cookie_map for key in JD_COOKIE_KEYS):
        return text, cookie_map, "config.py"

    cookies = harvest_cookies(USER_DATA_DIR)
    harvested = parse_cookie(cookies_to_text(cookies))
    if all(key in harvested for key in JD_COOKIE_KEYS):
        return cookies_to_text(cookies), harvested, "登录目录"

    missing = "、".join(k for k in JD_COOKIE_KEYS if k not in harvested)
    raise JdBlocked(
        "拿不到京东搜索需要的 Cookie。\n"
        f"  缺失：{missing}\n"
        "  处理：先运行一次并在弹出的 Chrome 里登录 jd.com，"
        "登录态会自动保存，之后就不用管了。"
    )


def parse_cookie(text):
    cookies = {}
    for part in re.split(r"[;\r\n]+", text):
        part = part.strip()
        if not part or "=" not in part:
            continue
        key, _, value = part.partition("=")
        key = key.strip()
        if key and not any(ch.isspace() for ch in key):
            cookies[key] = value.strip()
    return cookies


def collect_products(cookie_text, cookie_map):
    print()
    print("=" * 80)
    print("正在请求京东搜索数据...")
    print("=" * 80)
    print(f"搜索关键词：{SEARCH.strip()}")

    products, seen = [], set()

    for page_index in range(1, max(1, PAGE_COUNT) + 1):
        ware_list = search_page(SEARCH.strip(), page_index, cookie_text, cookie_map)
        print(f"第 {page_index} 页返回 {len(ware_list)} 个商品")

        if page_index < PAGE_COUNT:
            wait = random.uniform(SEARCH_MIN_INTERVAL, SEARCH_MAX_INTERVAL)
            print(f"  等待 {wait:.1f} 秒后翻页")
            time.sleep(wait)

        for item in ware_list:
            sku = clean_text(item.get("skuId") or item.get("wareId"))
            if not (sku.isascii() and sku.isdigit()) or sku in seen:
                continue
            seen.add(sku)
            products.append(item)

    if not products:
        raise JdBlocked("搜索结果里没有可用商品")

    return products


def parse_args():
    parser = argparse.ArgumentParser(description="京东商品 + SKU 采集（默认读 config.py）")
    parser.add_argument("--proxy", help="临时指定代理，覆盖 config.py 的 PROXY")
    parser.add_argument("--check-ip", action="store_true",
                        help="启动时查一次出口 IP，确认代理是否生效")
    parser.add_argument("--pages", type=int, help="搜索页数，覆盖 PAGE_COUNT")
    parser.add_argument("--keyword", help="搜索词，覆盖 SEARCH")
    return parser.parse_args()


def main():
    import sys
    sys.path.insert(0, str(__import__('pathlib').Path(__file__).resolve().parents[2] / 'app'))
    from search_collect import run
    raise SystemExit(run('jingdong', sys.modules[__name__]))


if __name__ == "__main__":
    main()
