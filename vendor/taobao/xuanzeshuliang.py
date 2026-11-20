import argparse
import os
import random
import re
import json
import time
from urllib.parse import urlparse, parse_qs, unquote

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

# 签名不需要 JS 引擎：aef.js 里的 eE(x) 就是 md5(x)
import hashlib

from openpyxl import Workbook
from playwright.sync_api import sync_playwright, TimeoutError as PlaywrightTimeoutError


# ============================================================
# 基础配置
# ============================================================

# Cookie 与搜索词统一放在同目录 config.py 里维护，本文件不用再改这两项。
try:
    from config import (COOKIE, SEARCH, PROXY, PAGES, MIN_INTERVAL, MAX_INTERVAL,
                    COOLDOWN_SECONDS, validate as validate_config)
except ImportError as error:
    raise SystemExit(
        f"缺少同目录 config.py，或其中缺少 COOKIE / SEARCH：{error}\n"
        "请确保 config.py 和本文件放在同一个文件夹里。"
    ) from None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

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

    淘宝 MTop 接口对 TLS 指纹敏感，裸 requests 的 JA3 是 Python 特征；
    curl_cffi 让握手像 Chrome。没装就退回 requests。
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
    print("  ⚠ 这个出口 IP 不适合跑淘宝：")
    for item in problems:
        print(f"      · {item}")
    print("     机房 IP 和海外异地登录是风控重点，只会更容易被判异常。")
    print("     建议 PROXY 留空，或者用手机热点。")
    print()




# Chrome 持久化登录目录（保存登录态，不是采集结果）
USER_DATA_DIR = r"D:\project\shangpin\taobao_browser"

# 采集结果只保存最终 Excel；商品详情 HTML 仅在内存中解析。
# Excel 直接存到本文件所在目录，跟脚本、config.py 放一起
EXCEL_DIR = BASE_DIR

# ============================================================
# 统一入口（电商采集/启动.py）用的覆盖层
# ============================================================
def _env(name, default=None):
    value = os.environ.get(name)
    return value if value not in (None, "") else default


SEARCH = _env("EC_SEARCH", SEARCH)
PROXY = _env("EC_PROXY", PROXY)
USER_DATA_DIR = _env("EC_PROFILE", USER_DATA_DIR)
EXCEL_DIR = _env("EC_OUTPUT", EXCEL_DIR)


# aef.js 跟脚本放在同一目录，换文件夹也不用改路径
AEF_JS_PATH = os.path.join(BASE_DIR, "aef.js")

# 淘宝 MTop 搜索接口
SEARCH_API_URL = (
    "https://h5api.m.taobao.com/h5/"
    "mtop.relationrecommend.wirelessrecommend.recommend/2.0/"
)

APP_KEY = "12574478"


# ============================================================
# Cookie 与搜索词
# ============================================================
# COOKIE / SEARCH 已移到同目录 config.py，见文件开头的 import。
# 需要换号或换商品时，只改 config.py，不要改这里。


# ============================================================
# 工具函数
# ============================================================

def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def safe_str(value):
    if value is None:
        return ""

    if isinstance(value, (dict, list)):
        return json.dumps(
            value,
            ensure_ascii=False
        )

    return str(value)


def get_unique_excel_path():
    """
    自动避免覆盖旧 Excel：

    taobao_result.xlsx
    taobao_result_1.xlsx
    taobao_result_2.xlsx
    ...
    """

    ensure_dir(EXCEL_DIR)

    base_path = os.path.join(
        EXCEL_DIR,
        "taobao_result.xlsx"
    )

    if not os.path.exists(base_path):
        return base_path

    index = 1

    while True:
        path = os.path.join(
            EXCEL_DIR,
            f"taobao_result_{index}.xlsx"
        )

        if not os.path.exists(path):
            return path

        index += 1


def get_search_keyword(search_url):
    parsed = urlparse(search_url)

    params = parse_qs(
        parsed.query
    )

    return params.get(
        "q",
        [""]
    )[0]


# ============================================================
# 选择爬取数量
# ============================================================

def _preset_count(total_count):
    """统一入口通过 EC_LIMIT 指定数量时，直接用它，不再交互。"""
    raw = os.environ.get("EC_LIMIT")
    if raw in (None, ""):
        return None
    try:
        wanted = int(raw)
    except (TypeError, ValueError):
        return None
    if wanted <= 0:
        return total_count          # 0 = 全部
    return min(wanted, total_count)


def choose_product_count(total_count):

    preset = _preset_count(total_count)
    if preset is not None:
        print(f"\n按设定取前 {preset} 个商品（共 {total_count} 个）")
        return preset

    print()
    print("=" * 80)
    print("商品数量选择")
    print("=" * 80)

    print(
        f"当前搜索到 {total_count} 个商品"
    )

    print(
        "直接回车或输入 0：爬取全部商品"
    )

    print(
        f"输入 1~{total_count}：爬取前 N 个商品"
    )

    while True:

        user_input = input(
            "\n请输入要爬取的商品数量："
        ).strip()

        if user_input == "":
            return total_count

        if not user_input.isdigit():
            print(
                "输入无效，请输入数字。"
            )
            continue

        count = int(user_input)

        if count == 0:
            return total_count

        if 1 <= count <= total_count:
            return count

        print(
            f"请输入 0~{total_count} 之间的数字。"
        )


# ============================================================
# 淘宝 MTop 搜索
# ============================================================

class TaobaoMTop:

    def __init__(self):

        self.start_url = SEARCH_API_URL

        self.headers = {
            "user-agent": (
                "Mozilla/5.0 (Linux; Android 10; "
                "HMA-AL00 Build/HUAWEIHMA-AL00; wv) "
                "AppleWebKit/537.36 "
                "(KHTML, like Gecko) "
                "Version/4.0 Chrome/120.0.0.0 "
                "Mobile Safari/537.36"
            ),

            "referer":
                "https://main.m.taobao.com/",

            "cookie":
                # 多行粘贴的 Cookie 需去掉换行和行首尾空白，才能作为请求头。
                "".join(line.strip() for line in COOKIE.splitlines()),

            "bx-ua":
                "fast-load",

            "accept":
                "*/*",

            "accept-language":
                "zh-CN,zh;q=0.9",
        }

        # 带 TLS 指纹伪装（+ 可选代理）的会话
        self.session = new_session()

        self.session.headers.update(
            self.headers
        )

        # ----------------------------------------------------
        # 签名算法
        # ----------------------------------------------------
        # aef.js 里 eE(x) 的实现就是 md5(输入)，已用 node 逐项比验证实
        # （英文/中文/空串/带 JSON 的串全部一致），所以不用再跑 JS。
        # aef.js 保留在目录里仅作对照，缺了也不影响运行。

    # --------------------------------------------------------
    # 获取 token
    # --------------------------------------------------------

    def get_token(self):

        token = ""

        match = re.search(
            r"_m_h5_tk=([^;]+)",
            COOKIE
        )

        if match:

            token = match.group(1)

            token = token.split("_")[0]

        return token

    # --------------------------------------------------------
    # 生成 sign
    # --------------------------------------------------------

    def get_sign(
        self,
        timestamp,
        data
    ):

        token = self.get_token()

        sign_text = (
            token
            + "&"
            + str(timestamp)
            + "&"
            + APP_KEY
            + "&"
            + data
        )

        return hashlib.md5(
            sign_text.encode("utf-8")
        ).hexdigest()

    # --------------------------------------------------------
    # 构造 data
    # --------------------------------------------------------

    def build_data(
        self,
        keyword,
        page=1
    ):

        params = {
            "device":
                "HMA-AL00",

            "brand":
                "HUAWEI",

            "q":
                keyword,

            "page":
                int(page),

            "pageSize":
                10,
        }

        data = {
            "appId":
                "34385",

            "params":
                json.dumps(
                    params,
                    ensure_ascii=False,
                    separators=(",", ":")
                )
        }

        return json.dumps(
            data,
            ensure_ascii=False,
            separators=(",", ":")
        )

    # --------------------------------------------------------
    # 构造请求参数
    # --------------------------------------------------------

    def build_params(
        self,
        keyword,
        page=1
    ):

        data = self.build_data(
            keyword,
            page
        )

        timestamp = str(
            int(
                time.time() * 1000
            )
        )

        sign = self.get_sign(
            timestamp,
            data
        )

        return {

            "jsv":
                "2.7.4",

            "appKey":
                APP_KEY,

            "t":
                timestamp,

            "sign":
                sign,

            "api":
                "mtop.relationrecommend."
                "wirelessrecommend.recommend",

            "v":
                "2.0",

            "type":
                "jsonp",

            "dataType":
                "jsonp",

            "callback":
                "mtopjsonp20",

            "data":
                data,
        }

    # --------------------------------------------------------
    # 搜索
    # --------------------------------------------------------

    def search(
        self,
        keyword,
        page=1
    ):

        print()
        print("=" * 80)
        print(f"正在请求淘宝搜索数据（第 {page} 页）...")
        print("=" * 80)

        print(
            f"搜索关键词：{keyword}"
        )

        params = self.build_params(
            keyword,
            page
        )

        response = self.session.get(
            self.start_url,
            params=params,
            timeout=30
        )

        text = response.text

        # mtop 令牌握手：第一次调用时服务端会用 Set-Cookie 下发 _m_h5_tk，
        # 此时签名是错的，需要拿新令牌重算一次再请求
        if "FAIL_SYS_TOKEN" in text or "令牌" in text[:600]:
            new_token = ""
            try:
                for cookie in response.cookies:
                    if getattr(cookie, "name", "") == "_m_h5_tk":
                        new_token = cookie.value
                        break
            except Exception:
                new_token = ""

            if new_token:
                print("  令牌过期，用新令牌重试……")
                if "_m_h5_tk=" in COOKIE:
                    updated_cookie = re.sub(
                        r"_m_h5_tk=[^;]+", f"_m_h5_tk={new_token}", COOKIE)
                else:
                    updated_cookie = COOKIE.rstrip("; ") + f"; _m_h5_tk={new_token}"
                globals()["COOKIE"] = updated_cookie

                self.session.headers["cookie"] = "".join(
                    line.strip() for line in COOKIE.splitlines())

                params = self.build_params(keyword, page)
                response = self.session.get(
                    self.start_url,
                    params=params,
                    timeout=30
                )
                text = response.text

        response.raise_for_status()

        match = re.search(
            r"mtopjsonp\d+\((.*)\)\s*$",
            text,
            re.S
        )

        if not match:

            raise RuntimeError(
                "淘宝搜索接口没有返回正常 JSONP 数据。\n"
                f"响应前500字符：\n{text[:500]}"
            )

        json_text = match.group(1)

        try:

            result = json.loads(
                json_text
            )

        except json.JSONDecodeError as e:

            raise RuntimeError(
                f"搜索 JSON 解析失败：{e}\n"
                f"响应前500字符：\n{text[:500]}"
            )

        return result


# ============================================================
# 提取搜索商品
# ============================================================

def collect_products(obj):

    products = {}

    def walk(value):

        if isinstance(value, dict):

            if (
                "item_id" in value
                and "auctionURL" in value
            ):

                item_id = str(
                    value.get("item_id")
                )

                products[item_id] = value

            for child in value.values():
                walk(child)

        elif isinstance(value, list):

            for child in value:
                walk(child)

    walk(obj)

    return list(
        products.values()
    )


# ============================================================
# 提取搜索结果字段
# ============================================================

def extract_search_product(product):

    shop_info = product.get(
        "shopInfo",
        {}
    )

    price_show = product.get(
        "priceShow",
        {}
    )

    return {

        "商品ID":
            safe_str(
                product.get("item_id")
            ),

        "商品标题":
            safe_str(
                product.get("title")
            ),

        "搜索原价":
            safe_str(
                product.get("price")
            ),

        "搜索当前价":
            safe_str(
                price_show.get("price")
            ),

        "销量":
            safe_str(
                product.get("realSales")
            ),

        "店铺名称":
            safe_str(
                shop_info.get("title")
            ),

        "商品链接":
            safe_str(
                product.get("auctionURL")
            ),

        "包装规格":
            safe_str(
                product.get("structuredUSPInfo")
            ),
    }


# ============================================================
# 淘宝商品详情解析器
# ============================================================

class TaobaoDetailParser:

    """
    淘宝新版 PC 商品详情解析器。

    SKU价格核心规则：

        skuBase.skus
             |
             | skuId
             v
        skuCore.sku2info[skuId]
             |
             +---- subPrice.priceText
             |
             +---- price.priceText

    优先：

        店铺优惠后价格 subPrice.priceText

    如果没有：

        优惠前价格 price.priceText

    严禁：

        使用 sku2info["0"] 覆盖所有 SKU。
    """

    def __init__(self, html):

        self.html = html

    @staticmethod
    def _dict(value):

        if isinstance(value, dict):
            return value

        return {}

    @staticmethod
    def _list(value):

        if isinstance(value, list):
            return value

        return []

    # --------------------------------------------------------
    # 提取 ICE JSON
    # --------------------------------------------------------

    @staticmethod
    def _extract_ice_json(html):

        decoder = json.JSONDecoder()

        patterns = [

            r"window\.__ICE_APP_CONTEXT__\s*=\s*",

            r"var\s+b\s*=\s*",
        ]

        for pattern in patterns:

            for match in re.finditer(
                pattern,
                html
            ):

                start = match.end()

                while (
                    start < len(html)
                    and html[start].isspace()
                ):
                    start += 1

                if start >= len(html):
                    continue

                if html[start] != "{":
                    continue

                try:

                    obj, _ = decoder.raw_decode(
                        html[start:]
                    )

                    if (
                        isinstance(obj, dict)
                        and "loaderData" in obj
                    ):

                        return obj

                except json.JSONDecodeError:

                    continue

        return None

    # --------------------------------------------------------
    # 获取优惠前价格
    # --------------------------------------------------------

    @staticmethod
    def _get_original_price(
        sku_info
    ):

        price = sku_info.get(
            "price"
        )

        if not isinstance(price, dict):
            return ""

        for key in (
            "priceText",
            "price"
        ):

            value = price.get(key)

            if value not in (
                None,
                ""
            ):

                return str(value)

        return ""

    # --------------------------------------------------------
    # 获取优惠后价格
    # --------------------------------------------------------

    @staticmethod
    def _get_sub_price(
        sku_info
    ):

        sub_price = sku_info.get(
            "subPrice"
        )

        if not isinstance(
            sub_price,
            dict
        ):

            return ""

        for key in (
            "priceText",
            "price"
        ):

            value = sub_price.get(
                key
            )

            if value not in (
                None,
                ""
            ):

                return str(value)

        return ""

    # --------------------------------------------------------
    # 核心：获取 SKU 当前价格
    # --------------------------------------------------------

    @classmethod
    def _get_sku_price(
        cls,
        sku_info
    ):

        # ----------------------------------------------------
        # 第一优先：
        # 店铺优惠后
        # ----------------------------------------------------

        sub_price = cls._get_sub_price(
            sku_info
        )

        if sub_price != "":
            return sub_price

        # ----------------------------------------------------
        # 第二优先：
        # 优惠前
        # ----------------------------------------------------

        original_price = cls._get_original_price(
            sku_info
        )

        if original_price != "":
            return original_price

        return ""

    # --------------------------------------------------------
    # 获取 SKU 价格详细信息
    # --------------------------------------------------------

    @classmethod
    def _get_sku_price_detail(
        cls,
        sku_info
    ):

        original_price = (
            cls._get_original_price(
                sku_info
            )
        )

        sub_price = (
            cls._get_sub_price(
                sku_info
            )
        )

        if sub_price != "":

            return {
                "SKU价格": sub_price,
                "SKU优惠前价格": original_price,
                "SKU价格来源": "subPrice.priceText",
            }

        if original_price != "":

            return {
                "SKU价格": original_price,
                "SKU优惠前价格": "",
                "SKU价格来源": "price.priceText",
            }

        return {
            "SKU价格": "",
            "SKU优惠前价格": "",
            "SKU价格来源": "",
        }

    # --------------------------------------------------------
    # 建立 SKU 属性映射
    # --------------------------------------------------------

    @staticmethod
    def _build_prop_map(
        sku_base
    ):

        prop_map = {}

        for prop in TaobaoDetailParser._list(
            sku_base.get("props")
        ):

            if not isinstance(
                prop,
                dict
            ):
                continue

            pid = str(
                prop.get("pid", "")
            )

            property_name = safe_str(
                prop.get("name", "")
            )

            for value in TaobaoDetailParser._list(
                prop.get("values")
            ):

                if not isinstance(
                    value,
                    dict
                ):
                    continue

                vid = str(
                    value.get("vid", "")
                )

                value_name = safe_str(
                    value.get("name", "")
                )

                if pid and vid:

                    prop_map[
                        f"{pid}:{vid}"
                    ] = {

                        "property_name":
                            property_name,

                        "value_name":
                            value_name,
                    }

        return prop_map

    # --------------------------------------------------------
    # 解析 propPath
    # --------------------------------------------------------

    @staticmethod
    def _parse_prop_path(
        prop_path,
        prop_map
    ):

        property_names = []

        value_names = []

        value_ids = []

        for part in str(
            prop_path or ""
        ).split(";"):

            part = part.strip()

            if (
                not part
                or ":" not in part
            ):
                continue

            pid, vid = part.split(
                ":",
                1
            )

            key = f"{pid}:{vid}"

            info = prop_map.get(
                key
            )

            if info:

                property_names.append(
                    info["property_name"]
                )

                value_names.append(
                    info["value_name"]
                )

            value_ids.append(
                str(vid)
            )

        return (

            " / ".join(
                x
                for x in property_names
                if x
            ),

            " / ".join(
                x
                for x in value_names
                if x
            ),

            " / ".join(
                x
                for x in value_ids
                if x
            ),
        )

    # --------------------------------------------------------
    # 商品参数
    # --------------------------------------------------------

    @staticmethod
    def _parse_parameters(
        res
    ):

        result = {}

        plus_view = TaobaoDetailParser._dict(
            res.get("plusViewVO")
        )

        industry = TaobaoDetailParser._dict(
            plus_view.get(
                "industryParamVO"
            )
        )

        for list_key in (
            "enhanceParamList",
            "basicParamList"
        ):

            for item in TaobaoDetailParser._list(
                industry.get(
                    list_key
                )
            ):

                if not isinstance(
                    item,
                    dict
                ):
                    continue

                name = item.get(
                    "name"
                )

                value = item.get(
                    "value"
                )

                if name is None:

                    name = item.get(
                        "paramName"
                    )

                if value is None:

                    value = item.get(
                        "paramValue"
                    )

                if name is not None:

                    result[
                        safe_str(name)
                    ] = safe_str(
                        value
                    )

        return result

    # --------------------------------------------------------
    # 主解析函数
    # --------------------------------------------------------

    def parse(self):

        ice = (
            self._extract_ice_json(
                self.html
            )
        )

        if not ice:

            raise RuntimeError(
                "没有找到有效的 "
                "__ICE_APP_CONTEXT__ "
                "商品详情数据。"
            )

        loader_data = self._dict(
            ice.get("loaderData")
        )

        home = self._dict(
            loader_data.get("home")
        )

        home_data = self._dict(
            home.get("data")
        )

        res = self._dict(
            home_data.get("res")
        )

        if not res:

            raise RuntimeError(
                "详情页 ICE 数据中没有找到 res。"
            )

        item = self._dict(
            res.get("item")
        )

        seller = self._dict(
            res.get("seller")
        )

        # ====================================================
        # 商品信息
        # ====================================================

        product_info = {

            "商品ID":
                safe_str(
                    item.get("itemId")
                    or home_data.get(
                        "ssrItemId"
                    )
                ),

            "商品标题":
                safe_str(
                    item.get("title")
                ),

            "商品销量":
                safe_str(
                    item.get(
                        "vagueSellCount"
                    )
                    or self._dict(
                        res.get("titleVO")
                    ).get(
                        "salesDesc"
                    )
                ),

            "SPU ID":
                safe_str(
                    item.get("spuId")
                ),

            "店铺名称":
                safe_str(
                    seller.get(
                        "shopName"
                    )
                    or seller.get(
                        "sellerNick"
                    )
                ),

            "卖家ID":
                safe_str(
                    seller.get(
                        "sellerId"
                    )
                    or seller.get(
                        "userId"
                    )
                ),

            "店铺ID":
                safe_str(
                    seller.get(
                        "shopId"
                    )
                ),
        }

        # ====================================================
        # 商品参数
        # ====================================================

        parameters = (
            self._parse_parameters(
                res
            )
        )

        # ====================================================
        # SKU 原始数据
        # ====================================================

        sku_base = self._dict(
            res.get("skuBase")
        )

        sku_core = self._dict(
            res.get("skuCore")
        )

        sku2info = self._dict(
            sku_core.get(
                "sku2info"
            )
        )

        # ====================================================
        # SKU 属性映射
        # ====================================================

        prop_map = (
            self._build_prop_map(
                sku_base
            )
        )

        sku_entries = (
            self._list(
                sku_base.get("skus")
            )
        )

        print()
        print("=" * 80)
        print("SKU价格映射检查")
        print("=" * 80)

        print(
            f"skuBase.skus 数量："
            f"{len(sku_entries)}"
        )

        print(
            f"skuCore.sku2info 数量："
            f"{len(sku2info)}"
        )

        # ====================================================
        # SKU
        # ====================================================

        sku_list = []

        for sku_entry in sku_entries:

            if not isinstance(
                sku_entry,
                dict
            ):
                continue

            sku_id_raw = (
                sku_entry.get(
                    "skuId"
                )
            )

            if sku_id_raw in (
                None,
                ""
            ):
                continue

            sku_id = str(
                sku_id_raw
            )

            prop_path = safe_str(
                sku_entry.get(
                    "propPath"
                )
            )

            property_name, value_name, value_id = (
                self._parse_prop_path(
                    prop_path,
                    prop_map
                )
            )

            # =================================================
            # 非常重要：
            #
            # 只能根据当前 SKU ID 找对应数据
            #
            # 禁止使用 sku2info["0"]
            # =================================================

            sku_info = sku2info.get(
                sku_id
            )

            if not isinstance(
                sku_info,
                dict
            ):

                print(
                    f"    ⚠ SKU {sku_id} "
                    f"没有找到对应 sku2info"
                )

                sku_info = {}

            # =================================================
            # 获取价格
            # =================================================

            price_detail = (
                self._get_sku_price_detail(
                    sku_info
                )
            )

            sku_price = (
                price_detail[
                    "SKU价格"
                ]
            )

            original_price = (
                price_detail[
                    "SKU优惠前价格"
                ]
            )

            price_source = (
                price_detail[
                    "SKU价格来源"
                ]
            )

            quantity = sku_info.get(
                "quantity",
                ""
            )

            limit = sku_info.get(
                "limit",
                ""
            )

            if limit in (
                None,
                ""
            ):

                limit = sku_info.get(
                    "quantityLimit",
                    ""
                )

            quantity_text = (
                sku_info.get(
                    "quantityText",
                    ""
                )
            )

            logistics = (
                sku_info.get(
                    "logistics",
                    ""
                )
            )

            if logistics in (
                None,
                ""
            ):

                logistics = (
                    sku_info.get(
                        "logisticsTime",
                        ""
                    )
                )

            # =================================================
            # 控制台打印
            # =================================================

            print(
                f"    SKU ID：{sku_id}"
            )

            print(
                f"        规格：{value_name}"
            )

            print(
                f"        propPath：{prop_path}"
            )

            print(
                f"        优惠前："
                f"{original_price}"
            )

            print(
                f"        当前价："
                f"{sku_price}"
            )

            print(
                f"        价格来源："
                f"{price_source}"
            )

            print(
                f"        库存："
                f"{quantity}"
            )

            # =================================================
            # 保存 SKU
            # =================================================

            sku_list.append({

                "SKU ID":
                    sku_id,

                "SKU规格":
                    value_name,

                "SKU属性":
                    property_name,

                "SKU属性ID":
                    value_id,

                "propPath":
                    prop_path,

                "SKU价格":
                    sku_price,

                "SKU优惠前价格":
                    original_price,

                "SKU价格来源":
                    price_source,

                "SKU库存":
                    quantity,

                "SKU限购":
                    limit,

                "SKU状态":
                    safe_str(
                        quantity_text
                    ),

                "物流信息":
                    safe_str(
                        logistics
                    ),
            })

        return {

            "商品信息":
                product_info,

            "商品参数":
                parameters,

            "SKU":
                sku_list,
        }


# ============================================================
# Chrome 商品详情
# ============================================================

class TaobaoDetailBrowser:

    def __init__(
        self,
        playwright
    ):

        self.playwright = playwright

        ensure_dir(
            USER_DATA_DIR
        )

        print()
        print("=" * 80)
        print("正在启动 Chrome...")
        print("=" * 80)

        launch_options = _apply_browser_env(dict(
            channel="chrome",
            headless=False,
            viewport={"width": 1440, "height": 900},
            locale="zh-CN",
            timezone_id="Asia/Shanghai",
            args=["--disable-blink-features=AutomationControlled",
                  "--mute-audio"],
        ))

        _proxy = parse_proxy()
        if _proxy:
            launch_options["proxy"] = _proxy
            print(f"网络出口：{mask_proxy()}")

        self.context = playwright.chromium.launch_persistent_context(
            USER_DATA_DIR, **launch_options
        )

        if self.context.pages:

            self.page = (
                self.context.pages[0]
            )

        else:

            self.page = (
                self.context.new_page()
            )

        self.logged_in = False

    # --------------------------------------------------------
    # 关闭
    # --------------------------------------------------------

    def close(self):

        try:

            self.context.close()

        except Exception:

            pass

    # --------------------------------------------------------
    # 商品详情 URL
    # --------------------------------------------------------

    @staticmethod
    def build_detail_url(
        item_id
    ):

        return (
            "https://detail.tmall.com/item.htm"
            f"?id={item_id}"
        )

    # --------------------------------------------------------
    # 登录页面
    # --------------------------------------------------------

    def is_login_page(self):

        url = (
            self.page.url.lower()
        )

        login_keywords = [

            "login.taobao.com",

            "login.tmall.com",

            "member/login",
        ]

        return any(
            keyword in url
            for keyword in login_keywords
        )

    # --------------------------------------------------------
    # 等待详情数据
    # --------------------------------------------------------

    def wait_for_detail_data(
        self,
        timeout=30000
    ):

        try:

            self.page.wait_for_function(
                """
                () => {
                    return !!(
                        window.__ICE_APP_CONTEXT__
                        &&
                        window.__ICE_APP_CONTEXT__.loaderData
                    );
                }
                """,
                timeout=timeout
            )

            return True

        except PlaywrightTimeoutError:

            return False

    # --------------------------------------------------------
    # 登录
    # --------------------------------------------------------

    def ensure_login(
        self,
        item_id
    ):

        if not self.is_login_page():

            return True

        print()
        print("=" * 80)
        print("检测到淘宝登录页面")
        print("=" * 80)

        print()
        print(
            "请在当前打开的 Chrome 中完成淘宝登录。"
        )

        print(
            "登录成功后不要关闭 Chrome。"
        )

        input(
            "\n登录完成后，回到终端按 Enter 继续..."
        )

        detail_url = (
            self.build_detail_url(
                item_id
            )
        )

        print(
            "\n正在重新打开商品详情页..."
        )

        try:

            self.page.goto(
                detail_url,
                wait_until="domcontentloaded",
                timeout=60000
            )

        except Exception as e:

            print(
                f"重新打开详情页出现异常：{e}"
            )

        time.sleep(3)

        if self.is_login_page():

            print(
                "\n仍然处于登录页面，本次商品跳过。"
            )

            return False

        self.logged_in = True

        return True

    # --------------------------------------------------------
    # 获取商品详情 HTML
    # --------------------------------------------------------

    def get_detail_html(
        self,
        item_id
    ):

        detail_url = (
            self.build_detail_url(
                item_id
            )
        )

        print(
            f"    打开详情页：{detail_url}"
        )

        try:

            self.page.goto(
                detail_url,
                wait_until="domcontentloaded",
                timeout=60000
            )

        except Exception as e:

            print(
                f"    页面打开异常：{e}"
            )

        # ----------------------------------------------------
        # 登录检查
        # ----------------------------------------------------

        if self.is_login_page():

            success = (
                self.ensure_login(
                    item_id
                )
            )

            if not success:
                return None

        # ----------------------------------------------------
        # 等待 ICE
        # ----------------------------------------------------

        loaded = (
            self.wait_for_detail_data(
                timeout=30000
            )
        )

        if not loaded:

            print(
                "    ⚠ 等待 __ICE_APP_CONTEXT__ 超时"
            )

        time.sleep(1.5)

        html = (
            self.page.content()
        )

        print(
            f"    HTML长度：{len(html):,}"
        )

        # ----------------------------------------------------
        # 判断详情数据
        # ----------------------------------------------------

        if "__ICE_APP_CONTEXT__" not in html:

            print(
                "    ✗ HTML中没有找到 "
                "__ICE_APP_CONTEXT__"
            )

            return None

        if "loaderData" not in html:

            print(
                "    ✗ HTML中没有找到 loaderData"
            )

            return None

        return html


# ============================================================
# Excel 表头
# ============================================================

# 价格列直接使用每个 SKU 的价格，与控制台输出一致。
EXCEL_HEADERS = [
    "商品ID",
    "商品标题",
    "商品销量",
    "SPU ID",
    "店铺名称",
    "卖家ID",
    "店铺ID",
    "商品链接",
    "SKU优惠前价格",
    "SKU价格",
    "SKU ID",
    "SKU规格",
    "SKU属性",
    "SKU属性ID",
    "propPath",
    "SKU价格来源",
    "SKU库存",
    "SKU限购",
    "SKU状态",
    "物流信息",
    "是否礼盒装",
    "生产酿造工艺",
    "包装规格",
    "酒精度数",
    "单件净含量",
    "原料成分",
    "包装方式",
    "品牌",
    "系列",
    "品名",
    "产地",
    "口味",
    "是否进口",
    "生产许可证编号",
    "厂名",
    "厂址",
    "保质期",
]


# ============================================================
# 商品 + SKU 合并
# ============================================================

def combine_product_and_sku(
    search_product,
    detail_data
):
    """
    这里只负责把已经解析正确的数据组合起来。

    特别注意：
    SKU 信息已经在 TaobaoDetailParser.parse()
    中解析完成。

    因此这里：
    不重新计算价格
    不重新解析 SKU
    不修改 SKU 数据
    不做任何 SKU fallback

    直接把 parser 返回的 SKU 字典写入 Excel。
    """

    rows = []

    # ========================================================
    # 1. 搜索页商品信息
    # ========================================================

    search_info = extract_search_product(
        search_product
    )

    # ========================================================
    # 2. 详情页商品信息
    # ========================================================

    product_info = detail_data.get(
        "商品信息",
        {}
    )

    if not isinstance(product_info, dict):
        product_info = {}

    # ========================================================
    # 3. 商品参数
    # ========================================================

    parameters = detail_data.get(
        "商品参数",
        {}
    )

    if not isinstance(parameters, dict):
        parameters = {}

    # ========================================================
    # 4. SKU
    # ========================================================

    sku_list = detail_data.get(
        "SKU",
        []
    )

    if not isinstance(sku_list, list):
        sku_list = []

    # ========================================================
    # 5. 商品基础信息
    # ========================================================

    base = {
        "商品ID": safe_str(
            product_info.get("商品ID")
            or search_info.get("商品ID")
        ),

        "商品标题": safe_str(
            product_info.get("商品标题")
            or search_info.get("商品标题")
        ),

        "商品销量": safe_str(
            product_info.get("商品销量")
            or search_info.get("销量")
        ),

        "SPU ID": safe_str(
            product_info.get("SPU ID")
        ),

        "店铺名称": safe_str(
            product_info.get("店铺名称")
            or search_info.get("店铺名称")
        ),

        "卖家ID": safe_str(
            product_info.get("卖家ID")
        ),

        "店铺ID": safe_str(
            product_info.get("店铺ID")
        ),

        "商品链接": safe_str(
            search_info.get("商品链接")
        ),

        "搜索原价": safe_str(
            search_info.get("搜索原价")
        ),

        "搜索当前价": safe_str(
            search_info.get("搜索当前价")
        ),
    }

    # ========================================================
    # 6. 如果没有 SKU
    # ========================================================

    if not sku_list:

        row = dict(base)

        for key, value in parameters.items():
            row[key] = safe_str(value)

        rows.append(row)

        return rows

    # ========================================================
    # 7. 一个 SKU = Excel 一行
    # ========================================================

    for sku in sku_list:

        if not isinstance(sku, dict):
            continue

        # ----------------------------------------------------
        # 直接复制商品基础信息
        # ----------------------------------------------------

        row = dict(base)

        # ----------------------------------------------------
        # 直接复制 SKU 数据
        #
        # 这里非常重要：
        #
        # parser 里面是什么值
        # Excel 就写什么值
        #
        # 不重新计算
        # 不重新匹配
        # 不 fallback
        # ----------------------------------------------------

        row["SKU ID"] = safe_str(
            sku.get("SKU ID", "")
        )

        row["SKU规格"] = safe_str(
            sku.get("SKU规格", "")
        )

        row["SKU属性"] = safe_str(
            sku.get("SKU属性", "")
        )

        row["SKU属性ID"] = safe_str(
            sku.get("SKU属性ID", "")
        )

        row["propPath"] = safe_str(
            sku.get("propPath", "")
        )

        row["SKU价格"] = safe_str(
            sku.get("SKU价格", "")
        )

        row["SKU优惠前价格"] = safe_str(
            sku.get("SKU优惠前价格", "")
        )

        row["SKU价格来源"] = safe_str(
            sku.get("SKU价格来源", "")
        )

        row["SKU库存"] = safe_str(
            sku.get("SKU库存", "")
        )

        row["SKU限购"] = safe_str(
            sku.get("SKU限购", "")
        )

        row["SKU状态"] = safe_str(
            sku.get("SKU状态", "")
        )

        row["物流信息"] = safe_str(
            sku.get("物流信息", "")
        )

        # ----------------------------------------------------
        # 商品参数
        # ----------------------------------------------------

        for key, value in parameters.items():

            # 商品参数不能覆盖已经写入的 SKU 价格、库存等字段。
            if key in EXCEL_HEADERS and key not in row:

                row[key] = safe_str(value)

        # ----------------------------------------------------
        # 加入结果
        # ----------------------------------------------------

        rows.append(row)

    return rows


# ============================================================
# 保存 Excel
# ============================================================

def save_excel(rows):

    if not rows:

        print(
            "\n没有可以写入 Excel 的数据。"
        )

        return None

    excel_path = (
        get_unique_excel_path()
    )

    print()
    print("=" * 80)
    print("正在写入 Excel...")
    print("=" * 80)

    wb = Workbook()

    ws = wb.active

    ws.title = "商品SKU"

    # --------------------------------------------------------
    # 表头
    # --------------------------------------------------------

    ws.append(
        EXCEL_HEADERS
    )

    # --------------------------------------------------------
    # 数据
    # --------------------------------------------------------

    for row in rows:

        ws.append([

            safe_str(
                row.get(
                    header,
                    ""
                )
            )

            for header in EXCEL_HEADERS
        ])

    # --------------------------------------------------------
    # 冻结
    # --------------------------------------------------------

    ws.freeze_panes = "A2"

    # --------------------------------------------------------
    # 自动筛选
    # --------------------------------------------------------

    ws.auto_filter.ref = (
        ws.dimensions
    )

    # --------------------------------------------------------
    # 自动列宽
    # --------------------------------------------------------

    for column_cells in ws.columns:

        max_length = 0

        column_letter = (
            column_cells[
                0
            ].column_letter
        )

        for cell in column_cells:

            value = safe_str(
                cell.value
            )

            length = 0

            for char in value:

                if ord(char) > 127:
                    length += 2
                else:
                    length += 1

            max_length = max(
                max_length,
                length
            )

        width = min(
            max(
                max_length + 2,
                10
            ),
            45
        )

        ws.column_dimensions[
            column_letter
        ].width = width

    # --------------------------------------------------------
    # 保存
    # --------------------------------------------------------

    wb.save(
        excel_path
    )

    print()
    print(
        f"Excel保存成功："
        f"{excel_path}"
    )

    print(
        f"数据行数："
        f"{len(rows)}"
    )

    return excel_path


# ============================================================
# 主程序
# ============================================================


def _apply_browser_env(options):
    """统一入口可能要求改用 playwright 自带内核（系统没装 Chrome 时）。"""
    executable = os.environ.get("EC_EXECUTABLE")
    if executable:
        options.pop("channel", None)
        options["executable_path"] = executable
    return options


def harvest_cookies(profile_dir, headless=True):
    """从 Chrome 登录目录读出 cookie，不打开任何网页。"""
    if not os.path.isdir(profile_dir):
        return []
    try:
        from playwright.sync_api import sync_playwright
    except ImportError:
        return []
    try:
        with sync_playwright() as pw:
            options = _apply_browser_env(dict(
                channel="chrome", headless=headless,
                args=["--disable-blink-features=AutomationControlled", "--mute-audio"]))
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


LOGIN_COOKIE_MARKERS = ("unb", "cookie2")


def login_via_browser(profile_dir, timeout=300, warm_url="https://www.taobao.com"):
    """打开一个有头的 Chrome 让用户登录，登录后把 cookie 读回来。

    以前这里是死循环：提示"在 Chrome 里登录"，但代码直接 return 了，
    Chrome 根本没打开，用户永远没法建立登录态。
    """
    from playwright.sync_api import sync_playwright

    deadline = time.time() + timeout
    cookie_text = ""

    print()
    print("=" * 80)
    print("需要登录淘宝")
    print("=" * 80)
    print()
    print("  接下来会打开一个 Chrome 窗口，请在里面登录淘宝。")
    print("  登录成功后程序会自动继续，最多等 "
          f"{timeout // 60} 分钟。")
    print("  登录态会保存下来，以后不用再登。")
    print()

    with sync_playwright() as pw:
        options = _apply_browser_env(dict(
            channel="chrome",
            headless=False,
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

            if detected:
                print("  检测到登录成功")
            else:
                print("  等待登录超时")

            # 再刷一次首页，让 _m_h5_tk 这类令牌下发
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


def ensure_cookie():
    """优先用 config.COOKIE；没有就从登录目录自动读。"""
    text = "".join(line.strip() for line in (COOKIE or "").splitlines())
    if text:
        return text, "config.py"

    cookies = harvest_cookies(USER_DATA_DIR)
    text = "; ".join(f"{c.get('name')}={c.get('value')}"
                     for c in cookies if c.get("name") and c.get("value"))
    if text:
        return text, "登录目录"
    return "", ""


def parse_args():
    parser = argparse.ArgumentParser(description="淘宝商品采集（默认读 config.py）")
    parser.add_argument("--proxy", help="临时指定代理，覆盖 config.py 的 PROXY")
    parser.add_argument("--check-ip", action="store_true",
                        help="启动时查一次出口 IP，确认代理是否生效")
    parser.add_argument("--keyword", help="搜索词，覆盖 SEARCH")
    return parser.parse_args()


def main():
    import sys
    sys.path.insert(0, str(__import__('pathlib').Path(__file__).resolve().parents[2] / 'app'))
    from search_collect import run
    raise SystemExit(run('taobao', sys.modules[__name__]))


# ============================================================
# 程序入口
# ============================================================

if __name__ == "__main__":
    main()
