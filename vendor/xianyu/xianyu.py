"""闲鱼商品采集：搜索结果 + 商品详情，把页面上看得见的数据尽量全抓下来。

两条通道互补，互为兜底：
    1) 接口通道 —— 真实 Chrome 打开页面，接住页面自己发的 mtop 响应，
       深度遍历把里面所有字段摊平成「键 → 值」，结构化字段优先从这里取。
    2) DOM 通道 —— 直接读渲染好的页面：标题、价格、原价、想要/浏览/收藏/已售、
       卖家与卖家信用、成色标签、描述正文，以及主图相册里的全部图片。
    DOM 决定"页面上看得见的"，接口补充"页面没显示但接口给了的"。

登录态存在 USER_DATA_DIR（默认本目录 xy_browser），首次运行扫码一次即可。
搜索词、页数、是否抓详情、是否下载图片都在同目录 config.py 里改。
运行：python xianyu.py
"""

import argparse
import hashlib
import json
import logging
import os
import random
import re
import shutil
import subprocess
import sys
import time
from urllib.parse import (parse_qsl, quote, unquote, urlencode, urlparse,
                          urlunparse)

try:
    from openpyxl import Workbook
    from openpyxl.drawing.image import Image as XLImage
    from openpyxl.styles import Alignment, Font, PatternFill
    from openpyxl.utils import get_column_letter
except ImportError:
    raise SystemExit("缺少依赖，请先运行：python -m pip install openpyxl")

try:
    from PIL import Image as PILImage
    PILLOW_OK = True
except ImportError:
    PILLOW_OK = False

try:
    from playwright.sync_api import (
        Error as PlaywrightError,
        TimeoutError as PlaywrightTimeoutError,
        sync_playwright,
    )
except ImportError:
    raise SystemExit("缺少依赖，请先运行：python -m pip install playwright")

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
    import urllib3
    from urllib3.exceptions import InsecureRequestWarning
    urllib3.disable_warnings(InsecureRequestWarning)
except Exception:
    pass

# 本机若有代理做 HTTPS 拦截，Python 的证书库会验不过；
# 图片是公开的 CDN 资源，首次失败后自动改为跳过校验。
_TLS_STATE = {"insecure": False}


def fetch_bytes(url, headers, timeout=30, label=""):
    """下载二进制内容，HTTPS 证书验不过时自动降级。"""
    try:
        if _TLS_STATE["insecure"]:
            return requests.get(url, headers=headers, timeout=timeout, verify=False)
        return requests.get(url, headers=headers, timeout=timeout)
    except requests.exceptions.SSLError:
        if not _TLS_STATE["insecure"]:
            _TLS_STATE["insecure"] = True
            print("  提示：HTTPS 证书校验失败（本机可能有代理拦截），"
                  "图片下载改为跳过证书校验")
        try:
            return requests.get(url, headers=headers, timeout=timeout, verify=False)
        except requests.RequestException:
            return None
    except requests.RequestException:
        return None


# ============================================================
# 配置
# ============================================================

try:
    from config import (
        SEARCH, PAGES, LIMIT, USER_DATA_DIR, ENRICH_DETAIL, DOWNLOAD_IMAGES,
        EMBED_IMAGES, MAX_MAIN_IMAGES, IMAGE_SIZE,
        FORCE_LOGIN, MIN_INTERVAL, MAX_INTERVAL, DEBUG_DUMP,
        USE_API, API_FALLBACK_BROWSER, TLS_IMPERSONATE, PROXY,
        validate as validate_config,
    )
except ImportError as error:
    raise SystemExit(
        f"缺少同目录 config.py，或其中缺少必需项：{error}\n"
        "请确保 config.py 和本文件放在同一个文件夹里。"
    ) from None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
PROFILE_DIR = os.path.expandvars(USER_DATA_DIR.strip()) or os.path.join(
    BASE_DIR, "xy_browser")

# 判断"是否已登录"只看身份类 cookie。
# cookie2 / _tb_token_ / _m_h5_tk 匿名访问也会下发，用它们判断会误报。
LOGIN_COOKIE_HINTS = ("unb", "tracknick", "lgc", "_nk_", "havana_lgc_exp")
EXCEL_DIR = BASE_DIR
EXCEL_NAME = "xianyu_result.xlsx"
IMAGE_DIR = os.path.join(BASE_DIR, "images")
TEMPLATE_SEARCH = os.path.join(BASE_DIR, "template_search.json")
TEMPLATE_DETAIL = os.path.join(BASE_DIR, "template_detail.json")

SEARCH_HOST = "https://www.goofish.com"

# ============================================================
# 统一入口（电商采集/启动.py）用的覆盖层
# ============================================================
def _env(name, default=None):
    value = os.environ.get(name)
    return value if value not in (None, "") else default


SEARCH = _env("EC_SEARCH", SEARCH)
PROXY = _env("EC_PROXY", PROXY)
try:
    _pages_env = _env("EC_PAGES")
    if _pages_env is not None:
        PAGES = max(1, int(_pages_env))
except (TypeError, ValueError):
    pass

try:
    _limit_env = _env("EC_LIMIT")
    if _limit_env is not None:
        LIMIT = int(_limit_env)
except (TypeError, ValueError):
    pass
USER_DATA_DIR = _env("EC_PROFILE", USER_DATA_DIR)
EXCEL_DIR = _env("EC_OUTPUT", EXCEL_DIR)
TEMPLATE_SEARCH = _env("EC_OUTPUT") and os.path.join(EXCEL_DIR, "template_search.json") or TEMPLATE_SEARCH
TEMPLATE_DETAIL = _env("EC_OUTPUT") and os.path.join(EXCEL_DIR, "template_detail.json") or TEMPLATE_DETAIL

USER_AGENT = (
    "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
    "(KHTML, like Gecko) Chrome/153.0.0.0 Safari/537.36"
)

# 文本列（放在图片列前面）
TEXT_HEADERS_BEFORE = [
    "序号", "搜索关键词", "商品ID", "商品标题", "价格", "原价", "库存", "地区",
    "卖家昵称", "卖家ID", "卖家在售商品数", "卖家累计销量", "卖家好评率", "卖家信用",
    "卖家签名", "想要人数", "浏览量", "收藏人数", "已售", "运费", "发布时间",
    "商品状态", "商品标签", "商品描述",
]
TEXT_WIDTHS_BEFORE = [
    6, 14, 20, 52, 12, 10, 8, 12,
    20, 16, 14, 14, 12, 26,
    50, 10, 10, 10, 8, 10, 20,
    12, 22, 60,
]

# 文本列（放在图片列后面）
TEXT_HEADERS_AFTER = ["商品链接", "卖家链接", "数据来源", "接口字段"]
TEXT_WIDTHS_AFTER = [44, 40, 16, 60]

# 图片列：主图一 / 主图二 / 主图三 …（闲鱼只有主图相册，没有详情图）
CHINESE_NUMBERS = "一二三四五六七八九十"

# 缩略图尺寸由 config.IMAGE_SIZE 决定：像素 → Excel 单位换算
#   行高单位是磅（1 磅 = 4/3 像素），列宽单位是"字符"（约 7 像素/字符）
THUMB_PX = IMAGE_SIZE
IMAGE_COL_WIDTH = min(255, round(IMAGE_SIZE / 7.0) + 2)          # 列宽（字符）
IMAGE_ROW_HEIGHT = min(409, round(IMAGE_SIZE * 0.75) + 8)        # 行高（磅）
JPEG_QUALITY = 90


def image_column_names(kind, count):
    names = []
    for index in range(1, count + 1):
        number = CHINESE_NUMBERS[index - 1] if index <= len(CHINESE_NUMBERS) else str(index)
        names.append(f"{kind}{number}")
    return names


def build_layout(records):
    """根据实际数据生成表头和列宽。

    两条规则：
      1. 图片列不设上限 —— 实际有几张就开几列；
      2. 文本列里整列全空的直接不输出（没有的东西就不要占位置）。
    返回 (headers, widths, before_idx, after_idx, main_slots, detail_slots, dropped)
    """
    # 闲鱼没有详情图，只有主图相册，所以这里只开主图列
    main_slots = 1
    detail_slots = 0
    if EMBED_IMAGES:
        main_slots = max((len(r.get("main_files") or []) for r in records), default=0)

    dropped = []
    before_idx, after_idx = [], []

    for i, name in enumerate(TEXT_HEADERS_BEFORE):
        if any(r["before"][i] not in (None, "") for r in records):
            before_idx.append(i)
        else:
            dropped.append(name)

    for i, name in enumerate(TEXT_HEADERS_AFTER):
        if any(r["after"][i] not in (None, "") for r in records):
            after_idx.append(i)
        else:
            dropped.append(name)

    headers, widths = [], []
    for i in before_idx:
        headers.append(TEXT_HEADERS_BEFORE[i])
        widths.append(TEXT_WIDTHS_BEFORE[i])

    image_names = image_column_names("主图", main_slots) if EMBED_IMAGES else ["主图链接"]
    for name in image_names:
        headers.append(name)
        widths.append(IMAGE_COL_WIDTH if EMBED_IMAGES else 60)

    for i in after_idx:
        headers.append(TEXT_HEADERS_AFTER[i])
        widths.append(TEXT_WIDTHS_AFTER[i])

    return headers, widths, before_idx, after_idx, main_slots, detail_slots, dropped


# 兼容：模块级完整布局（真实导出时按数据动态生成，这里只是全量定义）
EXCEL_HEADERS, EXCEL_WIDTHS, _, _, _, _, _ = build_layout(
    [{"before": list(TEXT_HEADERS_BEFORE),
      "after": list(TEXT_HEADERS_AFTER),
      "main_files": []}])

# 小图标特征：这些不是商品图
ICON_HINTS = ("-tps-32-32", "-tps-160-160", "-tps-242-150", "tps-480-144",
              "_110x100", "_240x100", "favicon")

# 闲鱼成色/状态标签
STATUS_HINTS = ("全新", "几乎全新", "轻微使用痕迹", "明显使用痕迹", "全新未拆封",
                "99新", "95新", "9成新", "8成新", "未拆封", "已拆封", "闲置")


class XyBlocked(RuntimeError):
    """闲鱼要求登录或限流；继续只会扩大风控范围。"""


HIDE = False            # 当前是否无头模式（无头没法过滑块）


def ensure_dir(path):
    os.makedirs(path, exist_ok=True)


def clean_text(value):
    if value is None:
        return ""
    return re.sub(r"[ \t\u3000]+", " ", str(value)).strip()


def one_line(value):
    return re.sub(r"\s+", " ", clean_text(value))


# ============================================================
# 接口响应处理
# ============================================================

def mtop_json(body):
    """mtop 可能返回纯 JSON，也可能包成 mtopjsonp1({...})。"""
    text = (body or "").strip()

    if not text:
        raise ValueError("响应为空")

    if text[0] not in "{[":
        left = text.find("(")
        right = text.rfind(")")
        if left != -1 and right > left:
            text = text[left + 1:right]

    return json.loads(text)


# 需要人机验证（Baxia 滑块）：登录没掉，但要求证明"你不是机器"
VERIFY_PROBLEMS = (
    "FAIL_SYS_USER_VALIDATE", "_____tmd_____", "punish", "captcha",
    "needVerify", "slider",
)
# 限流：短时间内请求太多，等几分钟就好，跟登录无关
RATE_PROBLEMS = (
    "被挤爆", "请稍后重试", "访问被限制", "访问过于频繁", "操作太频繁", "限流",
)
# 令牌类问题：刷新页面会自动换发，重试即可，不用重新登录
TOKEN_PROBLEMS = (
    "FAIL_SYS_TOKEN_EXOIRED", "FAIL_SYS_TOKEN_EXPIRED", "FAIL_SYS_TOKEN_EMPTY",
    "TOKEN_EXOIRED",
)
# 会话类问题：cookie 还在，但服务端那个会话已经作废了 —— 刷新页面没用，必须重新登录
SESSION_PROBLEMS = (
    "FAIL_SYS_SESSION_EXPIRED", "SESSION_EXPIRED", "SESSION_EMPTY",
)
# 真正的登录问题：只能让用户去登录
LOGIN_PROBLEMS = (
    "RGV587", "FAIL_SYS_ILLEGAL_ACCESS", "NOT_LOGIN", "SESSION_EXPIRED_NOTLOGIN",
    "用户未登录", "请先登录",
)


def mtop_problem(payload):
    """返回 (类型, 说明)；一切正常返回 (None, "")。

    类型：
        "login" —— 需要登录（会话失效或被要求验证）
        "session" —— 服务端会话作废，必须重新登录（刷新无效）
        "token" —— 令牌过期，刷新页面重试即可
        "rate"  —— 真的被限流
        "other" —— 其它业务错误

    注意：RGV587 的文案是"哎哟喂,被挤爆啦,请稍后重试"，看着像限流，
    实际多数情况是"请去登录"。唯一可靠的区分方式看响应里的 data.url：
    指向 passport.goofish.com/mini_login.htm 就是要求登录。
    """
    ret = payload.get("ret") or []
    joined = " ".join(str(item) for item in ret)

    data = payload.get("data")
    url = str(data.get("url") or "") if isinstance(data, dict) else ""

    # 1) 登录跳转优先判断 —— 这条最可靠
    if "passport.goofish.com" in url or "mini_login" in url:
        return "login", f"闲鱼要求登录：{joined or '返回了登录页'}"

    if "login" in url.lower() and "goofish.com" in url.lower():
        return "login", f"闲鱼要求登录：{joined or '返回了登录页'}"

    # 2) 会话类（服务端会话作废，刷新无效，只能重新登录）
    for key in SESSION_PROBLEMS:
        if key in joined:
            return "session", (
                f"登录态已失效：{joined}\n"
                "    这不是令牌问题，刷新页面没用 —— 需要重新扫码登录一次。")

    # 3) 令牌类（刷新页面可自动换发）
    for key in TOKEN_PROBLEMS:
        if key in joined:
            return "token", f"令牌过期：{joined}"

    # 4) 人机验证（登录没掉，过一下滑块就行）
    for key in VERIFY_PROBLEMS:
        if key in joined:
            return "verify", f"需要人机验证：{joined}"

    # 5) 限流类
    for key in RATE_PROBLEMS:
        if key in joined:
            return "rate", f"闲鱼限流：{joined}"

    # 6) 其它明确要求登录的
    for key in LOGIN_PROBLEMS:
        if key in joined:
            return "login", f"闲鱼要求登录：{joined}"

    if joined and "SUCCESS" not in joined:
        return "other", f"接口返回：{joined}"

    return None, ""


def walk_dicts(node, depth=0, max_depth=14):
    if depth > max_depth:
        return

    if isinstance(node, dict):
        yield node
        for value in node.values():
            yield from walk_dicts(value, depth + 1, max_depth)
    elif isinstance(node, list):
        for value in node[:300]:
            yield from walk_dicts(value, depth + 1, max_depth)


def flatten(node, prefix="", depth=0, out=None, max_depth=8):
    """把嵌套 JSON 摊平成 {"a.b.c": 标量}，用于兜住未知结构。"""
    if out is None:
        out = {}

    if depth > max_depth:
        return out

    if isinstance(node, dict):
        for key, value in node.items():
            path = f"{prefix}.{key}" if prefix else str(key)
            flatten(value, path, depth + 1, out, max_depth)
    elif isinstance(node, list):
        for index, value in enumerate(node[:60]):
            flatten(value, f"{prefix}[{index}]", depth + 1, out, max_depth)
    elif node is not None and not isinstance(node, bool):
        out[prefix] = node

    return out


def price_text(value):
    if value is None:
        return ""
    if isinstance(value, str):
        return one_line(value)
    if isinstance(value, (int, float)):
        return str(value)
    if isinstance(value, list):
        return one_line("".join(price_text(item) for item in value))
    if isinstance(value, dict):
        for key in ("text", "price", "value"):
            if key in value:
                return price_text(value[key])
    return ""


def pick(flat, patterns, *, exclude=(), numeric=False):
    """按 key 名正则批量取值，返回第一个看起来合理的命中。"""
    for pattern in patterns:
        regex = re.compile(pattern, re.I)
        for key, value in flat.items():
            if not regex.search(key):
                continue
            if any(re.search(bad, key, re.I) for bad in exclude):
                continue
            text = one_line(value) if isinstance(value, str) else price_text(value)
            if not text or text in ("0", "false", "False", "null"):
                continue
            if numeric and not re.search(r"\d", text):
                continue
            return text
    return ""


# ============================================================
# 搜索结果解析
# ============================================================

def extract_items(payload, keyword):
    items, seen = [], set()

    def add(ex, extra=None):
        item_id = clean_text(ex.get("itemId"))
        if not (item_id.isascii() and item_id.isdigit()) or item_id in seen:
            return
        seen.add(item_id)

        detail = ex.get("detailParams") or {}
        if not isinstance(detail, dict):
            detail = {}

        items.append({
            "itemId": item_id,
            "title": one_line(ex.get("title")),
            "price": price_text(ex.get("price")),
            "originPrice": one_line(detail.get("soldPrice") or ex.get("soldPrice")),
            "area": one_line(ex.get("area")),
            "seller": one_line(ex.get("userNick")),
            "wantCnt": one_line(ex.get("wantCnt") or detail.get("wantCnt")),
            "publishTime": one_line((extra or {}).get("publishTime")),
            "picUrl": one_line(ex.get("picUrl")),
            "keyword": keyword,
            "source": "搜索列表",
        })

    data = payload.get("data") or {}
    for entry in (data.get("resultList") or []) if isinstance(data, dict) else []:
        if not isinstance(entry, dict):
            continue
        node = entry.get("data") or {}
        item = (node.get("item") or {}) if isinstance(node, dict) else {}
        main = (item.get("main") or {}) if isinstance(item, dict) else {}
        ex = (main.get("exContent") or {}) if isinstance(main, dict) else {}
        if isinstance(ex, dict) and ex.get("itemId"):
            click = main.get("clickParam") or {}
            args = (click.get("args") or {}) if isinstance(click, dict) else {}
            add(ex, {"publishTime": args.get("publishTime")})

    if not items:
        for node in walk_dicts(payload):
            if "itemId" in node and ("title" in node or "price" in node):
                add(node)

    return items


# ============================================================
# 接口直连（逆向）
# ============================================================
#
# 实测结论（2026-11，登录态下用 Python 重放验证）：
#   · sign = md5(token & t & appKey & data)，token 取自 _m_h5_tk 下划线前 32 位
#   · bx-ua / bx-umidtoken / bx_et 三个风控参数：搜索接口可省，详情接口本就不发
#   · 详情接口 mtop.taobao.idle.pc.detail 的 body 只有 data
# 所以只要登录 cookie 有效，就能完全脱离浏览器取数据。
#
# 请求模板不硬编码 —— 第一次跑浏览器时把真实请求存下来，
# 之后按模板替换 data 里的关键词/页码/商品ID 再重算 sign。

API_HOST = "https://h5api.m.goofish.com"
API_VERSION = "1.0"


def load_template(path):
    try:
        with open(path, "r", encoding="utf-8") as handle:
            template = json.load(handle)
        if template.get("api") and isinstance(template.get("params"), dict):
            return template
    except (OSError, json.JSONDecodeError, ValueError):
        pass
    return None


def save_template(path, template):
    try:
        with open(path, "w", encoding="utf-8") as handle:
            json.dump(template, handle, ensure_ascii=False, indent=1)
        return True
    except OSError:
        return False


def patch_ids(node, old_id, new_id):
    """把模板里出现的旧商品ID换成新的（递归，保持结构不变）。"""
    if isinstance(node, dict):
        return {k: patch_ids(v, old_id, new_id) for k, v in node.items()}
    if isinstance(node, list):
        return [patch_ids(v, old_id, new_id) for v in node]
    if isinstance(node, str) and node == old_id:
        return new_id
    return node


class MtopClient:
    """用登录 cookie + sign 直接调 mtop，不启动浏览器。"""

    def __init__(self, cookies, templates):
        self.cookies = dict(cookies or {})
        self.templates = templates or {}

        # 用 curl_cffi 让 TLS 指纹像 Chrome，而不是 Python
        self.tls_profile = ""
        proxies = proxy_urls()

        if CURL_CFFI_OK and TLS_IMPERSONATE:
            try:
                self.session = curl_requests.Session(impersonate=TLS_IMPERSONATE,
                                                     proxies=proxies)
                self.tls_profile = f"curl_cffi({TLS_IMPERSONATE})"
            except TypeError:
                self.session = curl_requests.Session(impersonate=TLS_IMPERSONATE)
                if proxies:
                    self.session.proxies.update(proxies)
                self.tls_profile = f"curl_cffi({TLS_IMPERSONATE})"
            except Exception:
                self.session = requests.Session()
        else:
            self.session = requests.Session()

        if proxies and not proxies.get("_applied"):
            try:
                self.session.proxies.update(proxies)
            except Exception:
                pass
        try:
            import urllib3
            urllib3.disable_warnings()
        except Exception:
            pass

    # ---------------- 基础 ----------------

    def _cookie_line(self):
        return "; ".join(f"{k}={v}" for k, v in self.cookies.items())

    def _token(self):
        return str(self.cookies.get("_m_h5_tk", "")).split("_")[0]

    def _refresh_token_from_session(self):
        """mtop 在令牌失效时会在响应里下发新的 _m_h5_tk。"""
        jar = self.session.cookies

        try:
            for cookie in jar:
                name = getattr(cookie, "name", None)
                value = getattr(cookie, "value", None)
                if name == "_m_h5_tk" and value:
                    self.cookies["_m_h5_tk"] = value
                    return True
        except TypeError:
            pass

        # curl_cffi 的 cookie 容器不支持迭代时走这条
        try:
            data = jar.get_dict()
            if data.get("_m_h5_tk"):
                self.cookies["_m_h5_tk"] = data["_m_h5_tk"]
                return True
        except Exception:
            pass

        return False

    def call(self, template, data_obj, retry_token=True):
        api = template["api"]
        app_key = template["params"].get("appKey", "34839810")
        data_str = json.dumps(data_obj, ensure_ascii=False, separators=(",", ":"))

        t_ms = str(int(time.time() * 1000))
        token = self._token()
        sign = hashlib.md5(
            f"{token}&{t_ms}&{app_key}&{data_str}".encode("utf-8")
        ).hexdigest()

        params = {k: v for k, v in template["params"].items()
                  if k not in ("sign", "t", "data")}
        params["t"] = t_ms
        params["sign"] = sign

        # 复放时不能照搬浏览器的全部头：
        #   content-length —— 我们发的 data 和当初抓到的长度不一样，长度不符会直接 400
        #   host / connection / transfer-encoding —— 由 requests 自己处理
        #   cookie —— 下面用合并后的 cookie 覆盖
        # 这些头原样带过去，网关会在 HTTP 层就拒掉（返回 HTML 的 400 页）。
        drop_headers = {
            "content-length", "host", "connection", "cookie",
            "transfer-encoding", "accept-encoding", "upgrade",
            "proxy-connection", "if-none-match", "if-modified-since",
            "keep-alive", "te", "trailer",
        }
        headers = {k: v for k, v in (template.get("headers") or {}).items()
                   if not k.startswith(":") and k.lower() not in drop_headers}
        headers.setdefault("content-type", "application/x-www-form-urlencoded")
        headers["cookie"] = self._cookie_line()

        url = f"{API_HOST}/h5/{api}/{API_VERSION}/"

        try:
            response = self.session.post(url, params=params, data={"data": data_str},
                                        headers=headers, timeout=25, verify=False)
        except requests.RequestException as error:
            return None, f"网络错误：{type(error).__name__}"

        try:
            payload = mtop_json(response.text)
        except Exception:
            body = (response.text or "").strip()
            snippet = " ".join(body.split())[:200] or "（空响应）"
            hint = ""
            if response.status_code == 400:
                hint = ("\n    HTTP 400 通常是请求参数不被接受："
                        "令牌失效、cookie 不完整、或者模板来自一次失败的请求。"
                        "\n    处理：删掉 cookie_cache.json 和 template_*.json 后重跑一次。")
            return None, (f"响应不是 JSON（HTTP {response.status_code}）：{snippet}{hint}")

        kind, message = mtop_problem(payload)
        if kind == "token" and retry_token and self._refresh_token_from_session():
            return self.call(template, data_obj, retry_token=False)

        return payload, ""

    # ---------------- 具体两个接口 ----------------

    def search(self, keyword, page_index=1):
        template = self.templates.get("search")
        if not template:
            return None, "缺少搜索接口模板"

        data = dict(template.get("data") or {})
        data["keyword"] = keyword
        data["pageNumber"] = page_index
        data.setdefault("rowsPerPage", 30)

        payload, error = self.call(template, data)
        if payload is None:
            return None, error

        items = extract_items(payload, keyword)
        if items:
            return items, ""

        kind, message = mtop_problem(payload)
        return [], message or "接口返回成功但没解析出商品"

    def detail(self, item_id):
        template = self.templates.get("detail")
        if not template:
            return None, "缺少详情接口模板"

        old_id = str(template.get("itemId") or "")
        data = patch_ids(template.get("data") or {}, old_id, str(item_id))

        payload, error = self.call(template, data)
        if payload is None:
            return None, error

        flat, problem = XianyuBrowser.merge_detail_payloads(None, [("", payload)])
        if not flat:
            return None, problem or "详情接口没有返回 itemDO"

        detail = {"itemId": item_id, "apiFields": flat}
        detail.update(XianyuBrowser.map_item_fields(None, flat))
        if detail.get("apiMapped"):
            detail["source"] = "详情页(接口直连)"
        return detail, ""


# ============================================================
# 浏览器
# ============================================================

# profile 被别的 Chrome 占用时留下的锁文件
SINGLETON_FILES = ("SingletonLock", "SingletonCookie", "SingletonSocket")


def profile_holders(profile_dir):
    """返回正在使用该登录目录的 chrome 进程 pid 列表。"""
    escaped = os.path.abspath(profile_dir).replace("'", "''")
    script = (
        f"$d = '{escaped}';"
        "$procs = Get-CimInstance Win32_Process -Filter \"Name='chrome.exe'\" | "
        "Where-Object { $_.CommandLine -and $_.CommandLine.Contains($d) };"
        "$procs | ForEach-Object { $_.ProcessId }"
    )
    try:
        completed = subprocess.run(
            ["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
            capture_output=True, text=True, encoding="utf-8", errors="replace",
            timeout=25, check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return []

    pids = []
    for line in (completed.stdout or "").splitlines():
        line = line.strip()
        if line.isdigit():
            pids.append(int(line))
    return pids


def release_profile(profile_dir, *, verbose=True):
    """结束占用登录目录的 Chrome，并清掉残留锁文件。

    只对本脚本自己的登录目录生效；如果 USER_DATA_DIR 指向真实 Chrome 的
    User Data，则一律不动，避免把用户正在用的浏览器杀掉。
    """
    normalized = os.path.normcase(os.path.abspath(profile_dir))

    if "google\\chrome\\user data" in normalized:
        if verbose:
            print("登录目录指向真实 Chrome，跳过自动清理（请手动退出 Chrome 再运行）")
        return 0

    pids = profile_holders(profile_dir)

    if pids and verbose:
        print(f"发现 {len(pids)} 个 Chrome 进程正占用登录目录，正在结束：{pids}")

    for pid in pids:
        try:
            subprocess.run(["taskkill", "/F", "/PID", str(pid)],
                           capture_output=True, timeout=15, check=False)
        except (OSError, subprocess.SubprocessError):
            pass

    if pids:
        time.sleep(1.5)

    for name in SINGLETON_FILES:
        target = os.path.join(profile_dir, name)
        try:
            if os.path.exists(target):
                os.remove(target)
        except OSError:
            pass

    return len(pids)


def parse_proxy(text=None):
    """把 PROXY 字符串拆成 Playwright 需要的结构；空串返回 None。"""
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
    """给 requests / curl_cffi 用的 proxies 字典。"""
    text = (PROXY if text is None else text).strip()
    if not text:
        return None
    if "://" not in text:
        text = "http://" + text
    return {"http": text, "https": text}


def mask_proxy(text=None):
    """日志里显示代理时把账号密码打码。"""
    text = (PROXY if text is None else text).strip()
    if not text:
        return "(直连)"
    if "://" not in text:
        text = "http://" + text
    parsed = urlparse(text)
    host = parsed.hostname or "?"
    port = f":{parsed.port}" if parsed.port else ""
    auth = "带认证@" if parsed.username else ""
    return f"{parsed.scheme}://{auth}{host}{port}"


def api_name_of(url):
    """从 h5api 的 URL 里取 mtop 接口名，不是 h5api 返回 None。"""
    parsed = urlparse(url)
    if not parsed.netloc.startswith("h5api.m.goofish.com"):
        return None
    if "/h5/" not in parsed.path:
        return None
    return parsed.path.split("/h5/")[1].split("/")[0]


def find_item_id(node):
    """在请求 data 里找出商品ID（纯数字且长度 >= 10）。"""
    if isinstance(node, dict):
        for value in node.values():
            found = find_item_id(value)
            if found:
                return found
    elif isinstance(node, list):
        for value in node:
            found = find_item_id(value)
            if found:
                return found
    elif isinstance(node, str) and node.isdigit() and len(node) >= 10:
        return node
    return ""



def _apply_browser_env(options):
    """统一入口可能要求改用 playwright 自带内核（系统没装 Chrome 时）。"""
    executable = os.environ.get("EC_EXECUTABLE")
    if executable:
        options.pop("channel", None)
        options["executable_path"] = executable
    return options


COOKIE_CACHE = os.path.join(BASE_DIR, "cookie_cache.json")
COOKIE_CACHE_HOURS = 6


def token_expiry_seconds(cookies):
    """_m_h5_tk 的格式是 "<token>_<过期毫秒时间戳>"，据此算还剩多久。

    返回剩余秒数；解析不出来返回 None（当作未知，交给服务端判断）。
    """
    raw = str((cookies or {}).get("_m_h5_tk") or "")
    if "_" not in raw:
        return None
    tail = raw.rsplit("_", 1)[-1]
    if not tail.isdigit():
        return None
    try:
        expiry_ms = int(tail)
    except ValueError:
        return None
    # 毫秒时间戳才合理（13 位左右）；否则说明格式不是我们以为的那样
    if expiry_ms < 10 ** 12:
        return None
    return expiry_ms / 1000.0 - time.time()


def load_cookie_cache(max_hours=COOKIE_CACHE_HOURS):
    """读上次采集到的 cookie。

    有了它，接口模式下连 Chrome 都不用启动（读 cookie 需要解密
    Chrome 的 cookie 库，只能起浏览器；缓存下来就能跳过）。

    两道新鲜度检查：
      1. 缓存写入不超过 max_hours 小时
      2. _m_h5_tk 自带的到期时间还没过（留 5 分钟余量）
         以及必须有 cookie2 —— 少了它 mtop 会直接回 HTTP 400
    """
    try:
        with open(COOKIE_CACHE, "r", encoding="utf-8") as handle:
            data = json.load(handle)
        saved_at = float(data.get("saved_at") or 0)
        cookies = data.get("cookies") or {}
    except (OSError, json.JSONDecodeError, ValueError, TypeError):
        return None

    if not cookies or (time.time() - saved_at) >= max_hours * 3600:
        return None

    if not cookies.get("_m_h5_tk") or not cookies.get("cookie2"):
        return None

    left = token_expiry_seconds(cookies)
    if left is not None and left < 300:
        # 令牌快过期或已过期，缓存不可信，让调用方重新读登录目录
        return None

    return cookies


def save_cookie_cache(cookies):
    try:
        with open(COOKIE_CACHE, "w", encoding="utf-8") as handle:
            json.dump({"saved_at": time.time(), "cookies": dict(cookies)},
                      handle, ensure_ascii=False)
        return True
    except (OSError, TypeError):
        return False


def drop_cookie_cache():
    try:
        os.remove(COOKIE_CACHE)
    except OSError:
        pass


# Chrome 的 cookie 是按域隔离的。关键点：
#   www.goofish.com 上那套（页面域的）并不会被发到 h5api.m.goofish.com，
#   而 .goofish.com / .taobao.com 这种父域的会被带上。
# 所以这里按「这个 cookie 会不会被发到接口主机」来过滤，
# 不能简单把所有 cookie 平铺成 {name: value} —— 同名时页面域会盖掉接口域。
API_HOST_NAME = "h5api.m.goofish.com"


def _cookie_applies_to_api(cookie):
    """这个 cookie 会不会被浏览器发到接口主机？"""
    raw = str(cookie.get("domain") or "")
    domain = raw.lstrip(".").lower()

    if not domain:
        return True                     # host-only
    if domain == API_HOST_NAME:
        return True                     # 接口主机专属
    if domain.endswith("goofish.com") or domain.endswith("taobao.com"):
        # 父域（.goofish.com / .taobao.com）会带上；
        # 页面主机域（www.goofish.com）不会，必须排除
        return raw.startswith(".")
    return False                        # 其它站点一律不带


# 覆盖顺序：父域先写，接口主机最后写 → 接口域的值最终生效
DOMAIN_SPECIFICITY = {
    ".taobao.com": 1,
    ".goofish.com": 1,
    API_HOST_NAME: 5,
}


def _domain_rank(domain):
    return DOMAIN_SPECIFICITY.get(str(domain or ""), 1)


def harvest_cookies(playwright, headless=True, verbose=False):
    """从登录目录读出 cookie —— 不打开任何网页，零网络请求。

    按域过滤后再摊平，避免页面域和接口域的同名 cookie 互相覆盖。
    """
    release_profile(PROFILE_DIR, verbose=False)
    ensure_dir(PROFILE_DIR)

    options = dict(
        channel="chrome", headless=headless,
        user_agent=USER_AGENT, viewport={"width": 1440, "height": 900},
        locale="zh-CN", timezone_id="Asia/Shanghai",
        args=["--disable-blink-features=AutomationControlled", "--mute-audio"],
    )
    proxy = parse_proxy()
    if proxy:
        options["proxy"] = proxy

    context = playwright.chromium.launch_persistent_context(PROFILE_DIR, **options)
    try:
        raw = context.cookies()
    finally:
        try:
            context.close()
        except Exception:
            pass

    relevant = [c for c in raw if _cookie_applies_to_api(c)]

    # 从「最不具体」到「最具体」依次覆盖
    relevant.sort(key=lambda c: _domain_rank(c.get("domain")))

    merged = {}
    for cookie in relevant:
        name = cookie.get("name")
        if not name:
            continue
        merged[name] = cookie.get("value")

    if verbose:
        seen_domains = {}
        for cookie in raw:
            key = str(cookie.get("domain") or "")
            seen_domains[key] = seen_domains.get(key, 0) + 1
        print("  cookie 域分布：")
        for domain, count in sorted(seen_domains.items(), key=lambda kv: -kv[1]):
            mark = "接口用" if any(_cookie_applies_to_api(c) for c in raw
                                   if str(c.get("domain") or "") == domain
                                   ) else "忽略"
            print(f"    {domain or '(无域)':26} {count:3} 个   {mark}")
        print(f"  合并后 {len(merged)} 个（原始 {len(raw)} 个）")

    return merged


# 接口字段的精确路径 —— 从真实响应里核出来的，比正则猜准
ITEM_FIELD_PATHS = {
        "title":       ("data.itemDO.title",),
        "desc":        ("data.itemDO.desc",),
        "price":       ("data.itemDO.soldPrice", "data.itemDO.minPrice"),
        "originPrice": ("data.itemDO.oldPrice", "data.itemDO.maxPrice"),
        "stock":       ("data.itemDO.quantity",),
        "wantCnt":     ("data.itemDO.wantCnt",),
        "browseCnt":   ("data.itemDO.browseCnt",),
        "collectCnt":  ("data.itemDO.collectCnt",),
        "soldCnt":     ("data.itemDO.soldCnt",),
        "freight":     ("data.itemDO.transportFee",),
        "status":      ("data.itemDO.itemStatusStr",),
        "created":     ("data.itemDO.gmtCreate",),
        "area":        ("data.sellerDO.city", "data.sellerDO.publishCity"),
        "seller":      ("data.sellerDO.nick", "data.sellerDO.uniqueName"),
        "sellerId":    ("data.sellerDO.sellerId",),
        "sellerItems": ("data.sellerDO.itemCount",),
        "sellerSold":  ("data.sellerDO.hasSoldNumInteger",),
        "creditRate":  ("data.sellerDO.newGoodRatioRate",),
        "credit":      ("data.sellerDO.goodSellerTag.text",),
        "signature":   ("data.sellerDO.signature",),
}


class XianyuBrowser:
    def __init__(self, playwright, headless=False):
        global HIDE
        HIDE = bool(headless)
        existed = os.path.isdir(PROFILE_DIR) and any(
            os.scandir(PROFILE_DIR)
        ) if os.path.isdir(PROFILE_DIR) else False

        ensure_dir(PROFILE_DIR)

        print()
        print("=" * 80)
        print("正在启动 Chrome")
        print("=" * 80)
        print(f"登录目录：{PROFILE_DIR}")
        print(f"目录状态：{'已存在，尝试复用登录态' if existed else '首次创建'}")
        print(f"网络出口：{mask_proxy()}")

        self.context = self._launch(playwright, headless)

        self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
        self._responses = []
        self._templates = {}
        self._current_item = ""
        self.page.on("response", self._collect)
        self.page.on("request", self._collect_request)

        # 登录态是从 cookie 判断的，不靠"每次都问"
        try:
            cookie_names = {cookie.get("name") for cookie in self.context.cookies()}
        except Exception:
            cookie_names = set()
        self.logged_in_hint = any(name in cookie_names for name in LOGIN_COOKIE_HINTS)

        if self.logged_in_hint:
            print("登录态：本地有登录标记，先尝试复用；服务器判定过期时会打开页面重新登录")
        else:
            print("登录态：未检测到。第一次运行需要扫码登录一次，之后会一直复用。")

        if FORCE_LOGIN:
            print("FORCE_LOGIN=True，本次会强制走登录流程")

    def _launch(self, playwright, headless):
        options = _apply_browser_env(dict(
            channel="chrome",
            headless=headless,
            user_agent=USER_AGENT,
            viewport={"width": 1440, "height": 900},
            locale="zh-CN",
            timezone_id="Asia/Shanghai",
            args=["--disable-blink-features=AutomationControlled", "--mute-audio"],
        ))
        proxy = parse_proxy()
        if proxy:
            options["proxy"] = proxy

        try:
            return playwright.chromium.launch_persistent_context(PROFILE_DIR, **options)
        except PlaywrightError as error:
            text = str(error)

            # Chrome 发现登录目录已被别的进程占用时会"交给现有会话"然后自己退出，
            # Playwright 拿不到目标，报 TargetClosedError。
            occupied = (
                "TargetClosed" in type(error).__name__
                or "Target page, context or browser has been closed" in text
                or "已在现有的浏览器会话中打开" in text
            )

            if not occupied:
                raise

            print()
            print("登录目录被占用，正在清理后重试……")
            killed = release_profile(PROFILE_DIR)

            if not killed:
                print("没有找到占用进程，直接重试一次")

            time.sleep(1.0)

            try:
                return playwright.chromium.launch_persistent_context(PROFILE_DIR, **options)
            except PlaywrightError as retry_error:
                raise SystemExit(
                    "Chrome 启动失败：登录目录仍被占用。\n"
                    f"  登录目录：{PROFILE_DIR}\n"
                    "  处理：\n"
                    "    1) 手动关掉所有 Chrome 窗口（含任务栏托盘里的），\n"
                    "    2) 或者在任务管理器里结束所有 chrome.exe，\n"
                    "    3) 再重新运行本脚本。\n"
                    f"  原始错误：{type(retry_error).__name__}"
                ) from None

    def _collect(self, response):
        """响应一到就把 body 读出来存好。

        之前是先把 response 对象攒着、等页面稳定后再读，
        结果大响应（比如 40KB 的详情）读的时候已经过期，
        except 一吞就变成"接口没数据"，字段全空。
        """
        url = response.url
        if "alicdn" in url:
            return
        if not ("mtop" in url or "idle" in url or "goofish.com/api" in url):
            return

        try:
            body = response.text()
        except Exception:
            body = None

        self._responses.append({"url": url, "body": body})

    def _collect_request(self, request):
        """把搜索/详情接口的真实请求存成模板，供以后的接口直连复用。"""
        api = api_name_of(request.url)
        if api == "mtop.taobao.idlemtopsearch.pc.search":
            if "shade" in request.url or "activate" in request.url:
                return
            key = "search"
        elif api == "mtop.taobao.idle.pc.detail":
            key = "detail"
        else:
            return

        if key in self._templates:
            return

        params = dict(parse_qsl(urlparse(request.url).query, keep_blank_values=True))
        body = dict(parse_qsl(request.post_data or "", keep_blank_values=True))
        data_raw = body.get("data") or params.get("data") or ""
        try:
            data_obj = json.loads(data_raw) if data_raw else {}
        except (json.JSONDecodeError, ValueError):
            data_obj = {}

        template = {
            "api": api,
            "params": params,
            "headers": {k: v for k, v in request.all_headers().items()
                        if not k.startswith(":")},
            "data": data_obj,
        }
        if key == "detail":
            template["itemId"] = find_item_id(data_obj) or self._current_item

        self._templates[key] = template

        # 立刻落盘：抓到就算赚到，别等整个流程跑完
        # （以前是搜索失败就返回，模板跟着一起丢了，下次还得再跑浏览器）
        try:
            target = TEMPLATE_SEARCH if key == "search" else TEMPLATE_DETAIL
            if not os.path.isfile(target) and save_template(target, template):
                print(f"  已保存接口模板：{os.path.basename(target)}")
        except Exception:
            pass

    def close(self):
        try:
            self.context.close()
        except Exception:
            pass

    # --------------------------------------------------------
    # 页面状态
    # --------------------------------------------------------

    def has_login_prompt(self):
        try:
            return self.page.evaluate(
                """() => {
                    const t = document.body ? document.body.innerText : '';
                    if (/扫码登录|手机号登录|密码登录/.test(t)) return true;
                    return !!document.querySelector(
                        'iframe[src*="passport.goofish.com"], .login-dialog, [class*="loginDialog"]'
                    );
                }"""
            )
        except Exception:
            return False

    def wait_manual_verify(self):
        """等用户在窗口里手动过滑块，返回是否看起来通过了。"""
        try:
            self.page.goto(f"{SEARCH_HOST}/search?q={quote('键盘')}",
                           wait_until="domcontentloaded", timeout=60000)
        except Exception:
            pass

        try:
            if sys.stdin and sys.stdin.isatty():
                input("\n过完滑块后按 Enter 继续...")
            else:
                return False
        except EOFError:
            return False

        return not self.has_login_prompt()

    def ensure_login(self, force=False):
        """打开真实页面等待用户登录；不以残留 Cookie 作为成功依据。"""
        self.logged_in_hint = False
        if force:
            # 只清理本工具专用闲鱼浏览器上下文的 Cookie，保留整个配置目录。
            self.context.clear_cookies()
        try:
            self.page.goto(f"{SEARCH_HOST}/", wait_until="domcontentloaded", timeout=60000)
            self.page.bring_to_front()
        except Exception as error:
            print(f"无法打开闲鱼登录页面：{type(error).__name__}，请检查网络后重试。")
            return False

        # 首页通常提供登录入口；未找到时让用户在当前页面手动点击。
        try:
            self.page.get_by_text(re.compile(r"^(登录|登录/注册|立即登录|去登录)$")).first.click(timeout=5000)
        except Exception:
            pass
        print()
        print("=" * 80)
        print("闲鱼要求登录")
        print("=" * 80)
        print()
        print("请在当前打开的 Chrome 窗口里：")
        print("  1. 如果没有登录框，先点击页面上的「登录」，再手动扫码/手机号登录")
        print("  2. 登录后确认右上角出现你的头像")
        print("  3. 不要关闭窗口")
        print()
        print(f"登录态会保存到：{PROFILE_DIR}")
        print("以后会优先复用；服务器会话过期时需要再次登录。")

        # 有些环境 stdin 看似是终端但实际读不到（计划任务、管道），
        # 这里统一兜住 EOFError，避免整个程序崩掉。
        try:
            input("\n登录完成后回到终端按 Enter 继续...")
        except EOFError:
            print("\n（无法读取终端输入，本次按未登录处理）")
            return False

        try:
            self.page.goto(f"{SEARCH_HOST}/", wait_until="domcontentloaded", timeout=60000)
            time.sleep(2)
        except Exception:
            pass

        try:
            names = {cookie.get("name") for cookie in self.context.cookies()}
        except Exception:
            names = set()
        self.logged_in_hint = any(name in names for name in LOGIN_COOKIE_HINTS)
        if self.has_login_prompt() or not self.logged_in_hint:
            self.logged_in_hint = False
            print("\n还未检测到登录完成，已停止；下次运行会再次打开登录页面。")
            return False

        print("已检测到登录标记，接下来通过搜索请求确认服务器会话有效。")
        return True

    # --------------------------------------------------------
    # 收集本轮页面发出的接口数据
    # --------------------------------------------------------

    def drain_payloads(self):
        """返回 [(url, payload), ...]，跳过读不到或非 JSON 的。"""
        payloads = []

        for item in self._responses:
            text = item.get("body") if isinstance(item, dict) else None
            if not text:
                continue
            try:
                payload = mtop_json(text)
            except Exception:
                try:
                    payload = json.loads(text)
                except Exception:
                    continue
            payloads.append((item.get("url") if isinstance(item, dict) else "", payload))

        self._responses.clear()
        return payloads

    def _dump(self, payload, name):
        try:
            path = os.path.join(BASE_DIR, f"debug_{name}.json")
            with open(path, "w", encoding="utf-8") as handle:
                json.dump(payload, handle, ensure_ascii=False, indent=1)
            print(f"  已保存原始响应：{os.path.basename(path)}")
        except OSError:
            pass

    # --------------------------------------------------------
    # 搜索
    # --------------------------------------------------------

    SEARCH_API = "idlemtopsearch.pc.search"

    def _take_search_payload(self):
        """从已收集的响应里取出搜索接口那一条，取到就移出队列。"""
        for item in list(self._responses):
            url = item.get("url") or ""
            if self.SEARCH_API not in url:
                continue
            if "shade" in url or "activate" in url:
                continue
            self._responses.remove(item)
            text = item.get("body")
            if not text:
                return None
            try:
                return mtop_json(text)
            except Exception:
                return None
        return None

    def _wait_search_payload(self, timeout=45000):
        """等搜索接口响应；一边等一边滚，触发懒加载。"""
        deadline = time.time() + timeout / 1000
        ticks = 0

        while time.time() < deadline:
            self.page.wait_for_timeout(400)
            payload = self._take_search_payload()
            if payload is not None:
                return payload

            ticks += 1
            if ticks % 5 == 0:      # 每约 2 秒滚一次
                try:
                    self.page.mouse.wheel(0, random.randint(600, 1200))
                except Exception:
                    pass

        return None

    def search(self, keyword, page_index=1, timeout=45000):
        """打开搜索页并等搜索结果。

        闲鱼的结果是懒加载的：首屏只发首页推荐接口，搜索接口要等页面
        滚一下才发出。所以这里一边等搜索接口，一边周期性往下滚。
        """
        url = f"{SEARCH_HOST}/search?q={quote(keyword)}"
        if page_index > 1:
            url += f"&page={page_index}"

        self._responses.clear()

        try:
            self.page.goto(url, wait_until="domcontentloaded", timeout=60000)
        except PlaywrightTimeoutError:
            pass

        payload = self._wait_search_payload(timeout)

        if DEBUG_DUMP and payload is not None:
            self._dump([payload], f"search_{page_index}")

        problem = ""
        kind = None

        if payload is not None:
            items = extract_items(payload, keyword)
            if items:
                return items, ""
            kind, problem = mtop_problem(payload)

            if kind == "token":
                # 令牌过期：刷新一次页面让 mtop 重新换发，再试一次
                print("  令牌过期，刷新后重试……")
                self.page.reload(wait_until="domcontentloaded")
                time.sleep(2)
                payload = self._wait_search_payload()
                if payload is not None:
                    items = extract_items(payload, keyword)
                    if items:
                        return items, ""
                    kind, problem = mtop_problem(payload)

                if kind == "token":
                    # 刷新后还是令牌错误，说明不是令牌，是会话真的没了
                    kind = "session"
                    problem = (
                        f"登录态已失效：{problem}\n"
                        "    刷新页面也没能恢复 —— 需要重新扫码登录一次。")

            if kind == "session":
                print()
                print("=" * 72)
                print("闲鱼的登录态在服务端已经作废了")
                print("=" * 72)
                print("cookie 文件还在 —— 本地看起来是已登录状态，")
                print("但服务器不认这个会话了，刷新页面不会恢复。")
                print()
                print("程序会打开当前浏览器的登录页面，等待你手动重新登录。")
                print("=" * 72)
            if kind in ("session", "login"):
                # 登录过期不能用页面上的残留商品卡片当作成功结果。
                return [], problem
            if not problem:
                problem = "搜索接口返回成功，但没解析出商品"

        # 接口没等到 / 没解析出来 → 用页面 DOM 兜底
        dom_items = self.search_dom(keyword)
        if dom_items:
            return dom_items, ""

        if problem:
            return [], problem

        return [], "没等到搜索接口，页面上也没有商品卡片（可能被限流）"

    def search_dom(self, keyword):
        try:
            cards = self.page.evaluate(
                """() => {
                    const seen = new Set(); const out = [];
                    for (const a of document.querySelectorAll('a[href*="item"]')) {
                        const href = a.getAttribute('href') || '';
                        const m = /[?&]id=(\\d{6,})/.exec(href);
                        if (!m || seen.has(m[1])) continue;
                        seen.add(m[1]);
                        let box = a;
                        for (let i = 0; i < 4 && box.parentElement; i++) box = box.parentElement;
                        const lines = ((box.innerText || '').split('\\n')
                            .map(v => v.trim()).filter(Boolean)).slice(0, 10);
                        out.push({ itemId: m[1], lines });
                    }
                    return out;
                }"""
            )
        except Exception:
            return []

        items = []

        # 页面上 a[href*="item"] 不全是商品：横幅、公告、推荐位也带这个链接。
        # 真正的商品卡片一定有价格（¥xx），据此过滤掉非商品。
        noise_words = ("网页版发闲置", "升级啦", "立即登录", "下载APP", "客服",
                       "帮助中心", "举报", "关于我们", "用户协议", "隐私政策")

        skipped = 0
        for card in cards or []:
            lines = card.get("lines") or []
            price, title = "", ""
            for line in lines:
                if not price and re.search(r"[\u00a5\uffe5]\s*\d", line):
                    price = line
                elif len(line) > len(title):
                    title = line

            if not price:
                skipped += 1
                continue
            if not title or len(title) < 4:
                skipped += 1
                continue
            if any(word in title for word in noise_words):
                skipped += 1
                continue

            items.append({
                "itemId": card["itemId"], "title": title, "price": price,
                "originPrice": "", "area": "", "seller": "", "wantCnt": "",
                "publishTime": "", "picUrl": "", "keyword": keyword,
                "source": "页面DOM",
            })

        if skipped:
            print(f"  （DOM 兜底：跳过 {skipped} 个非商品链接）")

        return items

    # --------------------------------------------------------
    # 商品详情
    # --------------------------------------------------------

    @staticmethod
    def merge_detail_payloads(payloads):
        """从详情页的多个响应里挑出真正有用的那条。

        返回 (字段字典, 问题说明)。
        规则：谁含 data.itemDO 就用谁 —— 别的接口（登录态查询、红点、
        推荐位）返回什么错误都不影响详情数据。找不到 itemDO 时，
        才退而合并所有"没报错"的负载。
        """
        for url, payload in payloads:
            data = payload.get("data")
            if isinstance(data, dict) and "itemDO" in data:
                return flatten(payload), ""

        flat, problems = {}, []
        for url, payload in payloads:
            kind, message = mtop_problem(payload)
            if kind:
                problems.append(message)
                continue
            flat.update(flatten(payload))

        return flat, (problems[0] if problems and not flat else "")

    def enrich(self, item_id, timeout=30000):
        """返回 (结构化详情 dict, 是否拿到有效数据)。"""
        self._current_item = str(item_id)
        url = f"{SEARCH_HOST}/item?id={item_id}"
        self._responses.clear()

        # 搜索页的脚本可能还在跳转（懒加载 / SPA 重定向），
        # 会把我们的 goto 打断。这种错误重试即可，不是致命问题。
        navigated = False
        for attempt in range(3):
            try:
                self.page.goto(url, wait_until="domcontentloaded",
                               timeout=timeout + 15000)
                navigated = True
                break
            except PlaywrightTimeoutError:
                navigated = True
                break
            except PlaywrightError as error:
                text = str(error)
                if "interrupted by another navigation" in text or "net::ERR_ABORTED" in text:
                    if attempt < 2:
                        time.sleep(1.0 + attempt)
                        continue
                # 其它情况（比如页面被关掉）记下来，交给上层处理
                print(f"  打开详情页失败：{text.splitlines()[0][:140]}")
                break

        if not navigated:
            return {"itemId": item_id, "apiProblem": "详情页打不开",
                    "apiFields": {}}, False

        try:
            self.page.wait_for_function(
                """() => {
                    const t = document.body ? document.body.innerText : '';
                    return t.length > 400 && !/网络不见了|扫码登录/.test(t.slice(0, 200));
                }""",
                timeout=timeout,
            )
        except PlaywrightTimeoutError:
            pass

        self.page.wait_for_timeout(2500)

        payloads = self.drain_payloads()

        if DEBUG_DUMP and payloads:
            self._dump(payloads, f"detail_{item_id}")

        detail = {"itemId": item_id}
        flat, problem = self.merge_detail_payloads(payloads)

        if problem:
            detail["apiProblem"] = problem

        detail["apiFields"] = flat

        if flat:
            detail.update(self.map_item_fields(flat))

        # DOM 通道（页面渲染出来的东西）
        dom = self.dom_detail()
        detail.update(dom)

        # 接口有值就用接口的（精确），只有接口缺的才回落到 DOM
        for key in ("title", "price", "originPrice", "desc", "seller", "area", "credit"):
            if not detail.get(key) and detail.get(f"dom_{key}"):
                detail[key] = detail[f"dom_{key}"]

        for key in ("want", "browse", "collect", "sold"):
            if not detail.get(f"{key}Cnt") and detail.get(f"dom_{key}"):
                detail[f"{key}Cnt"] = detail[f"dom_{key}"]

        if detail.get("apiMapped"):
            detail["source"] = "详情页(接口)"
        elif detail.get("title") or detail.get("dom_images"):
            problem = detail.get("apiProblem") or ""
            if "验证" in problem:
                detail["source"] = "详情页(DOM,接口要滑块)"
            elif "登录" in problem:
                detail["source"] = "详情页(DOM,接口要登录)"
            elif "限流" in problem:
                detail["source"] = "详情页(DOM,接口限流)"
            elif problem:
                detail["source"] = "详情页(DOM,接口异常)"
            else:
                detail["source"] = "详情页(DOM)"
        else:
            detail["source"] = "仅列表"

        ok = bool(detail.get("title") or detail.get("dom_images") or flat)
        return detail, ok

    @staticmethod
    def map_item_fields(flat):
        """按精确路径取字段，再补标签和主图。"""
        result = {}

        for name, paths in ITEM_FIELD_PATHS.items():
            for path in paths:
                value = flat.get(path)
                if value in (None, ""):
                    continue
                result[name] = one_line(value)
                break

        # 标签：priceRelativeTags[*].text（包邮、验货宝 这类）
        tags = []
        for key in sorted(flat):
            if not re.fullmatch(r"data\.itemDO\.priceRelativeTags\[\d+\]\.text", key):
                continue
            text = one_line(flat[key])
            if text and text not in tags:
                tags.append(text)
        if tags:
            result["tags"] = "、".join(tags)

        # 主图：imageInfos[*].url，按原始顺序，跳过尺寸过小的图标
        pattern = re.compile(r"data\.itemDO\.imageInfos\[(\d+)\]\.url$")
        entries = []
        for key, value in flat.items():
            found = pattern.fullmatch(key)
            if not found:
                continue
            index = int(found.group(1))
            width = flat.get(f"data.itemDO.imageInfos[{index}].widthSize") or 0
            height = flat.get(f"data.itemDO.imageInfos[{index}].heightSize") or 0
            try:
                width, height = int(width), int(height)
            except (TypeError, ValueError):
                width = height = 0
            entries.append((index, one_line(value), width, height))

        images = []
        for _, url, width, height in sorted(entries):
            if not url:
                continue
            if not looks_like_product_image(width, height):
                continue
            images.append(url)

        if images:
            result["apiMainImages"] = dedupe_images(images)

        result["apiMapped"] = bool(result)
        return result

    # --------------------------------------------------------
    # DOM 提取
    # --------------------------------------------------------

    def dom_detail(self):
        try:
            raw = self.page.evaluate(
                """() => {
                    const text = (document.body.innerText || '');
                    const lines = text.split('\\n').map(s => s.trim()).filter(Boolean);

                    const imgs = [];
                    const push = (url, w, h) => {
                        if (!url || !/^https?:/.test(url)) return;
                        if (imgs.some(i => i.url === url)) return;
                        imgs.push({ url, w: w || 0, h: h || 0 });
                    };
                    // 只收商品图：至少 250px，比例不能太扁/太窄
                    const keep = (w, h) => {
                        if (!w || !h) return true;
                        if (w < 250 || h < 250) return false;
                        const r = w / h;
                        return r >= 0.35 && r <= 2.8;
                    };
                    for (const el of document.querySelectorAll('img')) {
                        push(el.currentSrc || el.src, el.naturalWidth, el.naturalHeight);
                    }
                    for (const el of document.querySelectorAll('*')) {
                        const bg = getComputedStyle(el).backgroundImage;
                        if (bg && bg.indexOf('url(') === 0) {
                            push(bg.slice(5, -2), el.clientWidth, el.clientHeight);
                        }
                    }

                    const links = [...document.querySelectorAll('a[href]')]
                        .map(a => a.getAttribute('href'))
                        .filter(h => h && /personal|userId=/.test(h));

                    // 结构化定位描述块：先找"描述不符包邮退"这类承诺条，
                    // 再取它后面的兄弟节点 —— 比数文本行可靠得多。
                    const descText = (() => {
                        const hints = ['描述不符', '支持验货', '7天无理由',
                                       '退货退款', '验货宝'];
                        const nodes = [...document.querySelectorAll('div,section,article')];
                        const cands = nodes.filter(el => {
                            const t = el.innerText || '';
                            if (t.length > 260) return false;
                            return hints.some(h => t.includes(h));
                        });
                        if (!cands.length) return '';
                        cands.sort((a, b) => (a.innerText.length - b.innerText.length));
                        let node = cands[0];
                        for (let i = 0; i < 5 && node; i++) {
                            const next = node.nextElementSibling;
                            if (next) {
                                const t = (next.innerText || '').trim();
                                if (t.length > 20) return t;
                            }
                            node = node.parentElement;
                        }
                        return '';
                    })();

                    return { lines, imgs, links, descText };
                }"""
            )
        except Exception:
            return {}

        lines = raw.get("lines") or []
        result = {"dom_lines": lines}

        nav_words = ("搜索", "发布", "消息", "我的", "首页", "登录", "注册", "客服")

        # 标题：只在开头 15 行里挑，且排除带句读的句子
        # （描述正文常比标题长，只按长度挑会把描述第一句当标题）
        title_bad = ("卖家信用", "包邮", "直接买", "人想要", "人浏览", "小刀",
                     "描述不符", "立即购买", "聊一聊", "前来过", "来闲鱼",
                     "卖出", "好评率", "信用", "宝贝")
        head = [line for line in lines[:15]
                if len(line) >= 6
                and line not in nav_words
                and not any(word in line for word in title_bad)
                and not re.search(r"[，。！？；|｜]", line)]
        if not head:
            head = [line for line in lines[:15] if len(line) >= 6 and line not in nav_words]
        if head:
            result["dom_title"] = max(head, key=len)

        # 价格：闲鱼页面常见写法是"2人小刀价 ¥0.10"、"直接买 ¥1.00"，
        # 也有单独一行 "¥" + 数字的组合。按可信度依次尝试。
        money = re.compile(r"[\u00a5\uffe5]\s*(\d+(?:\.\d+)?)")

        for line in lines[:80]:
            if any(key in line for key in ("小刀价", "直接买", "到手价", "一口价", "秒杀价")):
                found = money.search(line)
                if found:
                    result["dom_price"] = f"\u00a5{found.group(1)}"
                    break

        if not result.get("dom_price"):
            for line in lines[:80]:
                if re.fullmatch(r"[\u00a5\uffe5]\s*\d+(?:\.\d+)?", line):
                    result["dom_price"] = line
                    break

        if not result.get("dom_price"):
            # 最后兜底："¥" 独占一行 + 下一行数字（可能被拆成整数和小数两行）
            index = next((i for i, line in enumerate(lines[:80])
                          if re.fullmatch(r"[\u00a5\uffe5]", line)), -1)
            if index != -1 and index + 1 < len(lines):
                merged = lines[index] + lines[index + 1]
                if index + 2 < len(lines) and re.fullmatch(r"\.\d+", lines[index + 2]):
                    merged += lines[index + 2]
                result["dom_price"] = merged

        # 原价：只认明确标了"原价"的行，不再拿"第二个 ¥ 值"瞎猜
        for line in lines[:80]:
            if "原价" in line:
                found = money.search(line)
                if found:
                    result["dom_originPrice"] = f"\u00a5{found.group(1)}"
                    break

        # 人数类
        for line in lines:
            for key, pattern in (("want", r"(\d+)\s*人想要"),
                                 ("browse", r"(\d+)\s*人浏览"),
                                 ("collect", r"(\d+)\s*人收藏"),
                                 ("sold", r"已售\s*(\d+)")):
                found = re.search(pattern, line)
                if found and not result.get(f"dom_{key}"):
                    result[f"dom_{key}"] = found.group(1)

        # 成色标签
        for line in lines:
            if line in STATUS_HINTS:
                result["dom_status"] = line
                break

        # 卖家 + 信用
        # "卖家信用极好" 前面那行有时是昵称，有时是"赞 鱼小铺L5"这种徽章，
        # 所以要往前多找几行，跳过徽章再取昵称。
        badge_words = ("赞", "鱼小铺", "Lv", "LV", "级", "粉丝", "关注", "宝贝",
                       "前来过", "来闲鱼", "卖出", "好评率", "信用")
        for index, line in enumerate(lines):
            if "卖家信用" not in line:
                continue
            result["dom_credit"] = line
            for back in range(1, 5):
                pos = index - back
                if pos < 0:
                    break
                candidate = lines[pos].strip()
                if not (2 <= len(candidate) <= 24):
                    continue
                if any(word in candidate for word in badge_words):
                    continue
                if re.fullmatch(r"[\d\W]+", candidate):
                    continue
                result["dom_seller"] = candidate
                break
            break

        # 地区：只认"市/区/县/镇/街道/省/州/旗"结尾的短行。
        # 宁可空着，也不把"机械键盘/搜索/综合"这类词当地名 ——
        # 接口里有 data.sellerDO.city 时本来就轮不到这里。
        area_noise = (
            "隐私", "政策", "协议", "集团", "公司", "原因", "区域", "服务", "举报",
            "客服", "反馈", "关于", "版本", "执照", "信用代码", "搜索", "发布",
            "消息", "我的", "首页", "登录", "注册", "闲鱼", "全部", "综合",
            "降价", "价格", "确定", "个人", "闲置", "验货", "担保", "包邮",
            "超赞", "鱼小铺", "全新", "严选", "转卖", "加载", "标题", "描述",
        )
        title_now = result.get("dom_title") or ""
        for line in lines:
            if not (2 <= len(line) <= 10):
                continue
            if any(word in line for word in area_noise):
                continue
            if re.search(r"[\u00a5\uffe5]|想要|浏览|收藏|卖家|\d", line):
                continue
            if title_now and (line in title_now or title_now in line):
                continue
            if re.search(r"(市|区|县|镇|街道|省|州|旗)$", line):
                result["dom_area"] = line
                break

        # ---------------- 描述正文 ----------------
        # 顺序：① 结构化定位（承诺条后面的兄弟节点）② 文本行区块兜底
        # 之前取"最长的一行"，结果抓到页脚备案号；
        # 后来又取"最后一个锚点往后"，结果锚点落进推荐商品区，抓成一堆卖家昵称。
        desc_noise_keywords = (
            "备案", "网械", "平台备字", "许可证", "增值电信", "广播电视",
            "营业性演出", "人力资源服务", "信用代码", "统一社会", "营业执照",
            "阿里巴巴", "淘宝", "天猫", "1688", "阿里健康", "全球速卖通",
            "高德", "飞猪", "优酷", "大麦", "阿里影业", "钉钉", "支付宝",
            "阿里妈妈", "阿里云", "一淘", "版权", "侵权", "服务协议",
            "隐私", "政策", "举报", "客服", "反馈", "关于我们", "诚聘",
            "开放平台", "卖家信用", "已售", "人想要", "人浏览", "人收藏",
            "描述不符", "包邮退", "退货且", "运费由卖家承担",
            "为你推荐", "相似商品", "看过", "猜你喜欢",
        )
        desc_stop = ("聊一聊", "立即购买", "收藏", "担保交易", "展开", "收起",
                     "我要留言", "留言", "相似商品", "为你推荐", "看过")

        def clean_block(text):
            kept = []
            for raw_line in (text or "").split("\n"):
                line = raw_line.strip()
                if not line or len(line) < 2:
                    continue
                if any(word in line for word in desc_noise_keywords):
                    continue
                if re.fullmatch(r"[\u00a5\uffe5]?\s*\d+(?:\.\d+)?\s*元?", line):
                    continue
                if title_now and (line == title_now):
                    continue
                kept.append(line)
            return "\n".join(kept[:20])

        title_now = result.get("dom_title") or ""

        # ① 结构化
        structural = clean_block(raw.get("descText"))
        if len(structural) >= 10:
            result["dom_desc"] = structural

        # ② 文本行兜底：用"第一个"锚点，并且只往下一小段
        if not result.get("dom_desc"):
            start_index = -1
            for index, line in enumerate(lines):
                if re.search(r"人想要|人浏览|描述不符|直接买", line):
                    start_index = index
                    break          # 第一个锚点就在价格区，比最后一个可靠
            if start_index == -1:
                for index, line in enumerate(lines):
                    if re.search(r"^[\u00a5\uffe5]\s*\d", line):
                        start_index = index
                        break

            if start_index != -1:
                block = []
                for line in lines[start_index + 1:start_index + 30]:
                    if any(line.startswith(word) for word in desc_stop):
                        break
                    block.append(line)
                fallback = clean_block("\n".join(block))
                if len(fallback) >= 10:
                    result["dom_desc"] = fallback

        # 图片：只取主图相册里的图（闲鱼没有详情图）
        # 第一轮用严格条件（>=250px 且比例正常）
        strict, lenient = [], []
        for image in raw.get("imgs") or []:
            url = image["url"]
            if any(hint in url for hint in ICON_HINTS):
                continue
            width, height = image.get("w"), image.get("h")
            if looks_like_product_image(width, height):
                strict.append(url)
            elif not (width and height and width < 100 and height < 100):
                lenient.append(url)

        images = dedupe_images(strict) or dedupe_images(lenient)

        result["dom_images"] = images
        result["dom_main_images"] = images

        for link in raw.get("links") or []:
            if "userId=" in link or "personal" in link:
                result["dom_sellerLink"] = link if link.startswith("http") else SEARCH_HOST + link
                break

        return result


# ============================================================
# 图片下载
# ============================================================

# 阿里 CDN 图片的唯一文件 id，形如 O1CN01tha5zR1ucA6HWFfbU。
# 同一张图会以多种写法出现（img/gw 换域名、http/https、_120x120 缩放后缀），
# 用 URL 全串去重挡不住，必须按这个 id 认。
ALI_IMAGE_ID = re.compile(r"(O1CN[0-9A-Za-z]+)")


def image_identity(url):
    """一张图的身份：优先取阿里文件 id，取不到就用去掉参数的原 URL。"""
    found = ALI_IMAGE_ID.search(url or "")
    if found:
        return found.group(1)
    return (url or "").split("?")[0]


def image_score(url):
    """同一张图的多个变体里，分数越高越接近原图。"""
    score = 0
    clean = (url or "").split("?")[0]
    if not re.search(r"_\d+x\d+[a-z0-9]*\.(?:jpg|jpeg|png|webp)$", clean, re.I):
        score += 3              # 没有 _120x120 这类缩放后缀 = 原图
    if "tps-" not in clean:
        score += 1              # -tps- 多为小尺寸变体
    if clean.startswith("https://"):
        score += 1
    if "/imgextra/" in clean:
        score += 1
    return score


def dedupe_images(urls):
    """按文件 id 去重，保留顺序，同 id 留最清晰的那个变体。"""
    best, order = {}, []
    for url in urls or []:
        url = (url or "").strip()
        if not url:
            continue
        key = image_identity(url)
        if key not in best:
            best[key] = url
            order.append(key)
        elif image_score(url) > image_score(best[key]):
            best[key] = url
    return [best[key] for key in order]


def looks_like_product_image(width, height):
    """商品图判定：够大 + 比例正常，用来滤掉横幅、图标、装饰条。"""
    if not width or not height:
        return True                      # 尺寸未知时放行，后面还有缩略图兜底
    if width < 250 or height < 250:
        return False
    ratio = width / height
    return 0.35 <= ratio <= 2.8


def fetch_image_files(item_id, main_urls):
    """下载主图到 images/<商品ID>/，返回本地路径列表（顺序与 url 一致）。

    闲鱼只有"主图相册"，没有淘宝那种图文详情，所以这里只处理主图。
    已存在的文件直接复用，重跑不会重复下载。
    """
    folder = os.path.join(IMAGE_DIR, str(item_id))
    ensure_dir(folder)

    headers = {"User-Agent": USER_AGENT, "Referer": f"{SEARCH_HOST}/item?id={item_id}"}
    main_urls = dedupe_images(main_urls)
    picked = main_urls if not MAX_MAIN_IMAGES else main_urls[:MAX_MAIN_IMAGES]
    saved = []

    for index, url in enumerate(picked, 1):
        suffix = os.path.splitext(urlparse(url).path)[1].lower()
        if suffix not in (".jpg", ".jpeg", ".png", ".webp"):
            suffix = ".jpg"
        # 文件名带上图片 id，相册顺序变化时不会复用错的旧文件
        ident = image_identity(url)
        tag = f"_{ident[:12]}" if ALI_IMAGE_ID.fullmatch(ident) else ""
        target = os.path.join(folder, f"main_{index:02d}{tag}{suffix}")

        if not (os.path.exists(target) and os.path.getsize(target) > 0):
            response = fetch_bytes(url, headers, label="图片")
            if response is not None and response.status_code == 200 and response.content:
                with open(target, "wb") as handle:
                    handle.write(response.content)

        if os.path.exists(target) and os.path.getsize(target) > 0:
            saved.append(target)

    return saved


def make_thumbnail(path, holder):
    """把本地图片缩成缩略图，返回 openpyxl 可直接 add_image 的对象。"""
    from io import BytesIO

    with PILImage.open(path) as source:
        # 带透明通道的调色板图直接 convert("RGB") 会告警，先统一成 RGBA
        if source.mode in ("P", "LA", "RGBA"):
            source = source.convert("RGBA")
            canvas = PILImage.new("RGB", source.size, (255, 255, 255))
            canvas.paste(source, mask=source.split()[-1])
            image = canvas
        else:
            image = source.convert("RGB")
        image.thumbnail((THUMB_PX, THUMB_PX), PILImage.LANCZOS)
        buffer = BytesIO()
        image.save(buffer, format="JPEG", quality=JPEG_QUALITY)
        buffer.seek(0)
        size = image.size

    holder.append(buffer)      # 保活：openpyxl 在 wb.save 时才真正读流
    xl_image = XLImage(buffer)
    xl_image.width, xl_image.height = size
    return xl_image


def diagnose(keyword="测试"):
    """排查接口模式为什么用不了。

    只发 1 次搜索接口请求，把关键信息全打出来：
    模板、cookie、令牌、请求构造、服务端返回。
    """
    print("=" * 80)
    print("接口模式诊断（只发 1 次请求）")
    print("=" * 80)

    # ---------- 1) 模板 ----------
    print("\n[1] 接口模板")
    templates = {"search": load_template(TEMPLATE_SEARCH),
                 "detail": load_template(TEMPLATE_DETAIL)}
    for key, path in (("search", TEMPLATE_SEARCH), ("detail", TEMPLATE_DETAIL)):
        found = templates.get(key)
        mark = "有" if found else "缺失"
        print(f"  {mark:4} {os.path.basename(path)}")
        if found:
            print(f"       api = {found.get('api')}")
            params = found.get("params") or {}
            print(f"       参数 = {sorted(params)}")
            data_keys = sorted((found.get("data") or {}))
            print(f"       data 字段 = {data_keys}")

    if not (templates.get("search") and templates.get("detail")):
        print("\n  结论：模板不全，接口模式无法启用")
        print("  处理：先跑一次浏览器模式（不加 --diagnose），跑成功后会自动生成")
        return 1

    # ---------- 2) cookie ----------
    print("\n[2] 登录态")
    cookies = load_cookie_cache()
    source = "cookie 缓存"
    if not cookies:
        print("  缓存没有，正在启动无窗口 Chrome 读取……")
        with sync_playwright() as playwright:
            cookies = harvest_cookies(playwright, verbose=True)
        source = "登录目录"
        if cookies:
            save_cookie_cache(cookies)

    if not cookies:
        print("\n  结论：完全读不到 cookie，登录目录可能是空的")
        print(f"  目录：{PROFILE_DIR}")
        return 1

    print(f"  来源：{source}，共 {len(cookies)} 个")
    for name in ("_m_h5_tk", "unb", "cookie2", "tracknick", "sgcookie", "havana_lgc2_77"):
        value = cookies.get(name)
        if value is None:
            print(f"  {name:16} 缺失")
        else:
            shown = value if len(str(value)) <= 18 else f"{str(value)[:12]}…（{len(str(value))} 字符）"
            print(f"  {name:16} {shown}")

    token = str(cookies.get("_m_h5_tk", "")).split("_")[0]
    print(f"\n  签名用 token：{token or '（空！）'}")

    # ---------- 3) 发一次请求 ----------
    print("\n[3] 实测一次搜索请求")
    client = MtopClient(cookies, templates)
    print(f"  TLS 指纹：{client.tls_profile or 'requests 默认'}")
    print(f"  网络出口：{mask_proxy()}")

    template = templates["search"]
    data = dict(template.get("data") or {})
    data["keyword"] = keyword
    data["pageNumber"] = 1
    data.setdefault("rowsPerPage", 30)

    payload, error = client.call(template, data)

    if payload is None:
        print(f"\n  请求失败：{error}")
        return 1

    ret = " ".join(str(x) for x in (payload.get("ret") or []))
    print(f"  ret：{ret}")

    kind, message = mtop_problem(payload)
    items = extract_items(payload, keyword)
    print(f"  解析到商品：{len(items)} 条")

    print("\n[4] 结论")
    if items:
        print(f"  接口模式可用（关键词「{keyword}」拿到 {len(items)} 条）")
        print("  直接运行 python xianyu.py 就会走接口，不会打开浏览器。")
        return 0

    print(f"  接口模式不可用：{message or '未知原因'}")
    if kind == "session":
        print("  处理：python xianyu.py --relogin 重新扫码登录")
    elif kind == "token":
        print("  处理：再跑一次 --diagnose，看令牌握手能不能恢复")
    elif kind == "verify":
        print("  处理：在浏览器里手动过一次滑块，再跑 --diagnose")
    elif kind == "rate":
        print("  处理：等 30 分钟以上再试")
    elif kind == "login":
        print("  处理：python xianyu.py --relogin")
    else:
        print("  把上面的 ret 发出来，我按具体错误判断")
    return 1


def parse_args():
    parser = argparse.ArgumentParser(
        description="闲鱼商品采集（默认读同目录 config.py）")
    parser.add_argument("--keyword", help="搜索词，覆盖 config.py 的 SEARCH")
    parser.add_argument("--pages", type=int, help="搜索页数，覆盖 PAGES")
    parser.add_argument("--limit", type=int, help="抓前 N 条；0=全部，覆盖 LIMIT")
    parser.add_argument("--relogin", action="store_true",
                        help="清掉登录 cookie，重新扫码登录一次")
    parser.add_argument("--reset-profile", action="store_true",
                        help="删掉整个登录目录，连设备指纹一起换新（更彻底）")
    parser.add_argument("--proxy", help="临时指定代理，覆盖 config.py 的 PROXY")
    parser.add_argument("--check-ip", action="store_true",
                        help="启动时查一次出口 IP，确认代理生效")
    parser.add_argument("--cooldown", type=int, default=180,
                        help="被限流时等多少秒后自动重试一次；0=不重试（默认 180）")
    parser.add_argument("--diagnose", action="store_true",
                        help="排查接口模式：只发 1 次请求，打印模板/cookie/令牌/返回")
    return parser.parse_args()


def reset_profile_dir():
    """删掉登录目录，让下次启动是全新设备。"""
    release_profile(PROFILE_DIR, verbose=False)
    if os.path.isdir(PROFILE_DIR):
        print(f"正在删除登录目录：{PROFILE_DIR}")
        shutil.rmtree(PROFILE_DIR, ignore_errors=True)
        if os.path.isdir(PROFILE_DIR):
            print("删除未完成（可能有文件被占用），请手动删除该目录")
        else:
            print("登录目录已清空，下次启动需要重新扫码登录")


def choose_count(total):
    limit = LIMIT if override_limit is None else override_limit
    if limit:
        chosen = min(limit, total)
        print(f"\n按 LIMIT={limit}，本次取前 {chosen} 条")
        return chosen

    print(f"\n共 {total} 条商品")

    while True:
        try:
            value = input("抓取前几条？输入 0 或直接回车 = 全部：").strip()
        except EOFError:
            print("\n（无法读取终端输入，默认抓取全部）")
            return total

        if not value:
            return total

        if value.isascii() and value.isdigit():
            chosen = int(value)
            if chosen == 0:
                return total
            if chosen <= total:
                return chosen
            print(f"最多 {total} 条，请重新输入。")
        else:
            print("请输入非负整数。")


override_limit = None   # --limit 覆盖用


def get_unique_excel_path():
    ensure_dir(EXCEL_DIR)
    base = os.path.join(EXCEL_DIR, EXCEL_NAME)

    if not os.path.exists(base):
        return base

    index = 1
    while True:
        path = os.path.join(EXCEL_DIR, f"xianyu_result_{index}.xlsx")
        if not os.path.exists(path):
            return path
        index += 1


def save_excel(records):
    if not records:
        print("\n没有可以写入 Excel 的数据。")
        return None

    path = get_unique_excel_path()
    headers, widths, before_idx, after_idx, main_slots, detail_slots, dropped = \
        build_layout(records)

    print()
    print("=" * 80)
    print("正在写入 Excel...")
    print("=" * 80)
    print(f"实际输出 {len(headers)} 列（主图 {main_slots} 列）")

    if dropped:
        print(f"整列全空，已省略 {len(dropped)} 列：{'、'.join(dropped)}")
        print("  这些字段只有详情接口才有，接口被拦（滑块/登录）时就拿不到。")

    embed = EMBED_IMAGES and PILLOW_OK

    if EMBED_IMAGES and not PILLOW_OK:
        print("[警告] 没装 Pillow，图片列只能写链接：pip install pillow")

    wb = Workbook()
    ws = wb.active
    ws.title = "闲鱼商品"

    head_fill = PatternFill("solid", fgColor="FFDA00")
    head_font = Font(name="Microsoft YaHei", size=10, bold=True, color="333333")
    body_font = Font(name="Microsoft YaHei", size=10)
    wrap = Alignment(vertical="center", wrap_text=True)

    for col, title_text in enumerate(headers, 1):
        cell = ws.cell(1, col, title_text)
        cell.fill = head_fill
        cell.font = head_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 30

    main_start = len(before_idx) + 1
    after_start = main_start + main_slots

    buffers = []
    embedded = 0
    thumb_failed = 0

    for row_index, record in enumerate(records, 2):
        col = 1
        for i in before_idx:
            cell = ws.cell(row_index, col, record["before"][i])
            cell.font = body_font
            cell.alignment = wrap
            if headers[col - 1] == "商品ID":
                cell.number_format = "@"
            col += 1

        if embed:
            for offset, file_path in enumerate(record.get("main_files") or []):
                if offset >= main_slots:
                    break
                try:
                    picture = make_thumbnail(file_path, buffers)
                    ws.add_image(picture,
                                 f"{get_column_letter(main_start + offset)}{row_index}")
                    embedded += 1
                except Exception:
                    thumb_failed += 1
        else:
            cell = ws.cell(row_index, main_start, "\n".join(record.get("main_urls") or []))
            cell.font = body_font
            cell.alignment = wrap

        col = after_start
        for i in after_idx:
            cell = ws.cell(row_index, col, record["after"][i])
            cell.font = body_font
            cell.alignment = wrap
            col += 1

        has_picture = embed and bool(record.get("main_files"))
        ws.row_dimensions[row_index].height = IMAGE_ROW_HEIGHT if has_picture else 60

    for index, width in enumerate(widths, 1):
        ws.column_dimensions[get_column_letter(index)].width = width

    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:{get_column_letter(len(headers))}{len(records) + 1}"

    wb.save(path)
    wb.close()

    if embed:
        print(f"已嵌入 {embedded} 张缩略图" + (f"（{thumb_failed} 张失败）" if thumb_failed else ""))

    return path


def publish_text(item, detail):
    raw = detail.get("created") or item.get("publishTime") or ""
    if raw.isdigit() and len(raw) >= 10:
        return time.strftime("%Y-%m-%d %H:%M", time.localtime(int(raw[:10])))
    return raw


def build_record(index, item, detail=None, files=None):
    """组装一条记录：前段文本 + 图片文件路径 + 后段文本。"""
    detail = detail or {}
    files = files or []
    item_id = item["itemId"]

    main_urls = detail.get("apiMainImages") or detail.get("dom_main_images") or []
    if not main_urls and item.get("picUrl"):
        main_urls = [item["picUrl"]]

    # 闲鱼没有独立的详情图，主图相册就是这个商品的全部图片
    # 再兜一层去重：接口图和 DOM 图混在一起时可能有重复变体
    main_urls = dedupe_images(main_urls)

    api_fields = detail.get("apiFields") or {}
    api_json = json.dumps(api_fields, ensure_ascii=False) if api_fields else ""
    if len(api_json) > 32000:
        api_json = api_json[:32000] + "…(截断)"

    seller_link = detail.get("dom_sellerLink") or ""
    if not seller_link and detail.get("sellerId"):
        seller_link = f"{SEARCH_HOST}/personal?userId={detail['sellerId']}"

    before = [
        index,
        item.get("keyword", ""),
        item_id,
        item.get("title") or detail.get("title", ""),
        detail.get("price") or item.get("price", ""),
        detail.get("originPrice") or item.get("originPrice", ""),
        detail.get("stock", ""),
        detail.get("area") or item.get("area", ""),
        detail.get("seller") or item.get("seller", ""),
        detail.get("sellerId", ""),
        detail.get("sellerItems", ""),
        detail.get("sellerSold", ""),
        detail.get("creditRate", ""),
        detail.get("credit", ""),
        detail.get("signature", ""),
        detail.get("wantCnt") or item.get("wantCnt", ""),
        detail.get("browseCnt", ""),
        detail.get("collectCnt", ""),
        detail.get("soldCnt", ""),
        detail.get("freight", ""),
        publish_text(item, detail),
        detail.get("status", ""),
        detail.get("tags", ""),
        detail.get("desc", ""),
    ]

    after = [
        f"{SEARCH_HOST}/item?id={item_id}",
        seller_link,
        detail.get("source") or item.get("source", ""),
        api_json,
    ]

    return {
        "before": before,
        "after": after,
        "main_urls": main_urls,
        "main_files": files if isinstance(files, list) else (files.get("main") or []),
    }


# ============================================================
# 主流程
# ============================================================

def probe_exit_ip(timeout=20):
    """查出口 IP 的详细信息，用来判断这个 IP 能不能用来跑闲鱼。

    返回 {ip, country, isp, hosting} 或 None。
    hosting=True 表示云服务器 IP —— 电商风控会重点关照，强烈不建议。
    """
    proxies = proxy_urls()
    url = ("http://ip-api.com/json/?fields=query,country,regionName,"
           "city,isp,org,hosting")

    try:
        if CURL_CFFI_OK and TLS_IMPERSONATE:
            response = curl_requests.get(url, timeout=timeout, verify=False,
                                         proxies=proxies,
                                         impersonate=TLS_IMPERSONATE)
        else:
            response = requests.get(url, timeout=timeout, verify=False,
                                    proxies=proxies)
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
    """出口 IP 不适合跑电商时给出提示。"""
    if not info or not info.get("query"):
        return

    problems = []
    if info.get("hosting"):
        problems.append("这是云服务器/机房 IP（hosting=True）")
    if str(info.get("country") or "").lower() not in ("china",):
        problems.append(f"出口在海外（{info.get('country')}）")

    if not problems:
        return

    print()
    print("  ⚠ 这个出口 IP 不适合跑闲鱼：")
    for item in problems:
        print(f"      · {item}")
    print("     电商风控对机房 IP 和异地登录都很敏感，可能直接弹滑块")
    print("     甚至影响账号。建议 PROXY 留空走本机网络，")
    print("     或者用手机热点（运营商 IP 最接近真人用户）。")
    print()


def run_api_mode(keyword, pages, templates):
    """接口直连模式。成功导出 Excel 返回 True，失败返回 False。"""
    print("=" * 80)
    print("模式：接口直连（不渲染页面，请求量约为浏览器模式的 1/6）")
    print("=" * 80)

    # 先看缓存：命中就完全不启动 Chrome
    cookies = load_cookie_cache()
    cookie_source = "缓存"

    if cookies:
        print(f"已从缓存读取 {len(cookies)} 个 cookie（没有启动浏览器）")
    else:
        print("正在读取登录态（会短暂启动一个无窗口的 Chrome）……")
        with sync_playwright() as playwright:
            cookies = harvest_cookies(playwright)
        cookie_source = "登录目录"
        if cookies:
            save_cookie_cache(cookies)

    if not cookies.get("_m_h5_tk"):
        print("没有拿到 _m_h5_tk，无法计算签名，改用浏览器模式")
        return False

    print(f"cookie 来源：{cookie_source}，共 {len(cookies)} 个（未发任何网络请求）")
    client = MtopClient(cookies, templates)
    print(f"TLS 指纹：{client.tls_profile or 'requests 默认（Python 特征，建议装 curl_cffi）'}")
    print(f"网络出口：{mask_proxy()}")

    collected, seen = [], set()

    for page_index in range(1, max(1, pages) + 1):
        items, problem = client.search(keyword, page_index)
        if problem:
            print(f"第 {page_index} 页失败：{problem}")

            # 登录态/会话出问题时，缓存里的 cookie 已经没用了，
            # 清掉它，下次运行会重新从登录目录读（或提示重新登录）
            lowered = problem.lower()
            if ("失效" in problem or "登录" in problem or "session" in lowered
                    or "token" in lowered):
                drop_cookie_cache()
                print("  已清掉 cookie 缓存，下次运行会重新读取登录态")

            return False
        fresh = [row for row in items if row["itemId"] not in seen]
        for row in fresh:
            seen.add(row["itemId"])
        collected.extend(fresh)
        print(f"第 {page_index} 页解析到 {len(items)} 条，新增 {len(fresh)} 条")
        if page_index < pages:
            time.sleep(random.uniform(MIN_INTERVAL, MAX_INTERVAL))

    if not collected:
        print("没有采集到商品")
        return False

    target = choose_count(len(collected))
    collected = collected[:target]

    records = []
    detail_ok = detail_fail = image_saved = 0

    for index, item in enumerate(collected, 1):
        print()
        print("-" * 80)
        print(f"[{index}/{len(collected)}] {item['itemId']}  {item['title'][:34]}")
        print("-" * 80)

        detail = {}
        files = []

        if ENRICH_DETAIL:
            detail, problem = client.detail(item["itemId"])
            if detail:
                detail_ok += 1
                print(f"  标题：{clean_text(detail.get('title'))[:50]}")
                print(f"  价格：{detail.get('price') or '-'}"
                      f"   库存：{detail.get('stock') or '-'}"
                      f"   想要：{detail.get('wantCnt') or '-'}")
                print(f"  卖家：{detail.get('seller') or '-'}"
                      f"   {detail.get('credit') or ''}")
                main_urls = detail.get("apiMainImages") or []
            else:
                detail_fail += 1
                print(f"  详情失败：{problem}")
                detail = {"apiProblem": problem, "apiFields": {}}
                main_urls = []

            if not main_urls and item.get("picUrl"):
                main_urls = [item["picUrl"]]

            if (EMBED_IMAGES or DOWNLOAD_IMAGES) and main_urls:
                files = fetch_image_files(item["itemId"], main_urls)
                image_saved += len(files)
                print(f"  主图：{len(files)} 张")

            if index < len(collected):
                time.sleep(random.uniform(MIN_INTERVAL, MAX_INTERVAL))

        records.append(build_record(index, item, detail, files))

    excel_path = save_excel(records)

    print()
    print("=" * 80)
    print("采集完成")
    print("=" * 80)
    print(f"采集方式：接口直连")
    print(f"搜索关键词：{keyword}")
    print(f"采集条数：{len(records)}")
    if ENRICH_DETAIL:
        print(f"详情成功：{detail_ok}")
        print(f"详情失败：{detail_fail}")
        if EMBED_IMAGES or DOWNLOAD_IMAGES:
            print(f"图片文件：{image_saved} 张 → {IMAGE_DIR}")
    if excel_path:
        print(f"\nExcel：{excel_path}")
    print()

    return True


def main():
    import sys
    sys.path.insert(0, str(__import__('pathlib').Path(__file__).resolve().parents[2] / 'app'))
    from search_collect import run
    raise SystemExit(run('xianyu', sys.modules[__name__]))


if __name__ == "__main__":
    main()
