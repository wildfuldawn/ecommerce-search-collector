"""网络出口：代理解析、出口 IP 检查、机房/海外 IP 预警。

三个平台共用。判断逻辑和闲鱼那版一致：
hosting=True（云服务器）或海外出口都会让电商风控提高怀疑度。
"""

import os

try:
    from curl_cffi import requests as curl_requests
    CURL_CFFI_OK = True
except ImportError:
    curl_requests = None
    CURL_CFFI_OK = False

try:
    import requests
    REQUESTS_OK = True
except ImportError:
    requests = None
    REQUESTS_OK = False

try:
    import urllib3
    urllib3.disable_warnings()
except Exception:
    pass

from urllib.parse import unquote, urlparse


def parse_proxy(text):
    """把代理字符串拆成 Playwright 需要的结构；空串返回 None。"""
    text = (text or "").strip()
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

    result = {"server": server}
    if parsed.username:
        result["username"] = unquote(parsed.username)
    if parsed.password:
        result["password"] = unquote(parsed.password)
    return result


def proxy_urls(text):
    text = (text or "").strip()
    if not text:
        return None
    if "://" not in text:
        text = "http://" + text
    return {"http": text, "https": text}


def mask_proxy(text):
    text = (text or "").strip()
    if not text:
        return "(直连)"
    if "://" not in text:
        text = "http://" + text
    parsed = urlparse(text)
    host = parsed.hostname or "?"
    port = f":{parsed.port}" if parsed.port else ""
    auth = "带认证@" if parsed.username else ""
    return f"{parsed.scheme}://{auth}{host}{port}"


def new_session(proxy="", impersonate="chrome"):
    """带 TLS 指纹伪装和代理的 HTTP 会话。"""
    proxies = proxy_urls(proxy)

    if CURL_CFFI_OK:
        try:
            return curl_requests.Session(impersonate=impersonate, proxies=proxies)
        except TypeError:
            session = curl_requests.Session(impersonate=impersonate)
            if proxies:
                session.proxies.update(proxies)
            return session
        except Exception:
            pass

    if not REQUESTS_OK:
        raise RuntimeError(
            "缺少 requests 库，无法发 HTTP 请求。\n"
            "  处理：python -m pip install requests"
        )

    session = requests.Session()
    if proxies:
        session.proxies.update(proxies)
    return session


def probe_exit_ip(proxy="", timeout=20):
    """查出口 IP 详情。返回 dict 或 None。"""
    try:
        session = new_session(proxy)
    except RuntimeError:
        return None
    url = ("http://ip-api.com/json/?fields=query,country,regionName,"
           "city,isp,org,hosting")
    try:
        return session.get(url, timeout=timeout, verify=False).json()
    except Exception:
        return None


def describe_exit_ip(info):
    if not info or not info.get("query"):
        return "(查询失败，代理可能不通)"
    return (f"{info.get('query')}  {info.get('country')} / "
            f"{info.get('isp') or info.get('org') or ''}").strip()


def risk_warnings(info):
    """返回不适合跑电商出口的原因列表；空列表表示没问题。"""
    if not info or not info.get("query"):
        return []

    problems = []
    if info.get("hosting"):
        problems.append("这是云服务器/机房 IP（hosting=True）")
    if str(info.get("country") or "").lower() != "china":
        problems.append(f"出口在海外（{info.get('country')}）")
    return problems


def print_exit_report(proxy=""):
    """打印出口 IP 报告，返回是否可用。"""
    info = probe_exit_ip(proxy)
    print(f"出口 IP：{describe_exit_ip(info)}")

    troubles = risk_warnings(info)
    if troubles:
        print()
        print("  ⚠ 这个出口不适合跑国内电商：")
        for item in troubles:
            print(f"      · {item}")
        print("     机房 IP 和海外异地登录是风控重点，比直连更容易被拦。")
        print("     建议代理留空，或者用手机热点（运营商 IP 最接近真人）。")
        return False

    if info and info.get("query"):
        print("  （国内家宽/移动网络，适合跑电商采集）")
    return True
