"""淘宝/京东：打开专用浏览器，等待用户完成登录后才交还采集器。"""
import re


IDENTITY_COOKIES = {
    "taobao": (("unb", "cookie2"),),
    "jingdong": (("pt_key", "pt_pin"), ("thor", "pin"), ("thor", "pinId")),
}
HOME_URLS = {"taobao": "https://www.taobao.com", "jingdong": "https://www.jd.com"}
NAMES = {"taobao": "淘宝", "jingdong": "京东"}


def cookie_values(cookie_text):
    values = {}
    for part in (cookie_text or "").replace("\n", ";").split(";"):
        key, sep, value = part.strip().partition("=")
        if sep and key and value:
            values[key] = value
    return values


def has_identity(platform, cookie_text):
    values = cookie_values(cookie_text)
    return any(all(values.get(key) for key in group) for group in IDENTITY_COOKIES[platform])


def wait_for_login(module, platform, force=False):
    name = NAMES[platform]
    home = HOME_URLS[platform]
    options = module._apply_browser_env(dict(
        channel="chrome", headless=False, viewport={"width": 1440, "height": 900},
        locale="zh-CN", timezone_id="Asia/Shanghai",
        args=["--disable-blink-features=AutomationControlled", "--mute-audio"],
    ))
    proxy = module.parse_proxy()
    if proxy:
        options["proxy"] = proxy
    with module.sync_playwright() as pw:
        context = pw.chromium.launch_persistent_context(module.USER_DATA_DIR, **options)
        try:
            if force:
                context.clear_cookies()
            page = context.pages[0] if context.pages else context.new_page()
            page.goto(home, wait_until="domcontentloaded", timeout=60000)
            page.bring_to_front()
            try:
                page.get_by_text(re.compile(r"^(亲，请登录|你好，请登录|登录|登录/注册|立即登录|请登录)$")).first.click(timeout=5000)
            except Exception:
                pass
            print(f"已打开{name}页面，请在浏览器里手动完成登录；若没有登录框，请点击页面上的「登录」。")
            print("登录完成后回到终端按 Enter；未完成登录会继续等待，输入 q 可取消。")
            while True:
                try:
                    answer = input(f"等待{name}登录（Enter 确认，q 取消）：").strip().lower()
                except EOFError:
                    print("无法读取终端输入，本次未确认登录，已停止。")
                    return ""
                if answer == "q" or page.is_closed():
                    return ""
                page.goto(home, wait_until="domcontentloaded", timeout=60000)
                page.wait_for_timeout(1500)
                cookies = context.cookies()
                cookie_text = "; ".join(f"{c['name']}={c['value']}" for c in cookies
                                        if c.get("name") and c.get("value"))
                if not has_identity(platform, cookie_text):
                    print("还未检测到账号登录，请继续在浏览器中完成登录，再按 Enter。")
                    continue
                if platform == "jingdong" and not all(cookie_values(cookie_text).get(key) for key in module.JD_COOKIE_KEYS):
                    print("账号标记已出现，但搜索所需 Cookie 尚未齐全，请在浏览器中检查登录后再按 Enter。")
                    continue
                print(f"已检测到{name}账号登录标记，将通过搜索请求确认服务器会话有效。")
                return cookie_text
        finally:
            context.close()


class LoginRequired(RuntimeError):
    pass


def login_error(message):
    value = str(message).upper()
    return any(marker in value for marker in (
        "SESSION_EXPIRED", "SESSION_EMPTY", "NOT_LOGIN", "NOTLOGIN", "NO_LOGIN",
        "NEED_LOGIN", "FAIL_SYS_TOKEN", "未登录", "请登录", "先登录", "要求登录",
        "登录失效", "登录态失效", "登录态已失效", "会话过期", "验证身份", "CODE=605",
    ))


class ApiSearch:
    def __init__(self, platform, module, keyword, force_login=False):
        self.platform, self.module, self.keyword = platform, module, keyword
        self.recovered = False
        try:
            saved = module.ensure_cookie()
            cookie = saved[0]
        except Exception as error:
            if platform != "jingdong" or not isinstance(error, module.JdBlocked):
                raise
            cookie = ""
        if force_login or not has_identity(platform, cookie):
            cookie = wait_for_login(module, platform, force=force_login or bool(cookie))
        self.set_cookie(cookie)

    def set_cookie(self, cookie):
        if not has_identity(self.platform, cookie):
            raise RuntimeError(f"{NAMES[self.platform]}登录未完成；已停止采集")
        self.cookie = cookie
        self.cookie_map = cookie_values(cookie)
        # 使用刚登录取得的 Cookie，不能再次读回 config.py 中已经过期的值。
        self.module.COOKIE = cookie
        if self.platform == "taobao":
            self.spider = self.module.TaobaoMTop()

    def search(self, page):
        try:
            if self.platform == "jingdong":
                return self.module.search_page(self.keyword, page, self.cookie, self.cookie_map)
            payload = self.spider.search(self.keyword, page)
            ret = payload.get("ret") or []
            message = " ".join(map(str, ret)) if isinstance(ret, list) else str(ret)
            data = payload.get("data") or {}
            login_url = str(data.get("url") or "") if isinstance(data, dict) else ""
            if login_error(message) or ("login" in login_url.lower() and "taobao.com" in login_url.lower()):
                raise LoginRequired(message or "淘宝要求登录")
            if ret and "SUCCESS" not in message.upper():
                raise RuntimeError(f"淘宝搜索失败：{message}")
            return self.module.collect_products(payload)
        except Exception as error:
            if login_error(error):
                raise LoginRequired(str(error)) from error
            raise

    def fetch(self, page):
        try:
            return self.search(page)
        except LoginRequired:
            if self.recovered:
                raise
            self.recovered = True
            print(f"{NAMES[self.platform]}会话已失效或要求登录，打开页面等待手动登录后重试当前页。")
            self.set_cookie(wait_for_login(self.module, self.platform, force=True))
            return self.search(page)
