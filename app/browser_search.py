"""搜索一次后在当前结果页点击下一页；只读取列表响应和商品卡片。"""
import json
import os
import re
import time
from urllib.parse import quote, urlparse

from manual_login import HOME_URLS, NAMES, has_identity, login_error


SEARCH_URLS = {
    "taobao": "https://s.taobao.com/search?q={}",
    "jingdong": "https://search.jd.com/Search?keyword={}&enc=utf-8",
    "xianyu": "https://www.goofish.com/search?q={}",
}
NEXT_SELECTORS = (
    '.ant-pagination-next', '.next-pagination-item.next-next',
    '.pagination-next', '.pagination .next', '.ui-pager-next', '[class*="pagination-next"]',
    '[class*="Pagination"] [class*="next"]', '[class*="pagination"] [class*="next"]',
    'a.pn-next', '.next-btn', 'button[aria-label="Next page"]',
    '[aria-label="下一页"]', '[title="下一页"]',
)


class NoNextPage(RuntimeError):
    pass


class NeedLogin(RuntimeError):
    pass


def click_next_page(page):
    """只点击分页控件，绝不构造下一页 URL 或点击商品链接。"""
    page.evaluate('window.scrollTo(0, document.documentElement.scrollHeight)')
    candidates = [page.locator(selector) for selector in NEXT_SELECTORS]
    candidates += [page.get_by_role('button', name=re.compile(r'^下一页\s*[>›»]?$')),
                   page.get_by_role('link', name=re.compile(r'^下一页\s*[>›»]?$'))]
    found_disabled = False
    for candidate in candidates:
        for index in range(candidate.count()):
            control = candidate.nth(index)
            if not control.is_visible():
                continue
            classes = control.get_attribute('class') or ''
            disabled = (control.get_attribute('aria-disabled') == 'true'
                        or control.get_attribute('disabled') is not None
                        or bool(re.search(r'disabled|disable(?:\s|$)', classes, re.I)))
            if disabled or not control.is_enabled():
                found_disabled = True
                continue
            control.click(timeout=10000)
            return
    if found_disabled:
        raise NoNextPage('当前搜索结果已到最后一页（下一页不可用）')
    raise RuntimeError('未找到可用的「下一页」控件，已停止；不会重新搜索或拼接翻页地址')


DOM_CARDS = r'''platform => {
  const rules = {
    jingdong: ['li.gl-item[data-sku]', '[data-sku]'],
    taobao: ['a[class*="doubleCardWrapper"]', 'a[href*="item.taobao.com/item.htm"]', 'a[href*="detail.tmall.com/item.htm"]'],
    xianyu: ['a[href*="/item?"]', 'a[href*="/item?id="]']
  };
  const output = [], seen = new Set();
  for (const selector of rules[platform]) {
    for (const link of document.querySelectorAll(selector)) {
      const a = link.matches('a') ? link : link.querySelector('a[href]');
      const href = a ? a.href : '';
      const match = /[?&]id=(\d+)/.exec(href) || /\/(\d+)\.html/.exec(href);
      const id = link.getAttribute('data-sku') || (match && match[1]);
      if (!id || seen.has(id)) continue;
      let card = link;
      if (platform === 'jingdong') card = link.closest('li.gl-item') || link;
      if (platform === 'taobao') card = link.closest('[class*="doubleCardWrapper"]') || link;
      if (platform === 'xianyu') card = link.closest('[class*="feeds-item"]') || link;
      const lines = (card.innerText || '').split('\n').map(x=>x.trim()).filter(Boolean);
      const image = card.querySelector('img');
      const titleNode = card.querySelector('[class*="title"], .p-name em, [class*="Title"]');
      const title = (titleNode && titleNode.textContent.trim()) || (image && image.alt) ||
                    lines.find(x=>x.length > 4 && !/^[¥￥\d.,\s]+$/.test(x)) || '';
      const priceNode = card.querySelector('.p-price, [class*="price"], [class*="Price"]');
      const priceText = priceNode ? priceNode.textContent : (lines.find(x=>/[¥￥]\s*\d/.test(x)) || '');
      const price = (priceText.match(/[\d,]+(?:\.\d+)?/) || [''])[0];
      if (!title || !price) continue;
      const shopNode = card.querySelector('.p-shop, [class*="shopName"], [class*="shop-name"], [class*="seller"]');
      const shop = shopNode ? shopNode.textContent.trim() : '';
      let pic = image && (image.getAttribute('data-lazy-img') || image.getAttribute('data-src') || image.currentSrc || image.src);
      if (pic && pic.startsWith('//')) pic = 'https:' + pic;
      if (platform === 'jingdong') output.push({skuId:id, wareName:title, jdPrice:price, shopName:shop, imageurl:pic, source:'搜索页DOM'});
      if (platform === 'taobao') output.push({item_id:id, title, price, shopInfo:{title:shop}, auctionURL:href, pic_path:pic, source:'搜索页DOM'});
      if (platform === 'xianyu') output.push({itemId:id, title, price, seller:shop, picUrl:pic, source:'搜索页DOM'});
      seen.add(id);
    }
    if (output.length) break;
  }
  return output;
}'''


class BrowserSearch:
    def __init__(self, platform, module, keyword, force_login=False, timeout=45):
        self.platform, self.module, self.keyword = platform, module, keyword
        self.force_login, self.timeout = force_login, timeout
        self.current_page, self.pending_page = 0, 0
        self.responses, self.previous_ids = [], set()
        self.recovered = False

    def __enter__(self):
        self.playwright = self.module.sync_playwright().start()
        options = self.module._apply_browser_env(dict(
            channel='chrome', headless=False, viewport={'width': 1440, 'height': 900},
            locale='zh-CN', timezone_id='Asia/Shanghai',
            args=['--disable-blink-features=AutomationControlled', '--mute-audio'],
        ))
        proxy = self.module.parse_proxy()
        if proxy:
            options['proxy'] = proxy
        profile = os.environ.get('EC_PROFILE') or getattr(self.module, 'PROFILE_DIR', self.module.USER_DATA_DIR)
        try:
            self.context = self.playwright.chromium.launch_persistent_context(profile, **options)
            self.page = self.context.pages[0] if self.context.pages else self.context.new_page()
            self.page.on('response', self.on_response)
            if self.force_login or not self.has_login():
                self.login(force=self.force_login)
            return self
        except BaseException:
            if hasattr(self, 'context'):
                self.context.close()
            self.playwright.stop()
            raise

    def __exit__(self, *args):
        try:
            self.context.close()
        finally:
            self.playwright.stop()

    def has_login(self):
        cookies = self.context.cookies()
        if self.platform == 'xianyu':
            return any(c.get('name') in self.module.LOGIN_COOKIE_HINTS and c.get('value') for c in cookies)
        text = '; '.join(f"{c['name']}={c['value']}" for c in cookies if c.get('name') and c.get('value'))
        return has_identity(self.platform, text)

    def login(self, force=False):
        name = NAMES.get(self.platform, '闲鱼')
        home = HOME_URLS.get(self.platform, 'https://www.goofish.com')
        if force:
            self.context.clear_cookies()
        self.page.goto(home, wait_until='domcontentloaded', timeout=60000)
        self.page.bring_to_front()
        try:
            self.page.get_by_text(re.compile(r'^(亲，请登录|你好，请登录|登录|登录/注册|立即登录|请登录)$')).first.click(timeout=5000)
        except Exception:
            pass
        print(f'请在浏览器中手动登录{name}；登录完成后回到终端按 Enter，输入 q 取消。')
        while True:
            try:
                answer = input('等待登录（Enter 确认，q 取消）：').strip().lower()
            except EOFError:
                raise RuntimeError('无法读取输入，登录尚未确认')
            if answer == 'q' or self.page.is_closed():
                raise RuntimeError('已取消登录')
            if self.has_login():
                break
            print('还未检测到账号登录，请完成登录后再按 Enter。')
        # 新登录后只打开一次搜索结果，再逐次点击下一页。
        self.current_page = self.pending_page = 0
        self.previous_ids = set()

    def on_response(self, response):
        url = response.url.lower()
        relevant = {'xianyu': 'idlemtopsearch.pc.search' in url,
                    'jingdong': 'pc_search_searchware' in url,
                    'taobao': ('mtop' in url and ('search' in url or 'recommend' in url))}
        if not relevant[self.platform]:
            return
        try:
            body = response.text()
            if body.lstrip().startswith('{'):
                payload = json.loads(body)
            else:
                match = re.search(r'^[^(]*\((.*)\)\s*;?\s*$', body, re.S)
                if not match:
                    return
                payload = json.loads(match.group(1))
            self.responses.append(payload)
        except Exception:
            pass

    def parse_payload(self, payload):
        if self.platform == 'xianyu':
            kind, problem = self.module.mtop_problem(payload)
            if kind in ('session', 'login'):
                raise NeedLogin(problem)
            if kind == 'token':
                return []  # 页面自身会完成令牌握手，继续等待后续成功响应。
            if problem:
                raise RuntimeError(problem)
            return self.module.extract_items(payload, self.keyword)
        if self.platform == 'jingdong':
            problem = self.module.explain_failure(payload, '搜索页')
            if problem:
                if login_error(problem):
                    raise NeedLogin(problem)
                raise RuntimeError(problem)
            return (payload.get('data') or {}).get('wareList') or []
        ret = payload.get('ret') or []
        message = ' '.join(map(str, ret)) if isinstance(ret, list) else str(ret)
        data = payload.get('data') or {}
        redirect = str(data.get('url') or '') if isinstance(data, dict) else ''
        if 'login' in redirect.lower() and 'taobao.com' in redirect.lower():
            raise NeedLogin('淘宝要求登录')
        if 'FAIL_SYS_TOKEN' in message.upper():
            return []
        if login_error(message):
            raise NeedLogin(message)
        if ret and 'SUCCESS' not in message.upper():
            raise RuntimeError(f'淘宝搜索失败：{message}')
        return self.module.collect_products(payload)

    def identify(self, item):
        return str(item.get({'taobao': 'item_id', 'xianyu': 'itemId', 'jingdong': 'skuId'}[self.platform])
                   or item.get('wareId') or '')

    def wait_results(self):
        deadline = time.monotonic() + self.timeout
        dom_ids, stable = set(), 0
        while time.monotonic() < deadline:
            self.page.wait_for_timeout(400)
            items = []
            while self.responses:
                items.extend(self.parse_payload(self.responses.pop(0)))
            from_response = bool(items)
            if not items:
                items = self.page.evaluate(DOM_CARDS, self.platform)
            items = list({self.identify(item): item for item in items if self.identify(item)}.values())
            ids = {self.identify(item) for item in items if self.identify(item)}
            if ids and ids != self.previous_ids:
                if from_response:
                    return items
                stable = stable + 1 if ids == dom_ids else 0
                dom_ids = ids
                if stable >= 2:
                    return items
            try:
                host = urlparse(self.page.url).hostname or ''
                if host.startswith(('login.', 'passport.')):
                    raise NeedLogin('当前页面要求登录')
                dialogs = self.page.locator('iframe[src*="passport"], iframe[src*="login"]')
                if any(dialogs.nth(i).is_visible() for i in range(dialogs.count())):
                    raise NeedLogin('当前页面显示登录框')
            except NeedLogin:
                raise
            except Exception:
                pass
            # 只滚动已有结果以加载商品和分页控件，不再次提交搜索。
            self.page.mouse.wheel(0, 700)
        raise RuntimeError('等待结果页更新超时或商品与上一页完全相同，已停止以避免重复采集')

    def fetch(self, page_index):
        if page_index != self.current_page + 1:
            raise RuntimeError('必须按顺序从当前结果点击下一页')
        try:
            self.responses.clear()
            if self.current_page == 0:
                print(f'搜索一次：{self.keyword}')
                self.page.goto(SEARCH_URLS[self.platform].format(quote(self.keyword)),
                               wait_until='domcontentloaded', timeout=60000)
            else:
                print(f'在当前搜索结果中点击下一页 → 第 {page_index} 页')
                click_next_page(self.page)
            self.pending_page = page_index
            items = self.wait_results()
            self.current_page = page_index
            self.previous_ids = {self.identify(item) for item in items}
            return items
        except NoNextPage:
            return []
        except NeedLogin:
            if self.recovered:
                raise
            self.recovered = True
            print('服务器要求重新登录。登录后恢复搜索，再逐页点击回到中断位置。')
            self.login(force=True)
            for previous in range(1, page_index):
                self.fetch(previous)
            return self.fetch(page_index)
