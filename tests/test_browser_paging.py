"""真实浏览器中的分页回归测试；只使用拦截的本地测试页面，无外网/登录数据。"""
from pathlib import Path
from types import SimpleNamespace
import sys
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'app'))
from browser_search import BrowserSearch, NoNextPage


FIXTURE = '''<!doctype html><meta charset="utf-8">
<div id="cards"></div><button id="next" onclick="advance()">下一页</button>
<script>
let current=1; window.clicks=0;
function render() {
  let id=10000+current;
  document.querySelector('#cards').innerHTML = `
    <li class="gl-item" data-sku="${id}"><a href="https://item.jd.com/${id}.html">
      <span class="p-name"><em>商品标题 ${id}</em></span><span class="p-price">¥29.90</span></a></li>
    <a class="doubleCardWrapper" href="https://item.taobao.com/item.htm?id=${id}">
      <span class="title">商品标题 ${id}</span><span class="price">¥29.90</span></a>
    <a class="feeds-item" href="https://www.goofish.com/item?id=${id}">
      <span class="title">商品标题 ${id}</span><span class="price">¥29.90</span></a>`;
  document.querySelector('#next').disabled = current >= 3;
}
function advance() { window.clicks++; current++; setTimeout(render,100); }
render();
</script>'''


class BrowserPagingTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from playwright.sync_api import sync_playwright
        from common.env_check import chrome_path
        cls.pw = sync_playwright().start()
        options = {'headless': True}
        if not Path(cls.pw.chromium.executable_path).exists():
            executable = chrome_path()
            if not executable:
                cls.pw.stop()
                raise unittest.SkipTest('未安装可运行的浏览器')
            options['executable_path'] = executable
        cls.browser = cls.pw.chromium.launch(**options)

    @classmethod
    def tearDownClass(cls):
        cls.browser.close()
        cls.pw.stop()

    def setUp(self):
        self.context = self.browser.new_context()
        self.page = self.context.new_page()
        self.page.route('**/*', lambda route: route.fulfill(status=200, content_type='text/html', body=FIXTURE))
        self.navigations = []
        self.page.on('framenavigated', lambda frame: self.navigations.append(frame.url) if frame == self.page.main_frame else None)

    def tearDown(self):
        self.context.close()

    def search(self, platform):
        result = BrowserSearch(platform, SimpleNamespace(), '键盘', timeout=3)
        result.page, result.context = self.page, self.context
        return result

    def test_three_platforms_click_next_without_navigating_again(self):
        for platform in ('taobao', 'jingdong', 'xianyu'):
            with self.subTest(platform=platform), patch('browser_search.SEARCH_URLS', {platform: 'https://collector.test/search?q={}' }):
                self.navigations.clear()
                search = self.search(platform)
                ids = [search.identify(search.fetch(number)[0]) for number in (1, 2, 3)]
                self.assertEqual(ids, ['10001', '10002', '10003'])
                self.assertEqual(self.page.evaluate('window.clicks'), 2)
                self.assertEqual(len(self.navigations), 1)
                self.assertEqual(search.fetch(4), [])
                self.assertEqual(len(self.navigations), 1)

    def test_missing_next_button_stops_without_restarting_search(self):
        with patch('browser_search.SEARCH_URLS', {'xianyu': 'https://collector.test/search?q={}' }):
            search = self.search('xianyu')
            search.fetch(1)
            self.page.locator('#next').evaluate('(button) => button.remove()')
            with self.assertRaisesRegex(RuntimeError, '下一页'):
                search.fetch(2)
            self.assertEqual(len(self.navigations), 1)

    def test_click_that_does_not_update_results_never_returns_old_cards(self):
        with patch('browser_search.SEARCH_URLS', {'jingdong': 'https://collector.test/search?q={}' }):
            search = self.search('jingdong')
            search.fetch(1)
            self.page.evaluate('() => { window.advance = () => { window.clicks++; }; }')
            search.timeout = 0.9
            with self.assertRaisesRegex(RuntimeError, '上一页完全相同'):
                search.fetch(2)
            self.assertEqual(search.current_page, 1)
            self.assertEqual(len(self.navigations), 1)

    def test_out_of_order_pages_are_rejected_without_navigation(self):
        with patch('browser_search.SEARCH_URLS', {'taobao': 'https://collector.test/search?q={}' }):
            search = self.search('taobao')
            with self.assertRaisesRegex(RuntimeError, '按顺序'):
                search.fetch(2)
            self.assertEqual(self.navigations, [])

    def test_next_page_reads_browser_response_without_direct_api_call(self):
        payload = {'code': '0', 'data': {'wareList': [{'skuId': '20002', 'wareName': '第二页商品'}]}}
        self.page.route('**/api?**', lambda route: route.fulfill(json=payload))
        with patch('browser_search.SEARCH_URLS', {'jingdong': 'https://collector.test/search?q={}' }):
            search = self.search('jingdong')
            search.module = SimpleNamespace(explain_failure=lambda payload, name: None)
            self.page.on('response', search.on_response)
            search.fetch(1)
            self.page.evaluate("() => { window.advance = () => { window.clicks++; fetch('/api?functionId=pc_search_searchWare'); }; }")
            rows = search.fetch(2)
            self.assertEqual(rows, payload['data']['wareList'])
            self.assertEqual(self.page.evaluate('window.clicks'), 1)
            self.assertEqual(len(self.navigations), 1)


if __name__ == '__main__':
    unittest.main()
