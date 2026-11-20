import importlib.util
from pathlib import Path
import sys
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'app'))
from search_collect import collect, xianyu_search_fetch


class RecoveryTests(unittest.TestCase):
    def browser(self, responses):
        browser = Mock()
        browser.search.side_effect = responses
        browser.has_login_prompt.return_value = False
        browser.ensure_login.return_value = True
        return browser

    def test_expired_server_session_logs_in_and_retries_same_page(self):
        browser = self.browser([([], '登录态已失效：FAIL_SYS_SESSION_EXPIRED'),
                                ([{'itemId': '1'}], '')])
        rows = xianyu_search_fetch(browser, '键盘')(3)
        self.assertEqual(rows, [{'itemId': '1'}])
        browser.ensure_login.assert_called_once_with(force=True)
        self.assertEqual([c.args for c in browser.search.call_args_list], [('键盘', 3), ('键盘', 3)])

    def test_later_login_failure_preserves_previous_page(self):
        browser = self.browser([([{'itemId': '1'}], ''), ([], '闲鱼要求登录')])
        browser.ensure_login.return_value = False
        rows, state = collect(xianyu_search_fetch(browser, '键盘'), lambda x: x['itemId'], 2)
        self.assertEqual(rows, [{'itemId': '1'}])
        self.assertFalse(state['complete'])
        self.assertIn('登录未完成', state['reason'])

    def test_no_relogin_loop_after_server_rejects_new_session(self):
        browser = self.browser([([], 'SESSION_EXPIRED'), ([], 'SESSION_EXPIRED')])
        rows, state = collect(xianyu_search_fetch(browser, '键盘'), lambda x: x['itemId'], 2)
        self.assertEqual(rows, [])
        self.assertIn('SESSION_EXPIRED', state['reason'])
        browser.ensure_login.assert_called_once()
        self.assertEqual(browser.search.call_count, 2)

    def test_captcha_is_not_treated_as_expired_login(self):
        browser = self.browser([([], '需要人机验证：FAIL_SYS_USER_VALIDATE')])
        with self.assertRaisesRegex(RuntimeError, '人机验证'):
            xianyu_search_fetch(browser, '键盘')(1)
        browser.ensure_login.assert_not_called()


class LoginPageTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        vendor = Path(__file__).resolve().parents[1] / 'vendor/xianyu'
        with patch.dict(sys.modules):
            sys.modules.pop('config', None)
            sys.path.insert(0, str(vendor))
            try:
                spec = importlib.util.spec_from_file_location('xianyu_login_tests', vendor / 'xianyu.py')
                cls.module = importlib.util.module_from_spec(spec)
                spec.loader.exec_module(cls.module)
            finally:
                sys.path.remove(str(vendor))

    def browser(self):
        browser = self.module.XianyuBrowser.__new__(self.module.XianyuBrowser)
        browser.context = Mock()
        browser.page = Mock()
        browser.has_login_prompt = Mock(return_value=False)
        return browser

    def test_first_login_opens_page_and_accepts_manual_login(self):
        browser = self.browser()
        browser.context.cookies.return_value = [{'name': 'unb'}]
        with patch('builtins.input', return_value=''), patch.object(self.module.time, 'sleep'):
            self.assertTrue(browser.ensure_login())
        self.assertIn(self.module.SEARCH_HOST, browser.page.goto.call_args_list[0].args[0])
        browser.page.bring_to_front.assert_called_once()
        browser.context.clear_cookies.assert_not_called()
        self.assertTrue(browser.logged_in_hint)

    def test_forced_login_clears_stale_cookie_and_rejects_empty_login(self):
        browser = self.browser()
        browser.context.cookies.return_value = []
        with patch('builtins.input', return_value=''), patch.object(self.module.time, 'sleep'):
            self.assertFalse(browser.ensure_login(force=True))
        browser.context.clear_cookies.assert_called_once()
        self.assertFalse(browser.logged_in_hint)

    def test_server_login_error_cannot_fall_back_to_stale_dom_cards(self):
        browser = self.browser()
        browser._responses = []
        browser._wait_search_payload = Mock(return_value={
            'ret': ['FAIL_SYS_SESSION_EXPIRED::Session过期'], 'data': {}})
        browser.search_dom = Mock(return_value=[{'itemId': 'old'}])
        rows, problem = browser.search('键盘', 2)
        self.assertEqual(rows, [])
        self.assertIn('SESSION_EXPIRED', problem)
        browser.search_dom.assert_not_called()


if __name__ == '__main__':
    unittest.main()
