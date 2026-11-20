import sys
from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import MagicMock, Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'app'))
from manual_login import ApiSearch, LoginRequired, has_identity, wait_for_login
from search_collect import collect


TB_COOKIE = 'unb=user; cookie2=fresh'
JD_COOKIE = 'pt_key=fresh; pt_pin=user; __jdu=device; 3AB9D23F7A4B3CSS=device-token'


class ManualLoginTests(unittest.TestCase):
    def module(self, saved, cookies):
        module = SimpleNamespace(
            USER_DATA_DIR='dedicated-test-profile', JD_COOKIE_KEYS=('__jdu', '3AB9D23F7A4B3CSS'),
            _apply_browser_env=lambda options: options, parse_proxy=lambda: None,
            sync_playwright=MagicMock(), ensure_cookie=Mock(return_value=(saved, {}, 'saved')),
            JdBlocked=type('JdBlocked', (RuntimeError,), {}),
            search_page=Mock(), TaobaoMTop=Mock(), collect_products=Mock())
        context = module.sync_playwright.return_value.__enter__.return_value.chromium.launch_persistent_context.return_value
        page = Mock()
        page.is_closed.return_value = False
        context.pages = [page]
        context.cookies.side_effect = cookies
        return module, context, page

    def test_anonymous_cookies_are_not_account_login(self):
        self.assertFalse(has_identity('jingdong', '__jdu=device; 3AB9D23F7A4B3CSS=token'))
        self.assertFalse(has_identity('taobao', 'cookie2=anonymous; _m_h5_tk=token'))
        self.assertTrue(has_identity('jingdong', JD_COOKIE))
        self.assertTrue(has_identity('jingdong', 'thor=account-session; pin=user'))
        self.assertFalse(has_identity('jingdong', 'pin=user; __jdu=device'))
        self.assertTrue(has_identity('taobao', TB_COOKIE))

    def test_menu_login_check_uses_platform_and_rejects_guest_cookie(self):
        from common.env_check import login_state
        harvest = Mock(return_value=[{'name': '__jdu', 'value': 'device'}])
        logged, _, _ = login_state('jingdong', harvest)
        self.assertFalse(logged)
        harvest.assert_called_once_with('jingdong')

    def test_both_platforms_keep_waiting_after_enter_without_login(self):
        for platform, cookie in [('taobao', TB_COOKIE), ('jingdong', JD_COOKIE)]:
            cookies = [{'name': part.split('=')[0], 'value': part.split('=')[1]} for part in cookie.split('; ')]
            module, context, page = self.module('', [[{'name': '__jdu', 'value': 'anonymous'}], cookies])
            with patch('builtins.input', side_effect=['', '']) as prompt:
                result = wait_for_login(module, platform)
            self.assertTrue(has_identity(platform, result))
            self.assertEqual(prompt.call_count, 2)
            page.bring_to_front.assert_called_once()
            context.close.assert_called_once()

    def test_cancel_does_not_return_anonymous_cookies(self):
        module, context, _ = self.module('', [])
        with patch('builtins.input', return_value='q'):
            self.assertEqual(wait_for_login(module, 'taobao'), '')
        context.cookies.assert_not_called()
        context.close.assert_called_once()

    def test_missing_taobao_account_opens_login_before_search(self):
        module, _, _ = self.module('cookie2=anonymous', [])
        with patch('manual_login.wait_for_login', return_value=TB_COOKIE) as login:
            search = ApiSearch('taobao', module, '键盘')
        login.assert_called_once_with(module, 'taobao', force=True)
        self.assertEqual(module.COOKIE, TB_COOKIE)
        module.TaobaoMTop.assert_called_once()

    def test_missing_jd_cookie_uses_new_login_result_directly(self):
        module, _, _ = self.module('', [])
        module.ensure_cookie.side_effect = module.JdBlocked('missing cookie')
        with patch('manual_login.wait_for_login', return_value=JD_COOKIE) as login:
            search = ApiSearch('jingdong', module, '键盘')
        login.assert_called_once_with(module, 'jingdong', force=False)
        self.assertEqual(module.ensure_cookie.call_count, 1)
        self.assertEqual(search.cookie_map['pt_key'], 'fresh')

    def test_jd_expired_config_is_replaced_before_retry_same_page(self):
        module, _, _ = self.module(JD_COOKIE.replace('fresh', 'expired'), [])
        module.search_page.side_effect = [module.JdBlocked('code=605，京东要求先验证身份'), [{'skuId': '1'}]]
        search = ApiSearch('jingdong', module, '键盘')
        with patch('manual_login.wait_for_login', return_value=JD_COOKIE) as login:
            self.assertEqual(search.fetch(2), [{'skuId': '1'}])
        login.assert_called_once_with(module, 'jingdong', force=True)
        args = module.search_page.call_args.args
        self.assertEqual(args[:3], ('键盘', 2, JD_COOKIE))
        self.assertEqual(args[3]['pt_key'], 'fresh')
        self.assertEqual(module.COOKIE, JD_COOKIE)

    def test_taobao_login_payload_rebuilds_search_client(self):
        module, _, _ = self.module(TB_COOKIE.replace('fresh', 'expired'), [])
        old_spider, new_spider = Mock(), Mock()
        old_spider.search.return_value = {'ret': ['FAIL_SYS_SESSION_EXPIRED::Session过期']}
        new_spider.search.return_value = {'ret': ['SUCCESS::调用成功'], 'data': {}}
        module.TaobaoMTop.side_effect = [old_spider, new_spider]
        module.collect_products.return_value = [{'item_id': '1'}]
        search = ApiSearch('taobao', module, '键盘')
        with patch('manual_login.wait_for_login', return_value=TB_COOKIE):
            self.assertEqual(search.fetch(3), [{'item_id': '1'}])
        old_spider.search.assert_called_once_with('键盘', 3)
        new_spider.search.assert_called_once_with('键盘', 3)
        module.collect_products.assert_called_once_with(new_spider.search.return_value)

    def test_repeated_login_rejection_is_bounded_and_keeps_prior_rows(self):
        module, _, _ = self.module(JD_COOKIE, [])
        module.search_page.side_effect = [[{'skuId': '1'}], module.JdBlocked('NOT_LOGIN'), module.JdBlocked('NOT_LOGIN')]
        search = ApiSearch('jingdong', module, '键盘')
        with patch('manual_login.wait_for_login', return_value=JD_COOKIE) as login:
            rows, state = collect(search.fetch, lambda item: item['skuId'], 2)
        self.assertEqual(rows, [{'skuId': '1'}])
        self.assertFalse(state['complete'])
        self.assertIn('NOT_LOGIN', state['reason'])
        login.assert_called_once()

    def test_taobao_rate_limit_does_not_prompt_relogin(self):
        module, _, _ = self.module(TB_COOKIE, [])
        module.TaobaoMTop.return_value.search.return_value = {'ret': ['FAIL_SYS_FLOW_LIMIT::限流']}
        search = ApiSearch('taobao', module, '键盘')
        with patch('manual_login.wait_for_login') as login:
            with self.assertRaisesRegex(RuntimeError, '限流'):
                search.fetch(1)
        login.assert_not_called()


if __name__ == '__main__':
    unittest.main()
