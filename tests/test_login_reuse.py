from pathlib import Path
from types import SimpleNamespace
import json
import sys
import tempfile
import time
import unittest
from unittest.mock import Mock, patch

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'app'))
from browser_search import BrowserSearch, NeedLogin


class LoginReuseTests(unittest.TestCase):
    def search(self):
        result = BrowserSearch('xianyu', SimpleNamespace(LOGIN_COOKIE_HINTS=('unb',)), '键盘')
        result.context, result.page = Mock(), Mock()
        return result

    def cookie(self, name='unb', value='saved', expires=-1):
        return {'name': name, 'value': value, 'domain': '.goofish.com', 'path': '/', 'expires': expires}

    def test_session_cookie_survives_context_restart_via_local_snapshot(self):
        with tempfile.TemporaryDirectory() as directory:
            first = self.search()
            first.state_file = Path(directory) / 'login_state.json'
            first.context.storage_state.return_value = {'cookies': [self.cookie()], 'origins': []}
            first.save_login_state()
            second = self.search()
            second.state_file = first.state_file
            second.context.cookies.return_value = []
            second.restore_login_state()
            second.context.add_cookies.assert_called_once_with([self.cookie()])

    def test_backup_cannot_overwrite_fresh_cookie_or_restore_expired_cookie(self):
        with tempfile.TemporaryDirectory() as directory:
            search = self.search()
            search.state_file = Path(directory) / 'login_state.json'
            search.state_file.write_text(json.dumps({'cookies': [self.cookie(), self.cookie('old', expires=time.time()-5)]}), encoding='utf-8')
            search.context.cookies.return_value = [self.cookie(value='fresh')]
            search.restore_login_state()
            search.context.add_cookies.assert_not_called()

    def test_automatic_recovery_does_not_clear_cookies_or_open_homepage(self):
        search = self.search()
        search.wait_results = Mock(side_effect=[NeedLogin('SESSION_EXPIRED'), [{'itemId': '1'}]])
        search.login = Mock(side_effect=lambda **kwargs: setattr(search, 'current_page', 0))
        self.assertEqual(search.fetch(1), [{'itemId': '1'}])
        search.login.assert_called_once_with(force=False, navigate=False)
        search.context.clear_cookies.assert_not_called()

    def test_restored_cookie_is_used_before_any_login_prompt(self):
        search = self.search()
        # Verify entry restores state without deciding to force login from cookie markers alone.
        manager = Mock()
        manager.start.return_value.chromium.launch_persistent_context.return_value.pages = []
        search.module = SimpleNamespace(sync_playwright=Mock(return_value=manager),
            _apply_browser_env=lambda options: options, parse_proxy=lambda: None,
            USER_DATA_DIR='test-profile')
        search.restore_login_state = Mock()
        search.login = Mock()
        with patch.dict('os.environ', {'EC_PROFILE': 'test-profile'}):
            search.__enter__()
        search.restore_login_state.assert_called_once()
        search.login.assert_not_called()
        search.context.clear_cookies.assert_not_called()


if __name__ == '__main__':
    unittest.main()
