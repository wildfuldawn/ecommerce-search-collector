import sys
import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock
sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'app'))
from search_collect import collect, export


class SearchTests(unittest.TestCase):
    def test_duplicate_pages_continue_and_trim_exactly(self):
        fetch = Mock(side_effect=[[{'id': 1}, {'id': 1}, {'id': 2}],
                                  [{'id': 2}, {'id': 3}],
                                  [{'id': 4}, {'id': 5}]])
        rows, state = collect(fetch, lambda x: x['id'], 4)
        self.assertEqual([r['id'] for r in rows], [1, 2, 3, 4])
        self.assertEqual(fetch.call_count, 3)
        self.assertTrue(state['complete'])

    def test_blocked_later_page_keeps_previous_rows(self):
        fetch = Mock(side_effect=[[{'id': 1}], RuntimeError('blocked')])
        rows, state = collect(fetch, lambda x: x['id'], 10)
        self.assertEqual(len(rows), 1)
        self.assertIn('blocked', state['reason'])
        self.assertFalse(state['complete'])

    def test_empty_repeated_and_page_cap_stop(self):
        for data in [[[], []], [[{'id': 1}], [{'id': 1}]]]:
            fetch = Mock(side_effect=data)
            rows, state = collect(fetch, lambda x: x['id'], 10)
            self.assertFalse(state['complete'])
            self.assertLessEqual(fetch.call_count, 2)
        rows, state = collect(lambda p: [{'id': p}], lambda x: x['id'], 10, max_pages=2)
        self.assertEqual(len(rows), 2)
        self.assertIn('上限', state['reason'])

    def test_export_embeds_image_and_preserves_formula_like_title_as_text(self):
        from PIL import Image
        from io import BytesIO
        from openpyxl import load_workbook
        stream = BytesIO()
        Image.new('RGB', (20, 20), 'red').save(stream, 'PNG')
        response = Mock(content=stream.getvalue())
        client = Mock()
        client.get.return_value = response
        with tempfile.TemporaryDirectory() as folder:
            path = export('xianyu', 'test', [{'itemId': '1', 'title': '=1+1',
                          'price': '10', 'picUrl': 'https://example.com/image.png'}],
                          {'target': 1, 'actual': 1, 'pages': 1, 'complete': True, 'reason': 'ok'}, folder, session=client)
            wb = load_workbook(path)
            self.assertEqual(wb.active['B2'].value, '=1+1')
            self.assertEqual(wb.active['B2'].data_type, 's')
            self.assertEqual(len(wb.active._images), 1)
            self.assertEqual(wb.active['J2'].value, '已下载并嵌入')
            self.assertTrue(path.with_suffix('.json').exists())
            wb.close()


if __name__ == '__main__':
    unittest.main()
