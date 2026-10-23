# -*- coding: utf-8 -*-
"""全サイズはいったん左のタブに出さない(統合 1.2.50・現場の指示: いまは使わない)

機能は消さない: /all-size と印刷ビューは開ける。config/app.json の show_all_size を true にすれば出る。
"""
import unittest

from .test_http import HttpTestCase


class AllSizeHiddenTest(HttpTestCase):
    def tearDown(self):
        self.cfg.show_all_size = False
        self.ctx.wf.cfg.show_all_size = False

    def test_既定では左のタブに全サイズを出さない_機能は残す(self):
        self.assertFalse(self.cfg.show_all_size)
        for path in ("/", "/tare", "/labels", "/settings"):
            with self.subTest(path=path):
                _, html = self.get(path)
                self.assertNotIn('/all-size"', html, "左のタブに全サイズが出ている")
        status, html = self.get("/all-size")
        self.assertEqual(status, 200, "全サイズの画面そのものは開ける")
        self.assertIn("全サイズ", html)

    def test_設定で出せる(self):
        self.cfg.show_all_size = True
        self.ctx.wf.cfg.show_all_size = True
        _, html = self.get("/")
        self.assertIn('/all-size"', html)
        self.assertIn(">全サイズ</a>", html)


if __name__ == "__main__":
    unittest.main()
