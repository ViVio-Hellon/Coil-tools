# -*- coding: utf-8 -*-
"""印刷ビューは開いたらそのまま印刷の画面を出す。窓の通し試験だけはそれを止められる(統合 1.2.15)

Windows(WebView2)では開いたときの window.print() が本物の印刷の画面を出し、閉じるまでその窓が止まる。
WebDriver はその画面を押せないので、窓の通し試験は `NO_AUTOPRINT_ENV` を立てて「印刷する」ボタンから刷る。
ふだん(印が無いとき)は今までどおり開いたら印刷の画面を出す。
"""
import inspect
import os
import unittest
from unittest import mock

from modules.packing_pena_label.app.routes import pages

AUTO = "window.addEventListener('load',function(){window.print();})"


class AutoPrintTest(unittest.TestCase):
    def test_ふだんは開いたら刷る(self):
        for value in (None, "", "0"):
            with self.subTest(value=value):
                env = {k: v for k, v in os.environ.items() if k != pages.NO_AUTOPRINT_ENV}
                if value is not None:
                    env[pages.NO_AUTOPRINT_ENV] = value
                with mock.patch.dict(os.environ, env, clear=True):
                    self.assertIn(AUTO, pages._autoprint_script())

    def test_試験の印があれば開いても刷らない(self):
        with mock.patch.dict(os.environ, {pages.NO_AUTOPRINT_ENV: "1"}):
            self.assertEqual(pages._autoprint_script(), "")

    def test_4つの印刷ビューがこの1か所を使う(self):
        """風袋計算・羅列計算・ラベル台紙・全サイズの印刷ビュー。直に書くと試験の印が効かない。"""
        source = inspect.getsource(pages)
        self.assertEqual(source.count("_autoprint_script()"), 4 + 1, "呼ぶ4か所 + 定義")
        self.assertEqual(source.count(AUTO), 1, "window.print の差し込みは定義の1か所だけ")
