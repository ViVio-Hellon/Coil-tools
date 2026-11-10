"""ペナラベルの合言葉は、梱包明細の共有の管理者パスワードで確かめる(統合 1.2.75・構造点検の改善案4)

以前はプログラムに書いた決まった値で、変えられなかった。梱包明細(全ラインで共有)・資材計算(PCごと)と
3系統あり、梱包明細で変えてもペナラベルは古いままだった。
"""
import unittest

from modules.packing_details.meisai import admin_password as shared
from modules.packing_details.meisai import shared_settings
from modules.packing_details.tests import _db
from modules.packing_pena_label.app.services import settings as S

NEW = "kawatta-pass-1"


class SharedPasswordTest(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)          # 梱包明細の共有を、この試験だけの空の場所へ

    def test_変えていなければ既定で通る(self):
        self.assertTrue(S.password_ok(S.PATH_PASSWORD))
        self.assertFalse(S.password_ok(S.PATH_PASSWORD + "x"))
        self.assertFalse(S.password_ok(""))
        self.assertFalse(S.password_ok(None))

    def test_梱包明細で変えたら新しいほうだけ通る(self):
        self.assertTrue(shared.change(S.PATH_PASSWORD, NEW, NEW).ok)
        self.assertTrue(S.password_ok(NEW))
        self.assertFalse(S.password_ok(S.PATH_PASSWORD))

    def test_共有の控えが読めなければ通さない(self):
        self.assertTrue(shared.change(S.PATH_PASSWORD, NEW, NEW).ok)
        shared_settings.shared_path().write_text("{壊れている", encoding="utf-8")
        shared_settings.reset_for_tests()
        self.assertFalse(S.password_ok(S.PATH_PASSWORD))

    def test_既定の値は梱包明細と同じ(self):
        from modules.packing_details.meisai import config as details_config
        self.assertEqual(S.PATH_PASSWORD, details_config.ADMIN_PASSWORD)


if __name__ == "__main__":
    unittest.main()
