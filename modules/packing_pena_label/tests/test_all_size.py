# -*- coding: utf-8 -*-
"""テスラ全サイズ画面の検証（副番の採番が VBA と一致すること）。"""

import os
import shutil
import sys
import tempfile
import unittest

from modules.packing_pena_label.app.repositories.sqlite_store import Store
from modules.packing_pena_label.app.services.all_size import (AllSizeService, build_header_cells,
                                   build_sub_numbers, print_area, sheet_names,
                                   COL_KEN, COL_SIZE, COL_WT, ROW_H1, ROW_H2,
                                   SUB_COL_L, SUB_COL_R)

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
KATABAN_CSV = os.path.join(ROOT, "data", "型番.csv")


class TestSubNumbers(unittest.TestCase):
    """VBA 入力() にハードコードされていた副番と突き合わせる。"""

    def test_two_digit_sheet1(self):
        c = build_sub_numbers(1, three_digit=False)
        self.assertEqual(c[(3, SUB_COL_L)], "-0101")
        self.assertEqual(c[(83, SUB_COL_L)], "-0109")
        self.assertEqual(c[(3, SUB_COL_R)], "-0110")
        self.assertEqual(c[(83, SUB_COL_R)], "-0118")
        self.assertEqual(c[(96, SUB_COL_L)], "-0119")
        self.assertEqual(c[(176, SUB_COL_L)], "-0127")
        self.assertEqual(c[(96, SUB_COL_R)], "-0128")
        self.assertEqual(c[(176, SUB_COL_R)], "-0136")

    def test_two_digit_sheet2(self):
        c = build_sub_numbers(2, three_digit=False)
        self.assertEqual(c[(3, SUB_COL_L)], "-0201")
        self.assertEqual(c[(176, SUB_COL_R)], "-0236")

    def test_three_digit_sheet1(self):
        c = build_sub_numbers(1, three_digit=True)
        self.assertEqual(c[(3, SUB_COL_L)], "-1101")
        self.assertEqual(c[(83, SUB_COL_L)], "-1109")
        self.assertEqual(c[(3, SUB_COL_R)], "-1110")
        self.assertEqual(c[(83, SUB_COL_R)], "-1118")
        self.assertEqual(c[(96, SUB_COL_L)], "-1201")
        self.assertEqual(c[(176, SUB_COL_R)], "-1218")

    def test_three_digit_sheet2(self):
        c = build_sub_numbers(2, three_digit=True)
        self.assertEqual(c[(3, SUB_COL_L)], "-2101")
        self.assertEqual(c[(83, SUB_COL_R)], "-2118")
        self.assertEqual(c[(96, SUB_COL_L)], "-2201")
        self.assertEqual(c[(176, SUB_COL_R)], "-2218")

    def test_counts(self):
        for sheet in (1, 2):
            for three in (True, False):
                c = build_sub_numbers(sheet, three)
                self.assertEqual(len(c), 36, "副番は 18 行 × 左右 2 面 = 36")
                self.assertEqual(len(set(c.values())), 36, "副番が重複している")


class TestHeaderCells(unittest.TestCase):
    def test_two_digit_same_size_both_blocks(self):
        c = build_header_cells("X1", "W111111", "10", "10", three_digit=False)
        self.assertEqual(c[(ROW_H1, COL_SIZE)], "X1")
        self.assertEqual(c[(ROW_H2, COL_SIZE)], "X1")
        self.assertEqual(c[(ROW_H1, COL_KEN)], "W111111")

    def test_three_digit_appends_suffix(self):
        c = build_header_cells("X1", "W111111", "10", "11", three_digit=True)
        self.assertEqual(c[(ROW_H1, COL_SIZE)], "X1-1")
        self.assertEqual(c[(ROW_H2, COL_SIZE)], "X1-2")
        self.assertEqual(c[(ROW_H2, COL_WT)], "11")

    def test_empty_weight_clears_size_and_ken(self):
        """VBA: 重量空白時はサイズ・検番を消す。"""
        c = build_header_cells("X1", "W111111", "10", "", three_digit=True)
        self.assertEqual(c[(ROW_H1, COL_KEN)], "W111111")
        self.assertEqual(c[(ROW_H2, COL_KEN)], "")
        self.assertEqual(c[(ROW_H2, COL_SIZE)], "")


class TestPrintArea(unittest.TestCase):
    def test_three_digit_both(self):
        c = {(ROW_H1, COL_WT): "10", (ROW_H2, COL_WT): "11"}
        self.assertEqual(print_area(c, True, "x"), "$A$1:$BJ$186")

    def test_three_digit_bottom_only(self):
        c = {(ROW_H1, COL_WT): "", (ROW_H2, COL_WT): "11"}
        self.assertEqual(print_area(c, True, "x"), "$A$94:$BJ$186")

    def test_three_digit_top_only(self):
        c = {(ROW_H1, COL_WT): "10", (ROW_H2, COL_WT): ""}
        self.assertEqual(print_area(c, True, "x"), "$A$1:$BJ$93")

    def test_two_digit_short_sizes(self):
        for name in ("1.0×63.0×Coil", "1.0×73×Coil", "0.6×82.5×Coil"):
            self.assertEqual(print_area({}, False, name), "$A$1:$BJ$93")

    def test_two_digit_other_sizes(self):
        self.assertEqual(print_area({}, False, "1.0×33×Coil"), "$A$1:$BJ$186")


class TestService(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.store = Store(os.path.join(self.tmp, "s.sqlite3"))
        self.svc = AllSizeService(self.store, KATABAN_CSV)

    def tearDown(self):
        self.store.close()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_master_loaded(self):
        self.assertEqual(self.svc.master.count(), 7)
        self.assertEqual(self.svc.master.by_name("1.0×53.5×Coil")["kataban"],
                         "BJB7606500QR")

    def test_sheet_names(self):
        self.assertEqual(sheet_names("1.0×53.5×Coil"),
                         ("1.0×53.5×Coil1", "1.0×53.5×Coil2"))

    def test_submit_writes_both_sheets(self):
        r = self.svc.submit("1.0×53.5×Coil", "w111111", False,
                            "10", "12", "", "")
        self.assertTrue(r["ok"], r["message"])
        self.assertEqual(len(r["sheets"]), 2)
        s1 = self.svc.get_sheet("1.0×53.5×Coil1")
        self.assertEqual(s1["kensaNo"], "W111111")
        self.assertEqual(s1["weightTop"], "10")
        self.assertEqual(s1["cells"]["4,4"], "BJB7606500QR")   # 型番入力()
        self.assertEqual(s1["cells"]["97,4"], "BJB7606500QR")

    def test_submit_rejects_bad_kensa_no(self):
        r = self.svc.submit("1.0×53.5×Coil", "1111111", False, "10", "12", "", "")
        self.assertFalse(r["ok"])
        self.assertIn("7桁目に数字", r["message"])

    def test_submit_rejects_unknown_size(self):
        r = self.svc.submit("NOPE", "W111111", False, "10", "12", "", "")
        self.assertFalse(r["ok"])
        self.assertIn("おちつけシートがねぇぞ", r["message"])

    def test_submit_requires_digit_choice(self):
        r = self.svc.submit("1.0×53.5×Coil", "W111111", None, "10", "12", "", "")
        self.assertFalse(r["ok"])
        self.assertIn("2ケタ、3ケタどちらにもチェックが入っていません", r["message"])

    def test_print_targets_requires_weight(self):
        self.svc.prepare_sheets("1.0×53.5×Coil")
        r = self.svc.print_targets("1.0×53.5×Coil")
        self.assertFalse(r["ok"])
        self.assertEqual(r["message"], "重量インプットがありません")

        self.svc.submit("1.0×53.5×Coil", "W111111", False, "10", "12", "", "")
        r = self.svc.print_targets("1.0×53.5×Coil")
        self.assertTrue(r["ok"])
        self.assertEqual(len(r["targets"]), 2)

    def test_prepare_clears_previous(self):
        self.svc.submit("1.0×53.5×Coil", "W111111", False, "10", "12", "", "")
        self.assertTrue(self.svc.get_sheet("1.0×53.5×Coil1"))
        self.svc.prepare_sheets("1.0×53.5×Coil")
        self.assertEqual(self.svc.get_sheet("1.0×53.5×Coil1"), {})


if __name__ == "__main__":
    unittest.main()


class TestPrintOnlyInsideArea(unittest.TestCase):
    """全サイズの印刷は VBA の印刷範囲の外を刷らない(統合 1.2.31・現場の指示)。

    2ケタで条の少ないサイズ(63/73/82.5)は上のブロック 18 枚だけ。以前は上下 36 枚刷り、
    存在しないコイル番号(-0119〜-0136)のラベルが出た。"""

    def test_印刷範囲に入るブロック(self):
        from modules.packing_pena_label.app.services.all_size import blocks_in_area
        self.assertEqual(blocks_in_area("$A$1:$BJ$186"), (True, True))
        self.assertEqual(blocks_in_area("$A$1:$BJ$93"), (True, False))
        self.assertEqual(blocks_in_area("$A$94:$BJ$186"), (False, True))
        self.assertEqual(blocks_in_area(""), (False, False))


class TestPrintViewCounts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        from .test_http import HttpTestCase
        cls.http = HttpTestCase
        HttpTestCase.setUpClass()

    @classmethod
    def tearDownClass(cls):
        cls.http.tearDownClass()

    def _labels(self, combo):
        h = self.http()
        h.post("/api/all-size/prepare", {"comboText": combo})
        _, j = h.post("/api/all-size/submit", {"comboText": combo, "kensaNo": "W123456",
                                               "threeDigit": False, "w1": "10", "w2": "12"})
        self.assertTrue(j["ok"], j)
        _, html = h.get("/all-size/print?combo=" + __import__("urllib.parse").parse.quote(combo))
        return html.count('<div class="lbl">'), html

    def test_条の少ないサイズは上のブロックだけ(self):
        n, html = self._labels("1.0×63.0×Coil")
        self.assertEqual(n, 36)                 # 台紙 2 枚 × 上のブロック 18 枚
        self.assertNotIn("-0119", html)
        self.assertNotIn("-0236", html)

    def test_ほかのサイズは上下とも(self):
        n, _ = self._labels("1.0×53.5×Coil")
        self.assertEqual(n, 72)                 # 台紙 2 枚 × 36 枚

    def test_重量が数でなければ決定で止める(self):
        h = self.http()
        _, j = h.post("/api/all-size/submit", {"comboText": "1.0×53.5×Coil", "kensaNo": "W123456",
                                               "threeDigit": False, "w1": "-", "w2": "12"})
        self.assertFalse(j["ok"])
        self.assertIn("数になっていません", j["message"])
