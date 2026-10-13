# -*- coding: utf-8 -*-
"""丸めの再点検(統合 1.2.12)

- VBA の Val・Format は有効数字 15 桁の10進で見て丸める。0.565 × 14 + … = 8.444999999999999 は 8.45
- ラベルに刷る重量は、全角・桁区切りの付いた値も CDbl と同じに読む("２０．１５" → 20.2)
- 本数は整数。11.5 は空にする(計算に 11.5 を使い、見出しは「12本」になっていた)
"""
import unittest

from modules.packing_pena_label.app.services import validation as V
from modules.packing_pena_label.app.services.label_layout import weight_text
from modules.packing_pena_label.app.services.vba_compat import fmt, val, ws_roundup


class Rounding15Test(unittest.TestCase):
    def test_15桁で見て丸める(self):
        self.assertEqual(fmt(8.444999999999999, "0.00"), "8.45")
        self.assertEqual(fmt(0.563 * 5, "0.00"), "2.82")
        self.assertEqual(fmt(304.05, "0.0"), "304.1")
        self.assertEqual(ws_roundup(1230.0000000000002, -1), 1230.0, "割り切れていれば上げない")

    def test_Valは全角数字を読まない(self):
        self.assertEqual(val("１２"), 0.0)
        self.assertEqual(val("12abc"), 12.0)


class WeightTextTest(unittest.TestCase):
    def test_全角や桁区切りも読む(self):
        self.assertEqual(weight_text("２０．１５"), "20.2")
        self.assertEqual(weight_text("1,234"), "1234.0")
        self.assertEqual(weight_text("20"), "20.0")


class CoilCountTest(unittest.TestCase):
    def test_本数は整数だけ(self):
        self.assertEqual(V.filter_coil_text("11"), "11")
        self.assertEqual(V.filter_coil_text("11.0"), "11.0")
        self.assertEqual(V.filter_coil_text("11.5"), "")
        self.assertEqual(V.filter_coil_text("abc"), "")


if __name__ == "__main__":
    unittest.main()
