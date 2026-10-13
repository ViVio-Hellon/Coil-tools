"""数の丸め・読み方を VBA と同じにする(統合 1.2.12。丸めの再点検)

- 板厚・板幅の表示は VBA `Format` と同じ(有効数字 15 桁の10進で見て、半分は 0 から遠い側へ)
- 取り込みの数は VBA `Val` と同じ読み方("0.8mm" → 0.8、"１２" → 0、"1E-05" → 0.00001)
- 縦割数・横割数が整数でなければ切り捨てずに 0(前工程実績数は画面で入れる)
- 重量は数字だけを整数として読む(全角は半角に。0 は断る。float を通さない)
"""
from __future__ import annotations

import unittest

from modules.packing_details.meisai import import_specs, lot_repo, meisai_service
from modules.packing_details.meisai.strand_service import RefusedError


class FormatTest(unittest.TestCase):
    def test_表示は半分を0から遠い側へ(self):
        self.assertEqual([lot_repo.fmt(x, 1) for x in (104.25, 85.85, 100.05, 47.0)],
                         ["104.3", "85.9", "100.1", "47.0"])
        self.assertEqual([lot_repo.fmt(x, 3) for x in (0.8055, 1.0005, 0.8)],
                         ["0.806", "1.001", "0.800"])
        info = lot_repo.LotInfo(seizou_thickness=0.8055, seizou_width=104.25)
        self.assertEqual((info.seizou_thickness_text, info.seizou_width_text), ("0.806", "104.3"))


class ValTest(unittest.TestCase):
    def test_VBAのValと同じに読む(self):
        cases = {"0.8mm": 0.8, "1,234.5": 1.0, "１２": 0.0, "1E-05": 1e-05, " 1 2 ": 12.0,
                 "abc": 0.0, "-3.5": -3.5, ".5": 0.5, "12３": 12.0, "1.5D2": 150.0}
        for text, want in cases.items():
            with self.subTest(text=text):
                self.assertEqual(import_specs.vba_val(text), want)
                self.assertEqual(lot_repo.val(text), want)
        self.assertEqual(import_specs.to_real("0.8mm"), 0.8, "以前は読めずに 0 になっていた")
        self.assertIsNone(import_specs.to_real("abc"))

    def test_割数は整数でなければ切り捨てない(self):
        self.assertEqual([import_specs.to_whole(t) for t in ("2", "2.0", " 3 ")], [2, 2, 3])
        with self.assertLogs("meisai.import_specs", "WARNING"):
            self.assertIsNone(import_specs.to_whole("2.7"))
        self.assertEqual(import_specs.NULL_FALLBACKS[import_specs.to_whole], 0)


class WeightTest(unittest.TestCase):
    def test_数字だけを整数として読む(self):
        self.assertEqual(meisai_service.read_weights(["２５０２", "248", "250.0"], 3), [2502, 248, 250])
        for bad in ("2.5e3", "1_000", "2502.7", "-5", "+5", "1,000"):
            with self.subTest(bad=bad), self.assertRaises(RefusedError):
                meisai_service.read_weights([bad], 1)

    def test_0は断る(self):
        with self.assertRaises(RefusedError) as cm:
            meisai_service.read_weights(["0"], 1)
        self.assertIn("0", cm.exception.message)


if __name__ == "__main__":
    unittest.main()
