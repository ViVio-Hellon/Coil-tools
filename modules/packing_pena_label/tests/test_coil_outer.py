# -*- coding: utf-8 -*-
"""コイル外径から近いチップボールを選ぶ(統合 1.2.35・現場の指示)

- 検査番号 = 仕掛ファイルのロット番号。外径 = 内径 + 2 × 肉厚(資材計算と同じ部品で読む)
- 近いチップボール = 風袋計算がチップボールごとに使う外径(1100/1000/900mm)のうち、いちばん近いもの
- 重量反映で選び、どの外径からどれを選んだかを案内する。見つからなければ選ばずに理由を出す
- 仕掛パスが実績の形式でない(LS4LOT のまま)なら何もしない
"""
import os
import sqlite3
import tempfile
import unittest
from unittest import mock

from common import ls4_source
from modules.packing_pena_label.app.services import coil_outer as CO

from .test_workflow import WorkflowTestBase


def _results_file(rows):
    """実績の形式の仕掛ファイル(いまの年月の行)。"""
    from datetime import date
    folder = tempfile.mkdtemp(prefix="ppl_coil_")
    path = os.path.join(folder, "実績_LS4.sqlite3")
    month = date.today().strftime("%Y-%m-15 08:00:00")
    with sqlite3.connect(path) as c:
        c.execute(f'CREATE TABLE "仕掛" ("{ls4_source.COL_LOT}" TEXT, "{ls4_source.COL_BOX_TATE}" INTEGER,'
                  f' "{ls4_source.COL_BOX_YOKO}" INTEGER, "{ls4_source.COL_START}" TEXT,'
                  f' "{ls4_source.COL_INNER}" REAL, "{ls4_source.COL_THICK}" REAL)')
        c.executemany('INSERT INTO "仕掛" VALUES (?,?,?,?,?,?)',
                      [(lot, 2, 12, month, inner, thick) for lot, inner, thick in rows])
    return path


class _FakeResults:
    """資材計算の coil_results の代わり(仕掛パスの設定だけ差し替える。読み方は本物)。"""

    def __init__(self, path, wanted=True, second=None):
        from modules.packing_material_calculation.coil_tool import coil_results as real
        self.real = real
        self.path = path
        self.second = second
        self._wanted = wanted

    def wanted(self):
        return self._wanted

    def sources(self):
        from pathlib import Path
        return [Path(p) for p in (self.path, self.second) if p]

    def read_rows(self, path):
        return self.real.read_rows(path)


class NearestTipTest(unittest.TestCase):
    def test_いちばん近い外径のチップボール(self):
        self.assertEqual(CO.nearest_tip(1120), "TIP1000")
        self.assertEqual(CO.nearest_tip(1060), "TIP1000")
        self.assertEqual(CO.nearest_tip(1040), "TIP950")
        self.assertEqual(CO.nearest_tip(1050), "TIP1000")     # 同じ近さなら大きいほう
        self.assertEqual(CO.nearest_tip(960), "TIP950")
        self.assertEqual(CO.nearest_tip(940), "TIP820")
        self.assertEqual(CO.nearest_tip(780), "TIP820")

    def test_資材計算の部品が使える(self):
        self.assertIsNotNone(CO._material())


class ChooseTest(unittest.TestCase):
    def setUp(self):
        CO.forget()
        self.path = _results_file([("W123456", 508.0, 260.0),     # 1028mm
                                   ("W222222", 0, 0)])

    def tearDown(self):
        CO.forget()

    def _with(self, fake):
        return mock.patch.object(CO, "_material", return_value=(fake, ls4_source))

    def test_外径から選んで案内する(self):
        with self._with(_FakeResults(self.path)):
            tip, note = CO.choose("W123456", "TIP1000")
        self.assertEqual(tip, "TIP950")
        self.assertIn("コイル外径 1028mm", note)
        self.assertIn("内径 508.0 + 2 × 肉厚 260.0", note)
        self.assertIn("チップボール 950を選びました", note)
        self.assertIn("前: チップボール 1000", note)
        self.assertIn("違うときは選び直してください", note)

    def test_選んであるとおりなら変えない(self):
        with self._with(_FakeResults(self.path)):
            tip, note = CO.choose("W123456", "TIP950")
        self.assertIsNone(tip)
        self.assertIn("選んであるとおりです", note)

    def test_見つからない_内径肉厚が無いときは選ばずに理由(self):
        with self._with(_FakeResults(self.path)):
            tip, note = CO.choose("W999999", "TIP1000")
            self.assertIsNone(tip)
            self.assertIn("ロット W999999 がありません", note)
            self.assertIn("いまの選択: チップボール 1000", note)
            tip, note = CO.choose("W222222", "TIP820")
            self.assertIsNone(tip)
            self.assertIn("内径・肉厚が入っていません", note)

    def test_第1パスに内径_肉厚が無ければ第2パスも探す(self):
        """統合 1.2.45・現場の指摘(資材計算と同じ): 第1パスの行が空なら第2パスで埋める。"""
        second = _results_file([("W222222", 400.0, 250.0)])                # 900mm
        with self._with(_FakeResults(self.path, second=second)):
            tip, note = CO.choose("W222222", "TIP1000")
        self.assertEqual(tip, "TIP820")
        self.assertIn("コイル外径 900mm", note)
        empty = _results_file([("W222222", 0, 0)])
        CO.forget()                       # 探したファイルを覚えているので忘れさせる
        with self._with(_FakeResults(self.path, second=empty)):
            tip, note = CO.choose("W222222", "TIP1000")
        self.assertIsNone(tip)
        self.assertIn("第2パスにもありません", note)

    def test_実績の形式でなければ何もしない(self):
        with self._with(_FakeResults(self.path, wanted=False)):
            self.assertEqual(CO.choose("W123456", "TIP1000"), (None, ""))


class ApplyWeightChoosesTipTest(WorkflowTestBase):
    def setUp(self):
        super().setUp()
        CO.forget()
        self.path = _results_file([("W123456", 400.0, 200.0)])     # 800mm → 820

    def tearDown(self):
        CO.forget()
        super().tearDown()

    def test_重量反映でチップボールを選び_案内を残す(self):
        st = self._cb_state(5, kensa_no="W123456", weight1="10", weight2="12")
        with mock.patch.object(CO, "_material",
                               return_value=(_FakeResults(self.path), ls4_source)):
            r = self.wf.apply_weight(st)
        self.assertTrue(r.ok, r.message)
        self.assertEqual(st.tip, "TIP820")
        self.assertIn("チップボール 820を選びました", r.message)
        saved = self.wf.load_state()
        self.assertEqual(saved.tip, "TIP820")
        self.assertIn("コイル外径 800mm", saved.tip_note)

    def test_読めなくても重量反映は止めない(self):
        st = self._cb_state(5, kensa_no="W123456", weight1="10", weight2="12")
        with mock.patch.object(CO, "choose", side_effect=RuntimeError("壊れた")):
            r = self.wf.apply_weight(st)
        self.assertTrue(r.ok, r.message)
        self.assertEqual(st.tip, "TIP1000")


if __name__ == "__main__":
    unittest.main()
