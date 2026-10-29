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

    def test_どちらの部品も無ければ何もしない(self):
        with self._with(_FakeResults(self.path, wanted=False)), \
                mock.patch.object(CO, "_details", return_value=None):
            self.assertEqual(CO.choose("W123456", "TIP1000"), (None, ""))

    def test_実績の形式でなくても黙らない(self):
        """統合 1.2.57・現場の指摘「外径計算してチップボールを選んでいない」: 以前は資材計算の仕掛パスの
        名前が LS4LOT のままなら**何も言わずに**選ばなかった。いまは理由を出す。"""
        with self._with(_FakeResults(self.path, wanted=False)), \
                mock.patch.object(CO, "lookup_lot",
                                  return_value={"found": False, "reason": "仕掛のファイルにロット W123456 がありません",
                                                "notes": [], "kensa": "", "outer": None}):
            tip, note = CO.choose("W123456", "TIP1000")
        self.assertIsNone(tip)
        self.assertIn("自動で選んでいません", note)
        self.assertIn("ロット W123456 がありません", note)


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


def _ls4lot_file(rows, *, kensa_col="検査番号", name="LS4LOT.sqlite3"):
    """梱包明細が読む仕掛(名前は LS4LOT のまま)に、検査番号・内径・肉厚の列もあるファイル。"""
    folder = tempfile.mkdtemp(prefix="ppl_ls4lot_")
    path = os.path.join(folder, name)
    with sqlite3.connect(path) as c:
        c.execute(f'CREATE TABLE "仕掛" ("{ls4_source.COL_LOT}" TEXT, "{kensa_col}" TEXT,'
                  f' "当工程設計_縦割数" INTEGER, "当工程設計_横割数" INTEGER,'
                  f' "{ls4_source.COL_INNER}" REAL, "{ls4_source.COL_THICK}" REAL)')
        c.executemany('INSERT INTO "仕掛" VALUES (?,?,?,?,?,?)',
                      [(lot, ken, 2, 12, inner, thick) for lot, ken, inner, thick in rows])
    return path


class LotNoTest(WorkflowTestBase):
    """LOTNO から新検番とチップボール(統合 1.2.57・現場の指示)。

    LOTNO を探すデータは梱包明細が読む仕掛(LS4LOT)。ﾛｯﾄ番号がヒットしたら同じ行の検査番号を新検番に入れ、
    そのとき外径も計算してチップボールを選ぶ。ファイルの名前が LS4LOT のままでも、列があれば外径を出す。
    """

    def setUp(self):
        super().setUp()
        CO.forget()
        from modules.packing_details.meisai import lot_lookup
        self.lot_lookup = lot_lookup
        lot_lookup.forget()
        self.path = _ls4lot_file([("L6241W0", "N6X0945", 400.0, 200.0),      # 800mm → 820
                                  ("L7160C0", "H670129", 508.0, 260.0)])     # 1028mm → 950
        from pathlib import Path
        self.patch = mock.patch.object(lot_lookup, "sources", return_value=[Path(self.path)])
        self.patch.start()
        self.addCleanup(self.patch.stop)

    def test_LOTNOで同じ行の検査番号とチップボール(self):
        st = self._cb_state(5)
        r = self.wf.lookup_lot(st, "l6241w0")                     # 小文字でも
        self.assertTrue(r.ok, r.message)
        self.assertEqual(st.kensa_no, "N6X0945")
        self.assertEqual(st.lot_no, "L6241W0")
        self.assertEqual(st.tip, "TIP820")
        self.assertIn("検査番号 N6X0945 を新検番に入れました", r.message)
        self.assertIn("コイル外径 800mm", r.message)
        self.assertEqual(r.data["state"]["kensaNo"], "N6X0945")
        saved = self.wf.load_state()
        self.assertEqual((saved.kensa_no, saved.tip, saved.lot_no), ("N6X0945", "TIP820", "L6241W0"))

    def test_重量反映はLOTNOで外径を引く(self):
        """検査番号はロット番号と違う(L7160C0 の検査番号は H670129)。LOTNO で引いた検査番号なら LOTNO で引く。"""
        st = self._cb_state(5)
        self.wf.lookup_lot(st, "L7160C0")
        st.tip = "TIP1000"                                         # 手で選び直していた
        st.weight1, st.weight2 = "10", "12"
        r = self.wf.apply_weight(st)
        self.assertTrue(r.ok, r.message)
        self.assertEqual(st.tip, "TIP950")
        self.assertIn("コイル外径 1028mm", r.message)

    def test_見つからなければ断って何も変えない(self):
        st = self._cb_state(5, kensa_no="N1A2345")
        r = self.wf.lookup_lot(st, "L0000000")
        self.assertFalse(r.ok)
        self.assertIn("LOTNO L0000000 が見つかりません", r.message)
        self.assertEqual(st.kensa_no, "N1A2345")

    def test_検査番号の列が無ければそう知らせる(self):
        from pathlib import Path
        other = _ls4lot_file([("L6241W0", "X", 400.0, 200.0)], kensa_col="備考")
        self.lot_lookup.forget()
        with mock.patch.object(self.lot_lookup, "sources", return_value=[Path(other)]):
            st = self._cb_state(5, kensa_no="N1A2345")
            r = self.wf.lookup_lot(st, "L6241W0")
        self.assertTrue(r.ok)
        self.assertEqual(st.kensa_no, "N1A2345", "検査番号が無いので新検番は変えない")
        self.assertIn("検査番号の列が見つかりません", r.message)
        self.assertEqual(st.tip, "TIP820", "外径からは選ぶ")

    def test_画面のAPI(self):
        from modules.packing_pena_label.app.routes.api import ApiRoutes
        import time
        api = ApiRoutes(self.wf, getattr(self, "cfg", None), time.time())
        out = api._lot_lookup({"lotNo": "L6241W0", "current": {"tip": "TIP1000", "kensaNo": ""}})
        self.assertTrue(out["ok"], out)
        self.assertEqual(out["state"]["kensaNo"], "N6X0945")
        self.assertEqual(out["state"]["tip"], "TIP820")
        self.assertIn("/api/lot-lookup", ApiRoutes._STATE_WRITING)


if __name__ == "__main__":
    unittest.main()
