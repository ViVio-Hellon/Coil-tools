# -*- coding: utf-8 -*-
"""VBA 元ソースとの突き合わせで「Python を直す」とした 6 件(統合 1.2.35・現場の指示)

1. 資材マスタに無い資材・単重/係数が数でない資材 → 計算を止める(以前は 0kg で続けた。VBA は止まった)
2. 隠れている丈の梱包は計算しない・風袋の紙に出さない(以前は NW=0 のブロックが紙に出た)
3. 設定の確かめで止めたとき、何も消さない(以前は台紙の本数欄だけ消え、保存した NW・GW と食い違った)
4. 本数は整数の文字にそろえる(「12.0」が台紙・見出しに出た)
5. 本数の欄は数字だけ(画面)
6. 羅列計算も、ラベルの重量が数でなければ止める
"""
import csv
import os
import tempfile
import unittest

from modules.packing_pena_label.app.models import FormState
from modules.packing_pena_label.app.services import size_master as SM
from modules.packing_pena_label.app.services import validation as V

from . import test_http
from ._material import SAMPLE
from .test_http import HttpTestCase

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def _cur(**kw):
    base = {"selectedOb": 0, "selectedCb": 1, "namedCb": False, "tip": "TIP1000",
            "kensaNo": "Z987654", "weight1": "20.0", "weight2": "22.0",
            "coilH": {"1": "11", "2": "", "3": "11", "4": ""}}
    base.update(kw)
    return base


class _CustomMaterial(HttpTestCase):
    """資材マスタ CSV を書き換えた写しで動かす。"""

    @staticmethod
    def edit(rows):
        return rows

    @classmethod
    def setUpClass(cls):
        with open(SAMPLE, encoding="utf-8-sig", newline="") as f:
            rows = list(csv.reader(f))
        folder = tempfile.mkdtemp(prefix="ppl_bad_material_")
        path = os.path.join(folder, "資材重量.csv")
        with open(path, "w", encoding="utf-8", newline="") as f:
            csv.writer(f).writerows([rows[0]] + cls.edit([r for r in rows[1:] if r]))
        orig = test_http.business_csv
        test_http.business_csv = lambda: path
        try:
            super().setUpClass()
        finally:
            test_http.business_csv = orig


class MissingMaterialStopsTest(_CustomMaterial):
    @staticmethod
    def edit(rows):
        return [r for r in rows if r[1] != "PETバンド"]

    def test_無い資材があれば止めて何も書かない(self):
        _, j = self.post("/api/apply-weight", {"current": _cur()})
        self.assertTrue(j["ok"], j)
        _, j = self.post("/api/calc-tare", {"current": _cur()})
        self.assertFalse(j["ok"], j)
        self.assertEqual(j["level"], "error")
        self.assertIn("資材重量なし", j["message"])
        self.assertIn("PETバンド", j["message"])
        self.assertEqual(j["state"]["nw"]["1"], "")
        self.assertEqual(self.ctx.store.all_tare(), {})

    def test_羅列も止める(self):
        self.post("/api/apply-weight", {"current": _cur()})
        _, j = self.post("/api/calc-list", {"current": _cur()})
        self.assertFalse(j["ok"], j)
        self.assertIn("PETバンド", j["message"])


class NonNumericMaterialStopsTest(_CustomMaterial):
    @staticmethod
    def edit(rows):
        return [[r[0], r[1], "abc", r[3]] if r[1] == "テスラピン" else r for r in rows]

    def test_単重が数でなければ止める(self):
        self.post("/api/apply-weight", {"current": _cur()})
        _, j = self.post("/api/calc-tare", {"current": _cur()})
        self.assertFalse(j["ok"], j)
        self.assertIn("テスラピン", j["message"])
        self.assertIn("数でない", j["message"])


class HiddenTakeTest(HttpTestCase):
    def test_隠れた丈の梱包は計算しない_紙に出さない(self):
        # 丈1 だけのサイズ(奇数の OB)。丈2 の本数欄に値が残っている
        cur = _cur(selectedOb=13, selectedCb=0, coilH={"1": "11", "2": "", "3": "11", "4": "10"})
        _, j = self.post("/api/apply-weight", {"current": cur})
        self.assertTrue(j["ok"], j)
        _, j = self.post("/api/calc-tare", {"current": cur})
        self.assertTrue(j["ok"], j)
        self.assertEqual(sorted(self.ctx.store.all_tare()), [1])
        self.assertEqual(j["state"]["nw"]["3"], "")
        _, page = self.get("/tare/print")
        self.assertIn("丈1_1梱包目", page)
        self.assertNotIn("丈2_1梱包目", page)

    def test_表示中の丈に本数が無ければ止める(self):
        cur = _cur(selectedOb=13, selectedCb=0, coilH={"1": "", "2": "", "3": "11", "4": ""})
        self.post("/api/apply-weight", {"current": cur})
        _, j = self.post("/api/calc-tare", {"current": cur})
        self.assertFalse(j["ok"], j)
        self.assertIn("表示中の丈に本数が入力されていません", j["message"])


class StopKeepsEverythingTest(HttpTestCase):
    def test_設定で止めても前の計算結果と台紙は消えない(self):
        cur = _cur()
        self.post("/api/apply-weight", {"current": cur})
        _, ok = self.post("/api/calc-tare", {"current": cur})
        self.assertTrue(ok["ok"], ok)
        nw = ok["state"]["nw"]["1"]
        self.assertTrue(nw)
        _, j = self.post("/api/calc-tare", {"current": _cur(tip="")})
        self.assertFalse(j["ok"], j)
        self.assertIn("チップボールチェックがありません", j["message"])
        self.assertEqual(j["state"]["nw"]["1"], nw)
        st = self.ctx.wf.load_state()
        self.assertEqual(st.nw[1], nw)
        ob1, _ = self.ctx.master.cb_to_pair(1)
        cfg = self.ctx.master.get(ob1)
        self.assertEqual(self.ctx.store.get_cell(ob1, cfg.data_row, SM.COL_COILH), "11")
        self.assertTrue(self.ctx.store.all_tare())


class CoilCountTextTest(unittest.TestCase):
    def test_本数は整数の文字(self):
        self.assertEqual(V.filter_coil_text("12.0"), "12")
        self.assertEqual(V.filter_coil_text("１２"), "12")
        self.assertEqual(V.filter_coil_text("012"), "12")
        self.assertEqual(V.filter_coil_text("12.5"), "")

    def test_画面の本数欄は数字だけ(self):
        with open(os.path.join(ROOT, "app", "static", "js", "app.js"), encoding="utf-8") as f:
            js = f.read()
        part = js[js.index('kind === "coil"'):]
        part = part[:part.index("}")]
        self.assertIn("/^\\d+$/", part)
        self.assertIn('replace(/[^0-9]/g, "")', part)


class ListWeightTest(unittest.TestCase):
    def test_羅列はラベルの重量が数でなければ止める(self):
        from modules.packing_pena_label.app.services.workflow import Workflow
        st = FormState()
        st.selected_cb = 1
        st.lbl_weight1 = "-kg"
        st.lbl_weight2 = "22.0kg"
        msg = Workflow._list_weight_problem(st)
        self.assertIn("丈1の重量が数になっていません", msg)
        st.lbl_weight1 = "20.0kg"
        self.assertIsNone(Workflow._list_weight_problem(st))


if __name__ == "__main__":
    unittest.main()


class NextLotClearsTest(HttpTestCase):
    """刷ったあとは残し、次のロット(検番が変わった重量反映)で前の NW・GW・高さと台紙の本数欄を消す。"""

    def test_同じ検番なら残し_次の検番で消す(self):
        cur = _cur()
        self.post("/api/apply-weight", {"current": cur})
        _, ok = self.post("/api/calc-tare", {"current": cur})
        nw = ok["state"]["nw"]["1"]
        self.assertTrue(nw)
        ob1, _ = self.ctx.master.cb_to_pair(1)
        cfg = self.ctx.master.get(ob1)
        # 刷っても(印刷の対象を取っても)消えない
        self.post("/api/print-targets", {"current": cur})
        self.assertEqual(self.ctx.wf.load_state().nw[1], nw)
        # 同じ検番で重量反映し直しても消さない
        _, j = self.post("/api/apply-weight", {"current": cur})
        self.assertEqual(j["state"]["nw"]["1"], nw)
        self.assertNotIn("前のロット", j["message"])
        # 次のロット
        _, j = self.post("/api/apply-weight", {"current": _cur(kensaNo="Z111111")})
        self.assertTrue(j["ok"], j)
        self.assertIn("前のロットの NW・GW・高さと台紙の本数欄を消しました", j["message"])
        self.assertEqual(j["state"]["nw"]["1"], "")
        self.assertEqual(self.ctx.store.get_cell(ob1, cfg.data_row, SM.COL_COILH), "")
        self.assertEqual(j["state"]["coilH"]["1"], "11", "本数の入力欄は消さない(VBA のクリアと同じ)")


class AllSizeDigitTest(HttpTestCase):
    def test_2ケタ3ケタは既定で未選択(self):
        combo = self.ctx.wf.all_size.master.rows[0]["name"]
        self.post("/api/all-size/prepare", {"comboText": combo})
        _, page = self.get("/all-size")
        for value in ("2", "3"):
            tag = page[page.index(f'name="asDigit" value="{value}"'):]
            self.assertNotIn("checked", tag[:tag.index(">")], value)
        _, j = self.post("/api/all-size/submit", {"comboText": combo, "kensaNo": "W123456",
                                                  "threeDigit": None, "w1": "10", "w2": "12"})
        self.assertFalse(j["ok"])
        self.assertIn("2ケタ、3ケタどちらにもチェックが入っていません", j["message"])
        # 決定したあとは、選んだほうに印が付く
        _, j = self.post("/api/all-size/submit", {"comboText": combo, "kensaNo": "W123456",
                                                  "threeDigit": False, "w1": "10.5", "w2": "12"})
        self.assertTrue(j["ok"], j)
        _, page = self.get("/all-size")
        tag = page[page.index('name="asDigit" value="2"'):]
        self.assertIn("checked", tag[:tag.index(">")])

    def test_重量の入力は指定サイズと同じ(self):
        _, page = self.get("/all-size")
        with open(os.path.join(ROOT, "app", "templates", "index.html"), encoding="utf-8") as f:
            index = f.read()
        self.assertIn('inputmode="decimal" data-filter="num"', index)
        for w in ("asW1", "asW2", "asW4", "asW5"):
            tag = page[page.index(f'id="{w}"'):]
            tag = tag[:tag.index(">")]
            self.assertIn('inputmode="decimal"', tag)
            self.assertIn('data-filter="num"', tag)
