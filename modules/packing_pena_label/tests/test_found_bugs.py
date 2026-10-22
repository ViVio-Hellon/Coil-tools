# -*- coding: utf-8 -*-
"""構造点検(統合 1.2.26)で見つかった不具合の再現

直す前に、**直ったら通る試験**として置く(`expectedFailure`)。直したら印を外す。
印が付いたまま通ってしまったら(=気づかずに直った・別の所が変わった)、試験が落ちて知らせる。
"""
import unittest

from modules.packing_pena_label.app.models import FormState
from modules.packing_pena_label.app.routes.api import ApiRoutes
from modules.packing_pena_label.app.services import size_master as SM
from modules.packing_pena_label.app.services import validation as V

from .test_http import HttpTestCase


class AddedSizeWeightCheckTest(unittest.TestCase):
    """足したサイズ(台紙番号 15 以降)でも、重量が空なら重量反映を断る。"""

    def test_同梱のサイズは断る(self):
        self.assertEqual(V.check_weight_inputs(FormState(selected_ob=1)), "重量入力がありません　丈1")
        self.assertEqual(V.check_weight_inputs(FormState(selected_ob=2)), "重量入力がありません　丈2")

    def test_足したサイズの丈1も断る(self):
        for ob in (SM.FIRST_EXTRA_OB, SM.FIRST_EXTRA_OB + 2, SM.FIRST_EXTRA_OB + 4):
            self.assertEqual(V.check_weight_inputs(FormState(selected_ob=ob)),
                             "重量入力がありません　丈1", ob)

    def test_足したサイズの丈2も断る(self):
        for ob in (SM.FIRST_EXTRA_OB + 1, SM.FIRST_EXTRA_OB + 3):
            self.assertEqual(V.check_weight_inputs(FormState(selected_ob=ob)),
                             "重量入力がありません　丈2", ob)


    def test_重量が入っていれば通す(self):
        for ob, w in ((SM.FIRST_EXTRA_OB, "weight1"), (SM.FIRST_EXTRA_OB + 1, "weight2")):
            self.assertIsNone(V.check_weight_inputs(FormState(selected_ob=ob, **{w: "12.3"})), ob)

    def test_同梱のサイズの答えは前と同じ(self):
        """番号の決め打ち(前の判定)と、丈の決まり(いまの判定)が 0〜14 で全部同じ答え。"""
        def legacy(state, mode):
            w1, w2 = state.weight1 or "", state.weight2 or ""
            ob = state.selected_ob
            if mode == 1:
                return bool((ob and ob % 2 == 1 and ob <= 11) or ob == 13) and w1 == ""
            if mode == 2:
                return bool((ob and ob % 2 == 0 and ob <= 12) or ob == 14) and w2 == ""
            return V.inp_check(state, 3)
        for ob in range(0, 15):
            for cb in (0, 1, 3, 6):
                for named in (False, True):
                    for w1 in ("", "12.3"):
                        for w2 in ("", "9.8"):
                            st = FormState(selected_ob=ob, selected_cb=cb, named_cb=named,
                                           weight1=w1, weight2=w2)
                            for mode in (1, 2, 3):
                                self.assertEqual(V.inp_check(st, mode), legacy(st, mode),
                                                 (ob, cb, named, w1, w2, mode))

    def test_丈の決まりは1か所(self):
        for cfg in SM.SizeMaster().all():
            self.assertEqual(cfg.take_type, SM.take_of(cfg.ob_idx))


class ScreenGuardListTest(unittest.TestCase):
    """2枚目のタブを断る仕組みが、設定を書き換える口を全部守っている。"""

    def test_サイズの表示切替とサイズ追加も守る(self):
        for path in ("/api/settings/visible-sizes", "/api/settings/size-add"):
            self.assertIn(path, ApiRoutes._STATE_WRITING, path)


def _printable(html: str) -> bool:
    """印刷ビューが刷れる状態か(押せる「印刷する」があり、止められていない)。"""
    return "data-print-now>" in html and "停止中" not in html


class TarePrintStaleTest(HttpTestCase):
    """重量反映のあと計算し直していないとき、風袋計算の印刷ビューで
    **新しい検番に前の計算結果を並べて刷れない**(ラベルの印刷と同じ決まり)。"""

    def _cur(self, ken="Z987654", w2="22.0", n="11"):
        return {"selectedOb": 14, "selectedCb": 0, "namedCb": False,
                "tip": "TIP1000", "kensaNo": ken, "weight1": "", "weight2": w2,
                "coilH": {"1": "", "2": "", "3": n, "4": n}}

    def setUp(self):
        self.post("/api/clear", {"current": self._cur()})

    def test_計算したあとは刷れる(self):
        self.post("/api/apply-weight", {"current": self._cur()})
        self.post("/api/calc-tare", {"current": self._cur()})
        _, html = self.get("/tare/print")
        self.assertTrue(_printable(html))

    def test_入力を変えて計算し直していなければ_新しい検番で前の結果を刷れない(self):
        self.post("/api/apply-weight", {"current": self._cur()})
        self.post("/api/calc-tare", {"current": self._cur()})
        new = self._cur("Z111111", "33.0", "20")
        self.post("/api/apply-weight", {"current": new})
        _, html = self.get("/tare/print")
        self.assertFalse(_printable(html) and "検査番号: <b>Z111111" in html,
                         "新しい検番の見出しで、前の計算の風袋・GW を刷れてしまう")
        self.assertIn("入力が変わっています。印刷を止めています", html)
        self.assertNotIn("window.print();", html, "開いたとたんに刷る仕掛けが残っている")
        # 画面の風袋計算でも「このまま印刷」を止める
        _, screen = self.get("/tare")
        self.assertIn("このまま印刷（停止中）", screen)

    def test_クリアしてから別の検番を反映しても_前の結果は刷れない(self):
        """ラベル側の控えはクリアで消えるが、風袋計算の結果はクリアでも残る。"""
        self.post("/api/apply-weight", {"current": self._cur()})
        self.post("/api/calc-tare", {"current": self._cur()})
        self.post("/api/clear", {"current": self._cur()})
        new = self._cur("Z222222", "44.0", "9")
        self.post("/api/apply-weight", {"current": new})
        _, html = self.get("/tare/print")
        self.assertFalse(_printable(html), "クリアのあと、新しい検番で前の結果を刷れてしまう")

    def test_計算し直せば刷れる(self):
        self.post("/api/apply-weight", {"current": self._cur()})
        self.post("/api/calc-tare", {"current": self._cur()})
        new = self._cur("Z111111", "33.0", "20")
        self.post("/api/apply-weight", {"current": new})
        self.post("/api/calc-tare", {"current": new})
        _, html = self.get("/tare/print")
        self.assertTrue(_printable(html))
        self.assertIn("検査番号: <b>Z111111", html)


if __name__ == "__main__":
    unittest.main()


from .test_workflow import WorkflowTestBase  # noqa: E402


class ClearOnSelect535Test(WorkflowTestBase):
    """53.5mm のサイズを選ぶと、その丈のコイル高さ・NW・GW・TA を消す(足したサイズも同じ。統合 1.2.27)。"""

    def _filled(self):
        st = FormState()
        st.coil_h = {1: "1", 2: "2", 3: "3", 4: "4"}
        st.nw = {1: "10", 2: "20", 3: "30", 4: "40"}
        st.gw = {1: "11", 2: "21", 3: "31", 4: "41"}
        st.ta = {1: "5", 2: "6", 3: "7", 4: "8"}
        return st

    def test_同梱のサイズで消すのは前と同じ4つだけ(self):
        got = {c.ob_idx for c in self.master.all() if c.ob_idx < SM.FIRST_EXTRA_OB
               and c.clears_coil_on_select}
        self.assertEqual(got, {5, 6, 13, 14})

    def test_足した53_5mmのサイズも_選ぶとその丈を消す(self):
        ob = SM.FIRST_EXTRA_OB
        self.master.apply_extra([
            {"obIdx": ob, "baseName": "1.2mm×53.5mm", "widthMm": 53.5, "kataban": "T1", "symbol": ""},
            {"obIdx": ob + 2, "baseName": "1.2mm×45mm", "widthMm": 45.0, "kataban": "T2", "symbol": ""}])
        st = self._filled()
        self.assertTrue(self.wf.select_size(st, ob).ok)
        self.assertEqual((st.coil_h[1], st.nw[2], st.gw[1], st.ta[2]), ("", "", "", ""))
        self.assertEqual(st.coil_h[3], "3", "丈2の側は消さない")
        st = self._filled()
        self.assertTrue(self.wf.select_size(st, ob + 1).ok)
        self.assertEqual((st.coil_h[3], st.nw[4]), ("", ""))
        self.assertEqual(st.coil_h[1], "1", "丈1の側は消さない")
        st = self._filled()
        self.assertTrue(self.wf.select_size(st, ob + 2).ok)
        self.assertEqual(st.coil_h[1], "1", "53.5mm でないサイズは消さない")


class TipMustBeChosenTest(HttpTestCase):
    """チップボールを選んでいなければ計算しない(統合 1.2.32・現場の指示)。

    既定は 1000(画面を開いたときに選んである)。以前は選んでいなくても画面・サーバが 1000 や
    前の値で埋め、「チップボールチェックがありません」の停止(VBA IsValidSettings)に届かなかった。"""

    def _cur(self, tip):
        return {"selectedOb": 14, "selectedCb": 0, "namedCb": False, "tip": tip,
                "kensaNo": "Z987654", "weight1": "", "weight2": "22.0",
                "coilH": {"1": "", "2": "", "3": "11", "4": "11"}}

    def test_選んでいなければ止める(self):
        self.post("/api/apply-weight", {"current": self._cur("TIP1000")})
        _, j = self.post("/api/calc-tare", {"current": self._cur("")})
        self.assertFalse(j["ok"], j)
        self.assertIn("チップボールチェックがありません", j["message"])

    def test_既定は1000(self):
        from modules.packing_pena_label.app.models import FormState
        self.assertEqual(FormState().tip, "TIP1000")
        self.assertEqual(FormState.from_dict({"selectedOb": 1}).tip, "TIP1000")
        self.assertEqual(FormState.from_dict({"tip": ""}).tip, "")

    def test_選べば計算する(self):
        self.post("/api/apply-weight", {"current": self._cur("TIP950")})
        _, j = self.post("/api/calc-tare", {"current": self._cur("TIP950")})
        self.assertTrue(j["ok"], j)
