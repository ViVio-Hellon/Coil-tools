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

    @unittest.expectedFailure
    def test_足したサイズの丈1も断る(self):
        for ob in (SM.FIRST_EXTRA_OB, SM.FIRST_EXTRA_OB + 2, SM.FIRST_EXTRA_OB + 4):
            self.assertEqual(V.check_weight_inputs(FormState(selected_ob=ob)),
                             "重量入力がありません　丈1", ob)

    @unittest.expectedFailure
    def test_足したサイズの丈2も断る(self):
        for ob in (SM.FIRST_EXTRA_OB + 1, SM.FIRST_EXTRA_OB + 3):
            self.assertEqual(V.check_weight_inputs(FormState(selected_ob=ob)),
                             "重量入力がありません　丈2", ob)


class ScreenGuardListTest(unittest.TestCase):
    """2枚目のタブを断る仕組みが、設定を書き換える口を全部守っている。"""

    @unittest.expectedFailure
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

    @unittest.expectedFailure
    def test_入力を変えて計算し直していなければ_新しい検番で前の結果を刷れない(self):
        self.post("/api/apply-weight", {"current": self._cur()})
        self.post("/api/calc-tare", {"current": self._cur()})
        new = self._cur("Z111111", "33.0", "20")
        self.post("/api/apply-weight", {"current": new})
        _, html = self.get("/tare/print")
        self.assertFalse(_printable(html) and "検査番号: <b>Z111111" in html,
                         "新しい検番の見出しで、前の計算の風袋・GW を刷れてしまう")


if __name__ == "__main__":
    unittest.main()
