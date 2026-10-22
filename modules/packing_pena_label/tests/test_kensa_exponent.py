# -*- coding: utf-8 -*-
"""検査番号の穴(統合 1.2.36・現場の指摘)

- 5・6 文字目などに「E」を入れた W111E11 が通っていた。VBA の IsNumeric が「1E11」を指数表記の数と
  みなすため(VBA も Python も同じ)。そのままラベルとバーコードに載った
- 8 文字以上が通っていた(判定は 7 文字目までしか見ない)。画面は maxlength で、8 桁を貼り付けると
  黙って 7 桁に切られていた → 超える入力は入れずに知らせる
"""
import os
import unittest

from modules.packing_pena_label.app.services import validation as V

from .test_http import HttpTestCase

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def check(raw):
    return V.validate_kensa_no(V.normalize_kensa_no(raw))


class KensaDigitsTest(unittest.TestCase):
    def test_指数表記や記号は数字とみなさない(self):
        for raw in ("W111E11", "W1111E1", "W11E111", "W1E1111", "W111e11",
                    "W1+1111", "W1.1111", "W1-1111", "W1 1111", "W11D111"):
            with self.subTest(raw=raw):
                self.assertIsNotNone(check(raw))

    def test_EのときもVBAと同じ文言(self):
        self.assertIn("1～4桁目にアルファベットが入力されています", check("W111E11"))
        self.assertIn("5桁目にアルファベットが入力されています", check("W1E1111"))

    def test_8文字以上は弾く(self):
        self.assertEqual(check("W1111111"), "検査NOは7桁です(8桁入力されています)")
        self.assertIsNotNone(check("W11111111"))

    def test_1文字目は英字(self):
        self.assertIsNotNone(check("1111111"))
        self.assertIsNotNone(check("+111111"))

    def test_正しいものは通る(self):
        for raw in ("W111111", "w111111", "Z987654", "Ｗ１１１１１１"):
            with self.subTest(raw=raw):
                self.assertIsNone(check(raw))

    def test_これまで弾いていた入力の文言は変わらない(self):
        self.assertIn("6桁目にアルファベット", check("WA11111"))
        self.assertIn("1桁目にアルファベット", check("W12345"))
        self.assertIn("7桁目に数字", check("1111111"))
        self.assertIn("5桁目にアルファベット", check("W1A1111"))


class KensaOverHttpTest(HttpTestCase):
    def test_重量反映と全サイズの決定で止める(self):
        cur = {"selectedOb": 0, "selectedCb": 1, "namedCb": False, "tip": "TIP1000",
               "kensaNo": "W111E11", "weight1": "20.0", "weight2": "22.0",
               "coilH": {"1": "", "2": "", "3": "", "4": ""}}
        _, j = self.post("/api/apply-weight", {"current": cur})
        self.assertFalse(j["ok"], j)
        combo = self.ctx.wf.all_size.master.rows[0]["name"]
        _, j = self.post("/api/all-size/submit", {"comboText": combo, "kensaNo": "W111E11",
                                                  "threeDigit": False, "w1": "10", "w2": "12"})
        self.assertFalse(j["ok"], j)
        _, j = self.post("/api/all-size/submit", {"comboText": combo, "kensaNo": "W1111112",
                                                  "threeDigit": False, "w1": "10", "w2": "12"})
        self.assertFalse(j["ok"], j)
        self.assertIn("7桁", j["message"])


class KensaScreenTest(HttpTestCase):
    def test_欄は7桁を超えたら入れない(self):
        """maxlength(黙って切る)ではなく data-max(超えたら前の値に戻して知らせる)。"""
        _, index = self.get("/")
        _, allsize = self.get("/all-size")
        for page, el in ((index, 'id="kensaNo"'), (allsize, 'id="asKen"')):
            tag = page[page.index(el):]
            tag = tag[:tag.index(">")]
            self.assertIn('data-max="7"', tag)
            self.assertNotIn("maxlength", tag)
        with open(os.path.join(ROOT, "app", "static", "js", "app.js"), encoding="utf-8") as f:
            js = f.read()
        self.assertIn('getAttribute("data-max")', js)
        self.assertIn("桁の入力は入れていません", js)
        self.assertIn('addEventListener("beforeinput"', js)


if __name__ == "__main__":
    unittest.main()
