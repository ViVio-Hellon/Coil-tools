"""計算・印刷の記録(統合 1.2.31・現場の指示)

どの取得元(SQLite / Access / CSV / 古い内容)の、いつ時点の資材重量で計算し・刷ったかを残す。
印刷は、印刷の画面を閉じたときに画面から知らせてくる(/api/print-log)。
"""
from __future__ import annotations

import unittest

from .test_http import HttpTestCase


class PrintLogTest(HttpTestCase):
    def _cur(self):
        return {"selectedOb": 14, "selectedCb": 0, "namedCb": False, "tip": "TIP1000",
                "kensaNo": "Z987654", "weight1": "", "weight2": "22.0",
                "coilH": {"1": "", "2": "", "3": "11", "4": "11"}}

    def _calc(self):
        self.post("/api/clear", {"current": self._cur()})
        self.post("/api/apply-weight", {"current": self._cur()})
        _, j = self.post("/api/calc-tare", {"current": self._cur()})
        self.assertTrue(j["ok"], j)

    def _log(self):
        return self.ctx.wf.recent_log(20)

    def test_計算すると取得元と時点が残る(self):
        self._calc()
        row = self._log()[0]
        self.assertEqual((row["kind"], row["title"]), ("計算", "風袋計算"))
        m = row["data"]["material"]
        self.assertEqual(m["source"], "csv")
        self.assertRegex(m["asOf"], r"^\d{4}-\d{2}-\d{2} ")
        self.assertFalse(m["stale"])
        self.assertEqual(row["data"]["kensaNo"], "Z987654")

    def test_印刷の画面を閉じると_計算したときの資材重量と一緒に残る(self):
        self._calc()
        calc = self._log()[0]["data"]["material"]
        _, j = self.post("/api/print-log", {"path": "/labels/print", "search": "?ob=14"})
        self.assertTrue(j["ok"], j)
        row = self._log()[0]
        self.assertEqual((row["kind"], row["title"]), ("印刷画面", "小ラベル"))
        self.assertEqual(row["data"]["material"]["asOf"], calc["asOf"])
        self.assertEqual(row["data"]["sheets"][0]["kensaNo"], "Z987654")
        self.assertTrue(any(row["data"]["sheets"][0]["GW"]))
        _, j = self.post("/api/print-log", {"path": "/tare/print", "search": ""})
        self.assertEqual(self._log()[0]["title"], "風袋計算")

    def test_記録しない画面は残さない(self):
        before = len(self._log())
        _, j = self.post("/api/print-log", {"path": "/settings", "search": ""})
        self.assertFalse(j["ok"])
        self.assertEqual(len(self._log()), before)

    def test_記録の画面に出る(self):
        self._calc()
        self.post("/api/print-log", {"path": "/labels/print", "search": "?ob=14"})
        _, html = self.get("/print-log")
        self.assertIn("計算・印刷の記録", html)
        self.assertIn("小ラベル", html)
        self.assertIn("Z987654", html)


if __name__ == "__main__":
    unittest.main()
