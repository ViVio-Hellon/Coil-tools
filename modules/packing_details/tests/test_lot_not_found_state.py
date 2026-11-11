"""ロットが見つからなかったとき、前のロットの表示・知らせを残さない(統合 1.2.82・現場の指摘)

現場の指摘: 「該当するﾛｯﾄ情報が見つかりません。」「このロットは梱包課共有の仕掛(LS4LOT)に無いため、前工程実績数が0です。」が
ずっと残っていて、いまのロットの情報か分からない。ロット検索のときにリセットしてほしい。
見つからなかった返事に、いまの状態(前のロットは閉じている)を添え、画面はそれで描き直す。
"""
from __future__ import annotations

import unittest

from .test_web import _client, _post, _seed


class LotNotFoundStateTest(unittest.TestCase):
    def setUp(self):
        self.client, self.conn = _client(self)
        _seed(self.conn, "L5160Z0", tate="0", yoko="0")      # 割数 0 → 「前工程実績数が0です」が出るロット

    def test_見つからなければ_いまの状態を添えて前のロットは閉じている(self):
        opened = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        self.assertEqual(opened.status_code, 200, opened.get_json())
        self.assertIn("前工程実績数が0です", opened.get_json()["state"]["lot"]["jou_su_note"])
        res = _post(self.client, "/api/lot", {"lot_no": "ZZZ9999", "confirm": True})
        self.assertEqual(res.status_code, 404)
        body = res.get_json()
        self.assertEqual(body["error"]["code"], "not_found")
        self.assertIn("state", body, "画面を描き直せるよう状態を添える")
        self.assertFalse(body["state"]["found"])
        self.assertFalse(body["state"].get("lot"), "前のロットの情報(前工程実績数が0です の案内つき)を返さない")


if __name__ == "__main__":
    unittest.main()
