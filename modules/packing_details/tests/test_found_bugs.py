"""構造点検(統合 1.2.26)で見つかった不具合の再現

直す前に、**直ったら通る試験**として置く(`expectedFailure`)。直したら印を外す。
印が付いたまま通ってしまったら(=気づかずに直った・別の所が変わった)、試験が落ちて知らせる。
"""
from __future__ import annotations

import unittest
from unittest import mock

from modules.packing_details.meisai import db, history_repo, meisai_service, session

from .test_web import _client, _coil_states, _post, _seed


class _Base(unittest.TestCase):
    def setUp(self):
        self.client, self.conn = _client(self)
        _seed(self.conn)
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        _post(self.client, "/api/strands", {})

    def _out(self, keys):
        _post(self.client, "/api/stack-max", {"value": str(len(keys))})
        for key in keys:
            _post(self.client, "/api/stack", {"key": key})
        body = _post(self.client, "/api/output", {"confirm": True}).get_json()
        self.assertTrue(body.get("ok"), body)
        return body

    def _state(self):
        return _post(self.client, "/api/reverse", {}).get_json()["state"]


class StaleSnapshotRestoreTest(_Base):
    """途中経過が古くても(出力のあとの保存に失敗した)、開き直せば紙のコイルは使用済み。"""

    def test_出したあとの保存に失敗しても_開き直せば紙のコイルは使用済み(self):
        _post(self.client, "/api/stack-max", {"value": "2"})
        for key in ("1-12", "2-12"):
            _post(self.client, "/api/stack", {"key": key})
        with mock.patch.object(session, "save", side_effect=db.WriteError("試験の失敗")):
            body = _post(self.client, "/api/output", {"confirm": True}).get_json()
        self.assertTrue(body["ok"])
        self.assertIn("途中経過の保存に失敗しました", body["message"])
        session.clear()                                   # 落ちた想定
        state = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"}).get_json()["state"]
        states = _coil_states(state)
        self.assertEqual((states["1-12"], states["2-12"]), ("used", "used"))


class ZenKoteiBypassTest(_Base):
    """前工程実績数のロック(0.13.25)の抜け道: 積む前に変えておき、出力したあとで条番号表示を押す。"""

    def test_積む前に前工程実績数を変えたら_条番号表示し直すまで積めない(self):
        """丈数と同じ決まり(統合 1.2.26)。古い条番号のまま積むと、無いコイルを紙に載せてしまう。"""
        state = _post(self.client, "/api/zen-kotei", {"value": "10"}).get_json()["state"]
        self.assertEqual(state["strands"], [])
        state = _post(self.client, "/api/strands", {}).get_json()["state"]
        self.assertEqual(state["strands"], [5, 5])

    def test_配ったあとで数が変わった盤面は_積み上げを始めたら配り直さない(self):
        """古い版で作られた食い違い(配ったあとで前工程実績数だけ変わった)を再現して確かめる。"""
        self._out(("1-12", "2-12"))
        session.current().zen_kotei = 10
        res = _post(self.client, "/api/strands", {"confirm": True})
        self.assertEqual(res.status_code, 409)
        self.assertEqual(res.get_json()["error"]["code"], "strands_locked")
        state = self._state()
        self.assertEqual(state["strands"], [12, 12], "出した紙の条番号が盤面から消えた")
        self.assertIn("strands", state["locks"])
        res = _post(self.client, "/api/edit-output", {"no": 1})
        self.assertEqual(res.status_code, 200, res.get_json())

    def test_配分が変わらない再表示は_積み上げを始めても通す(self):
        self._out(("1-12", "2-12"))
        res = _post(self.client, "/api/strands", {"confirm": True})
        self.assertEqual(res.status_code, 200)
        self.assertEqual(_coil_states(res.get_json()["state"])["1-12"], "used")


class EditGuardTest(_Base):
    """出力済み修正の途中にしてはいけない操作。"""

    def test_修正中に廃棄してやめても_紙のコイルは廃棄にならない(self):
        self._out(("1-12", "2-12"))
        _post(self.client, "/api/edit-output", {"no": 1})
        _post(self.client, "/api/unstack", {"slot_no": 1})
        _post(self.client, "/api/discard", {"key": "1-12"})
        _post(self.client, "/api/edit-cancel", {})
        self.assertEqual(_coil_states(self._state())["1-12"], "used")
        self.assertNotIn("1-12", history_repo.load_snapshot(self.conn, "L5160Z0").discarded)

    def test_修正をやめたら積み条数は元に戻る(self):
        self._out(("1-12", "1-11", "2-12"))
        _post(self.client, "/api/stack-max", {"value": "2"})
        _post(self.client, "/api/edit-output", {"no": 1})
        state = _post(self.client, "/api/edit-cancel", {}).get_json()["state"]
        self.assertEqual(state["stack_max"], 2)

    def test_修正中に同じロットを開き直しても_黙って修正を捨てない(self):
        self._out(("1-12", "2-12"))
        _post(self.client, "/api/edit-output", {"no": 1})
        res = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        body = res.get_json()
        kept = body.get("state", {}).get("edit") is not None
        self.assertTrue(res.status_code == 409 or kept, (res.status_code, body))


class GuardMatrixTest(_Base):
    """操作 × 状態(ふつう・出力したあと・修正中)で、断る/通すが1か所の決まりどおりか。

    断るときは、画面に渡す `state.locks` と同じ文言で断る(画面とサーバで食い違わない)。"""

    ACTIONS = {
        "zen_kotei": ("/api/zen-kotei", {"value": "24"}),
        "jou_su": ("/api/jou-su", {"value": "2"}),
        "strands": ("/api/strands", {"confirm": True}),
        "lot": ("/api/lot", {"lot_no": "L5160Z0"}),
        "clear_lot": ("/api/clear-lot", {}),
        "edit_output": ("/api/edit-output", {"no": 1}),
        "stack_max": ("/api/stack-max", {"value": "2"}),
        "discard": ("/api/discard", {"key": "1-1"}),
        "reverse": ("/api/reverse", {}),
    }
    EXPECTED = {
        "normal": set(),
        "started": {"zen_kotei", "jou_su"},
        "editing": {"zen_kotei", "jou_su", "strands", "lot", "clear_lot", "edit_output"},
    }

    def _prepare(self, kind):
        if kind in ("started", "editing"):
            self._out(("1-12", "2-12"))
        if kind == "editing":
            _post(self.client, "/api/edit-output", {"no": 1})

    def test_断る操作は決まりどおりで_画面のlocksと同じ文言(self):
        for kind, refused in self.EXPECTED.items():
            for action, (path, body) in self.ACTIONS.items():
                with self.subTest(kind=kind, action=action):
                    self.setUp()
                    self._prepare(kind)
                    locks = self._state()["locks"]
                    self.assertEqual(set(locks), refused, kind)
                    res = _post(self.client, path, body)
                    if action in refused:
                        self.assertEqual(res.status_code, 409, res.get_json())
                        err = res.get_json()["error"]
                        self.assertEqual(err["code"], locks[action]["code"])
                        self.assertEqual(err["message"], locks[action]["message"])
                    else:
                        self.assertNotEqual(res.status_code, 409, (path, res.get_json()))

    def test_修正中の重量の打ち直しは_戻したことを知らせる(self):
        self._out(("1-12", "2-12"))
        _post(self.client, "/api/edit-output", {"no": 1})
        _post(self.client, "/api/weights", {"weights": ["300", "301"]})
        body = _post(self.client, "/api/output", {"confirm": True}).get_json()
        self.assertIn("重量は修正に入る前の値に戻しました", body["message"])
        self.assertEqual(meisai_service.find_output(self.conn, "L5160Z0", 1).weights, [300, 301])
        self.assertEqual(body["state"]["weights"], [250, 248])

    def test_修正中のロット切替は_了解があれば修正をやめて開く(self):
        self._out(("1-12", "2-12"))
        _post(self.client, "/api/edit-output", {"no": 1})
        res = _post(self.client, "/api/lot", {"lot_no": "L5160Z0", "cancel_edit": True})
        self.assertEqual(res.status_code, 200)
        state = res.get_json()["state"]
        self.assertIsNone(state["edit"])
        self.assertEqual(_coil_states(state)["1-12"], "used")

    def test_修正中は_紙のコイルを廃棄にできない(self):
        self._out(("1-12", "2-12"))
        _post(self.client, "/api/edit-output", {"no": 1})
        _post(self.client, "/api/unstack", {"slot_no": 1})
        body = _post(self.client, "/api/discard", {"key": "1-12"}).get_json()
        self.assertEqual(body.get("ignored"), "used")


class OutputThenSaveFailsTest(_Base):
    """出力は書けたのに、そのあとの途中経過の保存で失敗したとき。"""

    def test_書けていない_と言わず_同じコイルで次のNoを出させない(self):
        _post(self.client, "/api/stack-max", {"value": "2"})
        for key in ("1-12", "2-12"):
            _post(self.client, "/api/stack", {"key": key})
        with mock.patch.object(session, "save", side_effect=db.WriteError("試験の失敗")):
            res = _post(self.client, "/api/output", {"confirm": True})
        written = meisai_service.list_outputs(self.conn, "L5160Z0")
        # 書けたなら「書けた」、書けなかったなら1行も残さない。どちらかに揃える
        if written:
            self.assertEqual(res.status_code, 200, "No1 は書けているのに失敗と返した")
        else:
            self.assertNotEqual(res.status_code, 200)
        # どちらでも、同じコイルを次の No として出せてはいけない(積み上げが残ったままにしない)
        body = _post(self.client, "/api/output", {"confirm": True}).get_json()
        self.assertEqual(len(meisai_service.list_outputs(self.conn, "L5160Z0")), 1 if written else 0,
                         body)


if __name__ == "__main__":
    unittest.main()


class OtherLotRequestTest(_Base):
    """前のロットの画面で打った操作が、開いたロットに入らない(統合 1.2.26・通し試験で見つけた)。

    画面は押した順に1つずつ送る。ロットを開く要求のうしろに、開く前の画面で打った重量が並ぶと、
    前のロットの丈2の重量が、開いたロットに入っていた。画面は押したときのロットを
    `X-Screen-Lot` に付けて送り、サーバはいまのロットと違えば使わない。
    """

    def setUp(self):
        super().setUp()
        _seed(self.conn, lot="A123456", tate="2", yoko="6")

    def _post_from(self, screen_lot, path, data):
        from .test_web import HEADERS
        return self.client.post(path, json=data, headers={**HEADERS, "X-Screen-Lot": screen_lot})

    def test_開く前の画面で打った重量は_開いたロットに入れない(self):
        _post(self.client, "/api/lot", {"lot_no": "A123456", "confirm": True})
        res = self._post_from("L5160Z0", "/api/weights", {"weights": ["300", "248"]})
        self.assertEqual(res.status_code, 409)
        self.assertEqual(res.get_json()["error"]["code"], "lot_changed")
        self.assertEqual(session.current().lot_no, "A123456")
        self.assertNotIn(248, session.current().weights, "前のロットの重量が入った")

    def test_いまのロットの画面からなら通す(self):
        _post(self.client, "/api/lot", {"lot_no": "A123456", "confirm": True})
        res = self._post_from("A123456", "/api/weights", {"weights": ["300", "302"]})
        self.assertEqual(res.status_code, 200, res.get_json())
        self.assertEqual(list(session.current().weights), [300, 302])

    def test_ロットを開く口は_開く前の画面から押せる(self):
        res = self._post_from("L5160Z0", "/api/lot", {"lot_no": "A123456", "confirm": True})
        self.assertEqual(res.status_code, 200, res.get_json())
        self.assertEqual(session.current().lot_no, "A123456")

    def test_前のロットの画面で押したコイルを_開いたロットに積まない(self):
        _post(self.client, "/api/lot", {"lot_no": "A123456", "confirm": True})
        res = self._post_from("L5160Z0", "/api/stack", {"key": "1-6"})
        self.assertEqual(res.status_code, 409)
        self.assertEqual(res.get_json()["error"]["code"], "lot_changed")
