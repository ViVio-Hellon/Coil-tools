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

    @unittest.expectedFailure
    def test_出力したあとで条番号を配り直さない(self):
        _post(self.client, "/api/zen-kotei", {"value": "10"})    # 積む前なので通る
        self._out(("1-12", "2-12"))
        _post(self.client, "/api/strands", {"confirm": True})
        state = self._state()
        self.assertEqual(state["strands"], [12, 12], "出した紙の条番号が盤面から消えた")
        res = _post(self.client, "/api/edit-output", {"no": 1})
        self.assertEqual(res.status_code, 200, res.get_json())


class EditGuardTest(_Base):
    """出力済み修正の途中にしてはいけない操作。"""

    @unittest.expectedFailure
    def test_修正中に廃棄してやめても_紙のコイルは廃棄にならない(self):
        self._out(("1-12", "2-12"))
        _post(self.client, "/api/edit-output", {"no": 1})
        _post(self.client, "/api/unstack", {"slot_no": 1})
        _post(self.client, "/api/discard", {"key": "1-12"})
        _post(self.client, "/api/edit-cancel", {})
        self.assertEqual(_coil_states(self._state())["1-12"], "used")
        self.assertNotIn("1-12", history_repo.load_snapshot(self.conn, "L5160Z0").discarded)

    @unittest.expectedFailure
    def test_修正をやめたら積み条数は元に戻る(self):
        self._out(("1-12", "1-11", "2-12"))
        _post(self.client, "/api/stack-max", {"value": "2"})
        _post(self.client, "/api/edit-output", {"no": 1})
        state = _post(self.client, "/api/edit-cancel", {}).get_json()["state"]
        self.assertEqual(state["stack_max"], 2)

    @unittest.expectedFailure
    def test_修正中に同じロットを開き直しても_黙って修正を捨てない(self):
        self._out(("1-12", "2-12"))
        _post(self.client, "/api/edit-output", {"no": 1})
        res = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        body = res.get_json()
        kept = body.get("state", {}).get("edit") is not None
        self.assertTrue(res.status_code == 409 or kept, (res.status_code, body))


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
