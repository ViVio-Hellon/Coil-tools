"""作業の記録・テストモード・資材マスタの編集履歴 (VER 0.13.30)

**守ること**
- テストモード・記録の保存先フォルダーを変えるには管理者パスワードが要る。ライン名は要らない
- テストモードなら画面の上(と紙面の画面)に札が出る
- 出力・印刷画面・全クリアなどを、版・PC・ライン付きで月ごとのファイルに足す
  (この PC の records と、設定があれば <記録の保存先フォルダー>/<PC名>)。書けなくても止めない
- 明細の履歴の「モード」は、テストモードで出した紙だけ「テストモード」
- 右上の文字を変えたら編集履歴に前・後を残す。管理者パスワードは変えたことだけで、値はどこにも残さない
"""
from __future__ import annotations

import json
import re
import secrets
import sqlite3
import tempfile
import unittest
from pathlib import Path

from modules.packing_details.meisai import (config, db, distribution, edit_history, shared_settings,
                                           slip_history, user_settings, work_record)

from . import _db
from .test_web import HEADERS, TOKEN, _client, _post, _seed

PASSWORD = config.ADMIN_PASSWORD         # 一度も変えていない端末で通る値


def _records_dir(case: unittest.TestCase) -> Path:
    """この試験だけの records(この PC の控え)。"""
    folder = tempfile.TemporaryDirectory(prefix="packing-details-records-")
    case.addCleanup(folder.cleanup)
    case.addCleanup(setattr, config, "RECORDS_DIR", config.RECORDS_DIR)
    config.RECORDS_DIR = Path(folder.name) / "records"
    work_record.reset_for_tests()
    case.addCleanup(work_record.reset_for_tests)
    return Path(folder.name)


def _lines(folder: Path) -> list[dict]:
    out = []
    for path in sorted(folder.glob(f"{work_record.FILE_PREFIX}_*.jsonl")):
        out += [json.loads(t) for t in path.read_text(encoding="utf-8").splitlines() if t.strip()]
    return out


def _history_rows(path: Path) -> list[dict]:
    conn = sqlite3.connect(str(path))
    conn.row_factory = sqlite3.Row
    try:
        return [dict(r) for r in conn.execute(
            f'SELECT * FROM "{edit_history.TABLE}" ORDER BY rowid')]
    finally:
        conn.close()


class _Base(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)
        self.root = _records_dir(self)
        self.client, self.conn = _client(self)

    def _test_mode(self, on: bool, password: str = PASSWORD):
        return _post(self.client, "/api/settings/test-mode", {"enabled": on, "password": password})

    def _output_one(self):
        _seed(self.conn)
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        _post(self.client, "/api/strands", {})
        _post(self.client, "/api/stack-max", {"value": "2"})
        for key in ("1-10", "2-8"):
            _post(self.client, "/api/stack", {"key": key})
        body = _post(self.client, "/api/output", {"confirm": True}).get_json()
        self.assertTrue(body["ok"], body)
        return body


class SettingsTest(_Base):

    def test_テストモードは管理者パスワードが要る(self):
        for wrong in ("", "違う"):
            with self.subTest(password=wrong):
                res = self._test_mode(True, wrong)
                self.assertEqual(res.status_code, 403)
                self.assertEqual(res.get_json()["error"]["field"], "password")
                self.assertFalse(work_record.test_mode())
        res = self._test_mode(True)
        self.assertEqual(res.status_code, 200)
        self.assertTrue(res.get_json()["test_mode"])
        self.assertTrue(work_record.test_mode())
        # 切るのにも要る
        self.assertEqual(self._test_mode(False, "違う").status_code, 403)
        self.assertTrue(work_record.test_mode())
        self.assertFalse(self._test_mode(False).get_json()["test_mode"])
        self.assertFalse(work_record.test_mode())

    def test_記録の保存先フォルダーは管理者パスワードが要る(self):
        target = str(self.root / "共有の記録")
        res = _post(self.client, "/api/settings/record-dir", {"value": target, "password": "違う"})
        self.assertEqual(res.status_code, 403)
        self.assertEqual(work_record.record_dir_setting(), "")
        res = _post(self.client, "/api/settings/record-dir", {"value": target, "password": PASSWORD})
        self.assertEqual(res.status_code, 200, res.get_json())
        self.assertEqual(work_record.record_dir_setting(), target)
        self.assertEqual(res.get_json()["record_dir"],
                         str(Path(target) / work_record.pc_folder_name()))

    def test_ライン名はパスワード無しで保存できる(self):
        res = _post(self.client, "/api/settings/line-name", {"value": " 2ライン "})
        self.assertEqual(res.status_code, 200)
        self.assertEqual(work_record.line_name(), "2ライン")
        self.assertEqual(work_record.who()["line"], "2ライン")
        res = _post(self.client, "/api/settings/line-name", {"value": "あ" * 41})
        self.assertEqual(res.status_code, 400)
        self.assertEqual(work_record.line_name(), "2ライン")

    def test_設定画面に記録の設定が出る(self):
        user_settings.save(config.KEY_LINE_NAME, "3ライン")
        body = self.client.get("/api/settings", headers=HEADERS).get_json()
        self.assertEqual(body["line_name"], "3ライン")
        self.assertFalse(body["test_mode"])
        self.assertEqual(body["records_local"], str(config.RECORDS_DIR))
        local = next(g for g in body["storage"]["groups"] if g["kind"] == "local")
        self.assertIn(str(config.RECORDS_DIR), [r["path"] for r in local["rows"]])

    def test_配布するのは保存先フォルダーだけ(self):
        keys = {k for k, _l, _d in distribution.ITEMS}
        self.assertIn(config.KEY_RECORD_DIR, keys)
        self.assertNotIn(config.KEY_LINE_NAME, keys)
        self.assertNotIn(config.KEY_TEST_MODE, keys)

    def test_文字のfalseはテストモードにしない(self):
        user_settings.save(config.KEY_TEST_MODE, "false")
        self.assertFalse(work_record.test_mode())


class BadgeTest(_Base):

    def _page(self) -> str:
        return self.client.get(f"/meisai?t={TOKEN}").get_data(as_text=True)

    def test_テストモードなら上の帯に札が出る(self):
        badge = re.compile(r'<span class="testmode" id="testModeBadge"[^>]*>')
        self.assertRegex(badge.search(self._page()).group(0), r"\bhidden\b")
        self._test_mode(True)
        self.assertNotRegex(badge.search(self._page()).group(0), r"\bhidden\b")

    def test_紙面の画面にも札が出る(self):
        self._output_one()
        url = f"/report/L5160Z0/1?t={TOKEN}"
        self.assertNotIn('class="testmode"', self.client.get(url).get_data(as_text=True))
        self._test_mode(True)
        self.assertIn('class="testmode"', self.client.get(url).get_data(as_text=True))


class RecordTest(_Base):

    def test_出力すると版_PC_ライン付きで記録が残る(self):
        _post(self.client, "/api/settings/line-name", {"value": "2ライン"})
        shared_root = self.root / "共有の記録"
        _post(self.client, "/api/settings/record-dir",
              {"value": str(shared_root), "password": PASSWORD})
        self._output_one()

        items = [i for i in _lines(config.RECORDS_DIR) if i["kind"] == "出力"]
        self.assertEqual(len(items), 1)
        item = items[0]
        self.assertEqual(item["lot"], "L5160Z0")
        self.assertEqual(item["no"], 1)
        self.assertEqual(item["coils"], ["1-10", "2-8"])
        self.assertEqual(item["weights"], [250, 248])
        self.assertTrue(item["historyId"])
        self.assertEqual(item["version"]["details"], config.app_config.version())
        self.assertEqual(item["who"]["pc"], shared_settings.pc_name())
        self.assertEqual(item["who"]["line"], "2ライン")
        self.assertNotIn("mode", item["who"])
        # 設定の場所にも、PC ごとのフォルダーで同じ行
        shared = _lines(shared_root / work_record.pc_folder_name())
        self.assertIn(item, shared)

    def test_テストモードの記録には印が付く(self):
        self._test_mode(True)
        self._output_one()
        item = [i for i in _lines(config.RECORDS_DIR) if i["kind"] == "出力"][0]
        self.assertEqual(item["who"]["mode"], work_record.TEST_MODE_MARK)

    def test_印刷画面_紙面_全クリアも残る(self):
        self._output_one()
        _post(self.client, "/api/printed", {"lot": "L5160Z0", "nos": [1]})
        self.client.get(f"/report/L5160Z0/1?t={TOKEN}")
        _post(self.client, "/api/clear-all", {"confirm": True})
        kinds = [i["kind"] for i in _lines(config.RECORDS_DIR)]
        for kind in ("出力", "印刷画面", "紙面を開いた", "全クリア"):
            self.assertIn(kind, kinds)
        printed = next(i for i in _lines(config.RECORDS_DIR) if i["kind"] == "印刷画面")
        self.assertEqual((printed["lot"], printed["nos"]), ("L5160Z0", [1]))
        cleared = next(i for i in _lines(config.RECORDS_DIR) if i["kind"] == "全クリア")
        self.assertEqual(cleared["lot"], "L5160Z0")
        self.assertEqual(cleared["cancelled"], 1)

    def test_設定の場所に書けなくても出力は止めない(self):
        blocker = self.root / "ふさがっている"
        blocker.write_text("ファイル", encoding="utf-8")       # フォルダーを作れない
        user_settings.save(config.KEY_RECORD_DIR, str(blocker))
        body = self._output_one()
        self.assertEqual(body["seq_no"], 1)
        self.assertTrue(work_record.status()["shared_problem"])
        self.assertTrue(_lines(config.RECORDS_DIR))           # この PC には書けている
        html = self.client.get(f"/report/records?t={TOKEN}").get_data(as_text=True)
        self.assertIn("設定の保存先に書けていません", html)

    def test_新しい順に読める(self):
        for n in range(5):
            work_record.append("出力", {"lot": "L000000", "no": n + 1})
        items = work_record.recent(3)
        self.assertEqual([i["no"] for i in items], [5, 4, 3])


class SlipModeTest(_Base):

    def _modes(self) -> list[str]:
        return [r[0] for r in self.conn.execute("SELECT モード FROM 明細履歴 ORDER BY 出力日時")]

    def test_テストモードで出した紙だけ印が付く(self):
        self._output_one()
        self.assertEqual(self._modes(), [""])
        self._test_mode(True)
        _post(self.client, "/api/stack", {"key": "1-9"})
        _post(self.client, "/api/stack", {"key": "2-7"})
        _post(self.client, "/api/output", {"confirm": True})
        self.assertEqual(sorted(self._modes()), ["", work_record.TEST_MODE_MARK])

    def test_共有の履歴にも送る_古い表には列を足す(self):
        if sqlite3.sqlite_version_info < (3, 35, 0):
            self.skipTest("DROP COLUMN が使えない SQLite")
        shared_settings.history_dir().mkdir(parents=True, exist_ok=True)
        path = shared_settings.history_dir() / config.HISTORY_DB_NAME
        old = sqlite3.connect(str(path))
        old.executescript(slip_history.SHARED_DDL)
        old.execute("ALTER TABLE 明細履歴 DROP COLUMN モード")      # 前の版が作った表
        old.commit()
        old.close()
        self._test_mode(True)
        self._output_one()
        result = slip_history.send_pending(self.conn)
        self.assertTrue(result.ok, result.error)
        conn = sqlite3.connect(str(path))
        try:
            self.assertEqual(conn.execute("SELECT モード FROM 明細履歴").fetchall(),
                             [(work_record.TEST_MODE_MARK,)])
        finally:
            conn.close()
        found = slip_history.search(self.conn, "2000-01-01", "2999-12-31")
        self.assertEqual(found.rows[0]["モード"], work_record.TEST_MODE_MARK)

    def test_古い手元のDBにも列を足す(self):
        if sqlite3.sqlite_version_info < (3, 35, 0):
            self.skipTest("DROP COLUMN が使えない SQLite")
        conn = _db.memory_db()
        self.addCleanup(conn.close)
        conn.execute("ALTER TABLE 明細履歴 DROP COLUMN モード")
        conn.commit()
        db.apply_schema(conn)
        cols = [r[1] for r in conn.execute("PRAGMA table_info(明細履歴)")]
        self.assertEqual(tuple(c for c in cols if c not in slip_history.LOCAL_ONLY),
                         slip_history.HEADER_COLUMNS)


class EditHistoryTest(_Base):

    def setUp(self):
        super().setUp()
        self.master = _db.master_file(self)

    def test_右上の文字を変えたら前と後を残す(self):
        _post(self.client, "/api/settings/line-name", {"value": "2ライン"})
        self._test_mode(True)
        res = _post(self.client, "/api/settings/qa-mark",
                    {"value": "NLM.TEST.QA", "password": PASSWORD})
        self.assertEqual(res.status_code, 200, res.get_json())
        _post(self.client, "/api/settings/qa-mark", {"reset": True, "password": PASSWORD})
        rows = _history_rows(self.master)
        self.assertEqual([(r["操作"], r["前"], r["後"]) for r in rows],
                         [("変更", "NLM.NAGOYA.QA", "NLM.TEST.QA"),
                          ("既定に戻す", "NLM.TEST.QA", "NLM.NAGOYA.QA")])
        first = rows[0]
        self.assertEqual(first["表"], shared_settings.MASTER_TABLE)
        self.assertEqual(first["列"], shared_settings.MASTER_VALUE_COLUMN)
        self.assertEqual(first["PC"], shared_settings.pc_name())
        self.assertEqual(first["ライン"], "2ライン")
        self.assertEqual(first["モード"], work_record.TEST_MODE_MARK)

    def test_管理者パスワードは変えたことだけ残す(self):
        new = "pw" + secrets.token_hex(6)
        res = _post(self.client, "/api/settings/admin-password",
                    {"current": PASSWORD, "new": new, "confirm": new})
        self.assertEqual(res.status_code, 200, res.get_json())
        res = _post(self.client, "/api/settings/admin-password", {"reset": True, "current": new})
        self.assertEqual(res.status_code, 200, res.get_json())
        rows = _history_rows(self.master)
        self.assertEqual([(r["操作"], r["前"], r["後"]) for r in rows],
                         [("変更", edit_history.HIDDEN, edit_history.HIDDEN),
                          ("初期化", edit_history.HIDDEN, edit_history.HIDDEN)])
        # 値はどこにも無い(編集履歴・記録・記録の画面)
        dumped = json.dumps(rows, ensure_ascii=False)
        records = "".join(p.read_text(encoding="utf-8")
                          for p in config.RECORDS_DIR.glob("*.jsonl"))
        page = self.client.get(f"/report/records?t={TOKEN}").get_data(as_text=True)
        for secret in (PASSWORD, new):
            self.assertNotIn(secret, dumped)
            self.assertNotIn(secret, records)
            self.assertNotIn(secret, page)

    def test_記録の画面に出る(self):
        self._output_one()
        _post(self.client, "/api/settings/qa-mark", {"value": "NLM.TEST.QA", "password": PASSWORD})
        res = self.client.get(f"/report/records?t={TOKEN}")
        self.assertEqual(res.status_code, 200)
        html = res.get_data(as_text=True)
        self.assertIn("作業の記録", html)
        self.assertIn("刷った証拠にはなりません", html)
        self.assertIn("L5160Z0", html)
        self.assertIn("NLM.TEST.QA", html)                    # 編集履歴
        self.assertIn(edit_history.TABLE, html)

    def test_記録の画面もトークンが要る(self):
        self.assertEqual(self.client.get("/report/records").status_code, 401)


if __name__ == "__main__":
    unittest.main()
