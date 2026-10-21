"""仕掛の第1パス・第2パス(統合 1.2.28・現場の指示)

    第1パス  仕掛台帳のフォルダ   … SIKALOT / SIKAHIKI / SIKAODR(名前は固定)
             梱包課共有のフォルダ … LS4LOT(**名前も変えられる**。仕掛台帳とは別のフォルダ)
    第2パス  4ファイル共通のフォルダ … SIKALOT / SIKAHIKI / SIKAODR / LS4LOT(**4つとも名前を変えられる**)

第1パスに無いファイルは第2パスから読み、第1パスに無いロットは第2パスから足す(両方にあれば第1パス)。
"""
from __future__ import annotations

import tempfile
import unittest
from pathlib import Path

from modules.packing_details.meisai import config, data_sync, distribution, user_settings

from . import _db
from .test_web import HEADERS, _client, _post


def _ls4(path: Path, *lots: str) -> None:
    _db.make_source(path, _db.TOKOTEI_COLUMNS, [_db.tokotei_row(n) for n in lots])


def _tokotei(conn) -> list[str]:
    return sorted(r[0] for r in conn.execute('SELECT "ロット番号" FROM "仕掛当工程"'))


class LotPathsTest(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)
        self.conn = _db.case_db(self)
        root = Path(tempfile.mkdtemp())
        self.lot, self.konpo, self.second = root / "lot", root / "konpo", root / "second"
        for d in (self.lot, self.konpo, self.second):
            d.mkdir()
        user_settings.save(config.KEY_LOT_DB_DIR, str(self.lot))
        user_settings.save(config.KEY_KONPO_DB_DIR, str(self.konpo))

    def import_konpo(self):
        return data_sync.import_all(self.conn, force=True, only=config.KONPO_DB_FILES)

    # ---- 第1パス: LS4LOT の名前 ----
    def test_第1パスのLS4LOTは名前を変えられる(self):
        _ls4(self.konpo / "LS4LOT.sqlite3", "X999999")
        _ls4(self.konpo / "当工程.sqlite3", "A000001")
        user_settings.save(config.KEY_KONPO_DB_FILE, "当工程")
        self.assertEqual(data_sync.find_sources()["仕掛当工程"], self.konpo / "当工程.sqlite3")
        self.assertEqual(config.source_files()["仕掛当工程"], "当工程")
        self.import_konpo()
        self.assertEqual(_tokotei(self.conn), ["A000001"], "名前を変えたら LS4LOT.sqlite3 は読まない")

    def test_名前が空ならLS4LOT_拡張子まで書けばその名前そのもの(self):
        _ls4(self.konpo / "LS4LOT.sqlite3", "A000001")
        self.assertEqual(data_sync.find_sources()["仕掛当工程"], self.konpo / "LS4LOT.sqlite3")
        _ls4(self.konpo / "ls4_copy.bak", "B000001")
        user_settings.save(config.KEY_KONPO_DB_FILE, "ls4_copy.bak")
        self.assertEqual(data_sync.find_sources()["仕掛当工程"], self.konpo / "ls4_copy.bak")

    def test_見つからない文は変えた名前で言う(self):
        user_settings.save(config.KEY_KONPO_DB_FILE, "当工程")
        result = self.import_konpo()
        self.assertTrue(any("当工程" in e for e in result.errors), result.errors)

    # ---- 第2パス: LS4LOT も ----
    def test_第1パスにLS4LOTが無ければ第2パスの名前で読む(self):
        user_settings.save(config.KEY_LOT_DB_DIR2, str(self.second))
        user_settings.save(config.KEY_LOT_DB2_FILES["仕掛当工程"], "LS4_2")
        _ls4(self.second / "LS4LOT.sqlite3", "X999999")
        _ls4(self.second / "LS4_2.sqlite3", "B000001")
        self.assertEqual(data_sync.find_sources()["仕掛当工程"], self.second / "LS4_2.sqlite3")
        self.import_konpo()
        self.assertEqual(_tokotei(self.conn), ["B000001"])

    def test_第1パスに無いロットは第2パスのLS4LOTから足す_両方にあれば第1パス(self):
        user_settings.save(config.KEY_LOT_DB_DIR2, str(self.second))
        _ls4(self.konpo / "LS4LOT.sqlite3", "A000001")
        _db.make_source(self.second / "LS4LOT.sqlite3", _db.TOKOTEI_COLUMNS,
                        [_db.tokotei_row("A000001", yoko="99"), _db.tokotei_row("C000001")])
        result = self.import_konpo()
        self.assertEqual(_tokotei(self.conn), ["A000001", "C000001"])
        yoko = self.conn.execute('SELECT "当工程設計_横割数" FROM "仕掛当工程"'
                                 ' WHERE "ロット番号" = ?', ("A000001",)).fetchone()[0]
        self.assertEqual(str(yoko), "12", "両方にあるロットは第1パスの値")
        self.assertTrue(any("第2パス" in n and "仕掛当工程" in n for n in result.notes), result.notes)

    def test_第2パスは4ファイルとも名前を変えられる(self):
        self.assertEqual(set(config.KEY_LOT_DB2_FILES),
                         {"仕掛ロット", "仕掛引当", "仕掛受注", "仕掛当工程"})
        user_settings.save(config.KEY_KONPO_DB_FILE, "当工程")
        self.assertEqual(config.lot_db2_files()["仕掛当工程"], "当工程",
                         "第2パスの名前が空なら第1パスと同じ名前")


class LotPathsScreenTest(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)
        self.client, _conn = _client(self)
        root = Path(tempfile.mkdtemp())
        self.konpo = root / "konpo"
        self.konpo.mkdir()

    def test_設定画面で第1パスのLS4LOTの名前を保存する(self):
        res = _post(self.client, "/api/settings",
                    {"key": config.KEY_KONPO_DB_DIR, "value": str(self.konpo), "file": "当工程"})
        self.assertEqual(res.status_code, 200, res.get_json())
        self.assertEqual(user_settings.get(config.KEY_KONPO_DB_FILE), "当工程")
        state = self.client.get("/api/settings", headers=HEADERS).get_json()
        self.assertEqual(state["konpo_file_setting"], "当工程")
        self.assertEqual(state["konpo_files"], ["当工程"])
        self.assertEqual(state["files"]["仕掛当工程"], "当工程")

    def test_名前にフォルダや使えない文字は断る(self):
        for bad in ("sub\\LS4", "a/b", "x:y"):
            with self.subTest(bad=bad):
                res = _post(self.client, "/api/settings",
                            {"key": config.KEY_KONPO_DB_DIR, "value": str(self.konpo), "file": bad})
                self.assertEqual(res.status_code, 400)
                self.assertEqual(res.get_json()["error"]["field"], "file")

    def test_第2パスの状態に4ファイル_第1パスと同じファイルはそう言う(self):
        lot = self.konpo.parent / "lot"
        lot.mkdir()
        user_settings.save(config.KEY_LOT_DB_DIR, str(lot))
        _db.make_source(lot / "SIKALOT.sqlite3", _db.LOT_COLUMNS, [_db.lot_row("A000001")])
        body = _post(self.client, "/api/settings/lot2", {"dir": str(lot), "files": {}}).get_json()
        rows = {r["table"]: r for r in body["lot_second"]}
        self.assertEqual(set(rows), {"仕掛ロット", "仕掛引当", "仕掛受注", "仕掛当工程"})
        self.assertTrue(rows["仕掛ロット"]["same_as_first"])
        self.assertFalse(rows["仕掛引当"]["same_as_first"])

    def test_配布設定に入る(self):
        keys = {k for k, _, _ in distribution.ITEMS}
        self.assertIn(config.KEY_KONPO_DB_FILE, keys)
        self.assertIn(config.KEY_LOT_DB2_FILES["仕掛当工程"], keys)
        self.assertEqual(distribution._check(config.KEY_KONPO_DB_FILE, "当工程"), "")
        self.assertTrue(distribution._check(config.KEY_KONPO_DB_FILE, "a/b"))


if __name__ == "__main__":
    unittest.main()
