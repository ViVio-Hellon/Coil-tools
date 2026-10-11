"""仕掛台帳の2つ目(統合 1.2.7。python-web-tools VER4.4.0・VER4.8.15 と同じ決まり)

    1つ目のフォルダ … 決まった名前(SIKALOT / SIKAHIKI / SIKAODR)で探す
    2つ目のフォルダ … 2つ目の名前(表ごとに設定。空なら1つ目と同じ名前)だけを探す

    ファイルが1つ目に無い      → そのファイルは2つ目から読む
    1つ目にロット・受注が無い  → 1つ目を取り込んだあと、2つ目にだけあるものを足す
    両方にある                 → 1つ目を正とする(2つ目の行は足さない)

名前は変えてよいが、中身の要る列はそろえる約束(現場)。足りない列は設定画面の「いまの状態」に出る。
"""
from __future__ import annotations

import os
import tempfile
import time
import unittest
from pathlib import Path

from modules.packing_details.meisai import config, data_sync, distribution, source_db, user_settings

from . import _db
from .test_web import HEADERS, _client, _post


def _dirs() -> tuple[Path, Path]:
    root = Path(tempfile.mkdtemp())
    first, second = root / "first", root / "second"
    first.mkdir()
    second.mkdir()
    return first, second


def _lot(path: Path, *lots: str) -> None:
    _db.make_source(path, _db.LOT_COLUMNS, [_db.lot_row(n) for n in lots])


def _hiki(path: Path, *pairs: tuple[str, str]) -> None:
    _db.make_source(path, _db.HIKI_COLUMNS,
                    [{"ﾛｯﾄ番号": lot, "引当番号": "1", "受注番号": odr} for lot, odr in pairs])


def _odr(path: Path, *orders: tuple[str, str]) -> None:
    _db.make_source(path, _db.ODR_COLUMNS,
                    [{"受注番号": odr, "包装仕様NO": spec} for odr, spec in orders])


def _lots(conn) -> list[str]:
    return sorted(r[0] for r in conn.execute('SELECT "ロット番号" FROM "仕掛ロット"'))


class SecondDirTest(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)
        self.conn = _db.case_db(self)
        self.first, self.second = _dirs()
        user_settings.save(config.KEY_LOT_DB_DIR, str(self.first))

    def use_second(self, **names: str) -> None:
        user_settings.save(config.KEY_LOT_DB_DIR2, str(self.second))
        for table, name in names.items():
            user_settings.save(config.KEY_LOT_DB2_FILES[table], name)

    def import_lot(self):
        return data_sync.import_all(self.conn, force=True, only=config.LOT_DB_FILES)

    def test_使わなければ今までどおり(self):
        _lot(self.first / "SIKALOT.sqlite3", "A000001")
        self.assertIsNone(config.lot_db_dir2())
        self.assertEqual(data_sync.find_second_lots(), {})
        self.import_lot()
        self.assertEqual(_lots(self.conn), ["A000001"])

    def test_1つ目に無いファイルは2つ目の名前で読む(self):
        self.use_second(仕掛ロット="LOT2")
        _lot(self.second / "LOT2.sqlite3", "B000001")
        _hiki(self.first / "SIKAHIKI.sqlite3", ("B000001", "O1"))
        _odr(self.first / "SIKAODR.sqlite3", ("O1", "1C0001"))
        found = data_sync.find_sources()
        self.assertEqual(found["仕掛ロット"], self.second / "LOT2.sqlite3")
        self.assertEqual(found["仕掛引当"], self.first / "SIKAHIKI.sqlite3")
        result = self.import_lot()
        self.assertTrue(result.ok, result.errors)
        self.assertEqual(_lots(self.conn), ["B000001"])

    def test_名前を変えたら2つ目の同じ名前のファイルは読まない(self):
        """SIKALOT は名前が同じでも中身が違う(圧縮版)ことがあった。取り違えない。"""
        self.use_second(仕掛ロット="LOT2")
        _lot(self.second / "SIKALOT.sqlite3", "X999999")
        self.assertNotIn("仕掛ロット", data_sync.find_second_lots())

    def test_名前が空なら1つ目と同じ名前_拡張子なしは_db_も見る(self):
        self.use_second(仕掛受注="ODR2")
        _lot(self.second / "SIKALOT.sqlite3", "C000001")
        _odr(self.second / "ODR2.db", ("O9", "1C0009"))
        found = data_sync.find_second_lots()
        self.assertEqual(found["仕掛ロット"], self.second / "SIKALOT.sqlite3")
        self.assertEqual(found["仕掛受注"], self.second / "ODR2.db")
        self.assertEqual(config.lot_db2_files()["仕掛引当"], "SIKAHIKI.sqlite3")

    def test_1つ目に無いロット受注を2つ目から足す_両方にあれば1つ目(self):
        self.use_second(仕掛ロット="LOT2", 仕掛受注="ODR2")
        _lot(self.first / "SIKALOT.sqlite3", "A000001")
        _lot(self.second / "LOT2.sqlite3", "A000001", "D000001")
        _odr(self.first / "SIKAODR.sqlite3", ("O1", "1C0001"))
        _odr(self.second / "ODR2.sqlite3", ("O1", "2X9999"), ("O2", "1C0002"))
        result = self.import_lot()
        self.assertEqual(_lots(self.conn), ["A000001", "D000001"])
        specs = dict(self.conn.execute('SELECT "受注番号", "包装仕様NO" FROM "仕掛受注"').fetchall())
        self.assertEqual(specs, {"O1": "1C0001", "O2": "1C0002"}, "O1 は1つ目の値")
        self.assertTrue(any("2つ目" in n for n in result.notes), result.notes)

    def test_2つ目だけ変わっても取り込み直す(self):
        self.use_second(仕掛ロット="LOT2")
        _lot(self.first / "SIKALOT.sqlite3", "A000001")
        _lot(self.second / "LOT2.sqlite3", "A000001")
        data_sync.import_all(self.conn, only=config.LOT_DB_FILES)
        self.assertEqual(_lots(self.conn), ["A000001"])
        (self.second / "LOT2.sqlite3").unlink()
        _lot(self.second / "LOT2.sqlite3", "A000001", "E000001")
        later = time.time() + 5
        os.utime(self.second / "LOT2.sqlite3", (later, later))
        data_sync.import_all(self.conn, only=config.LOT_DB_FILES)       # force なし
        self.assertEqual(_lots(self.conn), ["A000001", "E000001"])

    def test_鍵の列が無い2つ目からは足さない(self):
        self.use_second(仕掛ロット="LOT2")
        _lot(self.first / "SIKALOT.sqlite3", "A000001")
        _db.make_source(self.second / "LOT2.sqlite3", ("用途名",), [{"用途名": "x"}])
        result = self.import_lot()
        self.assertEqual(_lots(self.conn), ["A000001"])
        self.assertTrue(any("鍵の列" in e for e in result.errors), result.errors)
        self.assertEqual(data_sync.second_missing_columns(self.second / "LOT2.sqlite3", "仕掛ロット")[:1],
                         ["ﾛｯﾄ番号"])

    def test_1つ目と同じフォルダなら2つ目としては使わない(self):
        user_settings.save(config.KEY_LOT_DB_DIR2, str(self.first))
        self.assertIsNone(config.lot_db_dir2())

    def test_見つからない文は2つ目の名前で言う(self):
        self.use_second(仕掛ロット="LOT2")
        result = self.import_lot()
        self.assertTrue(any("LOT2" in e for e in result.errors), result.errors)

    def test_find_named(self):
        (self.second / "台帳.bak").write_text("x")
        self.assertEqual(source_db.find_named(self.second, "台帳.bak"), self.second / "台帳.bak")
        self.assertIsNone(source_db.find_named(self.second, ""))


class SecondDirScreenTest(unittest.TestCase):
    """設定画面(/api/settings/lot2)と配布設定。"""

    def setUp(self):
        _db.settings_file(self)
        self.client, _conn = _client(self)
        self.first, self.second = _dirs()
        user_settings.save(config.KEY_LOT_DB_DIR, str(self.first))

    def test_保存して状態を見る_足りない列を出す(self):
        _db.make_source(self.second / "LOT2.sqlite3", ("ﾛｯﾄ番号",), [{"ﾛｯﾄ番号": "A000001"}])
        res = _post(self.client, "/api/settings/lot2",
                    {"dir": str(self.second), "files": {"仕掛ロット": "LOT2"}})
        self.assertEqual(res.status_code, 200, res.get_json())
        body = res.get_json()
        rows = {r["table"]: r for r in body["lot_second"]}
        self.assertTrue(rows["仕掛ロット"]["found"])
        self.assertIn("用途名", rows["仕掛ロット"]["missing_columns"])
        self.assertEqual(rows["仕掛引当"]["name"], "SIKAHIKI.sqlite3")
        self.assertFalse(rows["仕掛引当"]["found"])
        self.assertEqual(body["lot_db_dir2_setting"], str(self.second))
        names = {f["table"]: f["name"] for f in body["lot_files2"]}
        self.assertEqual(names, {"仕掛ロット": "LOT2", "仕掛引当": "", "仕掛受注": ""})
        self.assertEqual(self.client.get("/api/settings/lot2", headers=HEADERS)
                         .get_json()["lot_second"][0]["name"], "LOT2")

    def test_ファイル名にフォルダや使えない文字は断る(self):
        for bad in ("sub\\LOT2", "a/b", "x:y", "LOT?"):
            with self.subTest(bad=bad):
                res = _post(self.client, "/api/settings/lot2",
                            {"dir": str(self.second), "files": {"仕掛ロット": bad}})
                self.assertEqual(res.status_code, 400)
                self.assertEqual(res.get_json()["error"]["code"], "bad_name")

    def test_空にすれば2つ目を使わない(self):
        _post(self.client, "/api/settings/lot2", {"dir": str(self.second), "files": {}})
        res = _post(self.client, "/api/settings/lot2", {"dir": "", "files": {}})
        self.assertIn("使いません", res.get_json()["note"])
        self.assertIsNone(config.lot_db_dir2())

    def test_配布設定に入る(self):
        keys = {k for k, _, _ in distribution.ITEMS}
        self.assertIn(config.KEY_LOT_DB_DIR2, keys)
        self.assertTrue(set(config.KEY_LOT_DB2_FILES.values()) <= keys)
        self.assertEqual(distribution._check(config.KEY_LOT_DB2_FILES["仕掛ロット"], "LOT2"), "")
        self.assertTrue(distribution._check(config.KEY_LOT_DB2_FILES["仕掛ロット"], "a/b"))
        self.assertTrue(distribution._check(config.KEY_LOT_DB2_FILES["仕掛ロット"], 3))


if __name__ == "__main__":
    unittest.main()
