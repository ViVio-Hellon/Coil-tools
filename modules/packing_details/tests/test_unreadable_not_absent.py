"""「読めない」を「無い」と取り違えない(統合 1.2.61)

別のツールで、共有ファイルが読めなかったのを「実績の表が無い」と取り違え、手元の実績を送り直して重複を
大量に作った事案があった。梱包明細で同じ取り違えをしていた所を直した。どれも、**読めない・届かないときは
書かない・消さない・覚えない**。
"""
from __future__ import annotations

import json
import os
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from modules.packing_details.meisai import (admin_password, config, data_sync, db, distribution, history_repo,
                                           lot_lookup, shared_settings, slip_history, source_db,
                                           user_settings)

from . import _db
from .test_ls4_results import RESULT_COLUMNS, _row, _warisu, _ym


def _lots(conn) -> list[str]:
    return sorted(r[0] for r in conn.execute('SELECT "ロット番号" FROM "仕掛ロット"'))


class FirstPathUnreachableTest(unittest.TestCase):
    """第1パスのフォルダに一瞬届かないとき、第2パスの別のファイルで手元を総入れ替えしない。"""

    def setUp(self):
        _db.settings_file(self)
        self.conn = _db.case_db(self)
        root = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, root, True)
        self.first, self.second = root / "first", root / "second"
        self.first.mkdir()
        self.second.mkdir()
        user_settings.save(config.KEY_LOT_DB_DIR, str(self.first))
        user_settings.save(config.KEY_LOT_DB_DIR2, str(self.second))
        user_settings.save(config.KEY_LOT_DB2_FILES["仕掛ロット"], "LOT2")
        _db.make_source(self.first / "SIKALOT.sqlite3", _db.LOT_COLUMNS, [_db.lot_row("A000001")])
        _db.make_source(self.second / "LOT2.sqlite3", _db.LOT_COLUMNS, [_db.lot_row("B000001")])

    def import_lot(self):
        return data_sync.import_all(self.conn, force=True, only=["仕掛ロット"])

    def test_前は第1パスから取り込んでいたなら_届かないあいだは前の中身のまま(self):
        self.import_lot()
        self.assertEqual(_lots(self.conn), ["A000001", "B000001"])
        self.first.rename(self.first.with_name("first_unreachable"))         # 届かない
        result = self.import_lot()
        self.assertEqual(_lots(self.conn), ["A000001", "B000001"], "第2パスだけで入れ替えない")
        self.assertTrue(any("届かないので" in e for e in result.errors), result.errors)

    def test_フォルダに届いてファイルだけ無いなら_いままでどおり第2パスから(self):
        self.import_lot()
        (self.first / "SIKALOT.sqlite3").unlink()
        result = self.import_lot()
        self.assertEqual(_lots(self.conn), ["B000001"])
        self.assertFalse(any("届かないので" in e for e in result.errors), result.errors)

    def test_前から代わりのファイルを使っている端末はいままでどおり(self):
        self.first.rename(self.first.with_name("first_unreachable"))
        self.import_lot()
        self.assertEqual(_lots(self.conn), ["B000001"])


class ResultsEmptyTest(unittest.TestCase):
    """実績の形式の元の表が**空**(書き直し中)なら、手元を空にしない。今月の行が無いのとは別。"""

    def setUp(self):
        _db.settings_file(self)
        self.conn = _db.case_db(self)
        self.konpo = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, self.konpo, True)
        user_settings.save(config.KEY_KONPO_DB_DIR, str(self.konpo))
        user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")

    def test_元の表が空なら前の中身のまま(self):
        _db.make_source(self.konpo / "実績_LS4.sqlite3", RESULT_COLUMNS,
                        [_row("A000001", "2", "6", f"{_ym()}-05 08:00:00")])
        data_sync.import_all(self.conn, force=True, only=config.KONPO_DB_FILES)
        self.assertEqual(_warisu(self.conn), {"A000001": (2, 6)})
        (self.konpo / "実績_LS4.sqlite3").unlink()                   # 書き直し中で空
        _db.make_source(self.konpo / "実績_LS4.sqlite3", RESULT_COLUMNS, [])
        result = data_sync.import_all(self.conn, force=True, only=config.KONPO_DB_FILES)
        self.assertEqual(_warisu(self.conn), {"A000001": (2, 6)})
        self.assertTrue(any("表が空" in n and "前の中身のまま" in n for n in result.notes), result.notes)


class UserSettingsTest(unittest.TestCase):
    """設定ファイルが読めないのを「何も設定していない」と取り違えて、ほかの設定を消さない。"""

    def setUp(self):
        self.path = _db.settings_file(self)
        user_settings.save(config.KEY_LOT_DB_DIR, r"\\server\lot")

    def test_読めないときは保存しない(self):
        before = self.path.read_text(encoding="utf-8")
        with mock.patch.object(Path, "read_text", side_effect=PermissionError("押さえられています")):
            self.assertFalse(user_settings.save(config.KEY_THEME if hasattr(config, "KEY_THEME")
                                                else "x_test", "dark"))
        self.assertEqual(self.path.read_text(encoding="utf-8"), before, "書き換えていない")
        self.assertEqual(user_settings.get(config.KEY_LOT_DB_DIR), r"\\server\lot")

    def test_壊れたファイルは退けてから保存する(self):
        self.path.write_text("{壊れた", encoding="utf-8")
        self.assertTrue(user_settings.save("x_test", 1))
        self.assertEqual(json.loads(self.path.read_text(encoding="utf-8")), {"x_test": 1})
        aside = [p for p in self.path.parent.iterdir() if ".壊れていた_" in p.name]
        self.assertEqual(len(aside), 1)
        self.assertEqual(aside[0].read_text(encoding="utf-8"), "{壊れた")

    def test_配布設定も読めないあいだは上書きしない(self):
        bundle = distribution.Bundle(settings={config.KEY_LOT_DB_DIR: r"\\other\lot"})
        with mock.patch.object(user_settings, "load_strict",
                               side_effect=user_settings.SettingsUnreadable("押さえられています")):
            result = distribution._apply(bundle, overwrite=False)
        self.assertFalse(result.ok)
        self.assertEqual(user_settings.get(config.KEY_LOT_DB_DIR), r"\\server\lot")


class AdminPasswordTest(unittest.TestCase):
    def test_控えのJSONが読めなければ既定のパスワードで通さない(self):
        bad = shared_settings.Snapshot({}, "shared", json_problem="共有の設定(JSON)を読めません")
        self.assertFalse(admin_password.verify(config.ADMIN_PASSWORD, bad))
        good = shared_settings.Snapshot({}, "shared")
        self.assertTrue(admin_password.verify(config.ADMIN_PASSWORD, good))


class SharedSettingsTest(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)

    def test_マスタのフォルダに届かないのを無いとみなさない(self):
        with mock.patch.object(source_db, "find_strict", side_effect=OSError("ネットワーク名が見つかりません")):
            master = shared_settings._read_master(config.SHARED_DIR)
        self.assertTrue(master.problem)
        self.assertEqual(master.state, "error")

    def test_前は有った控えのJSONが見えないなら作り直さない(self):
        folder = shared_settings.json_dir()
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / shared_settings.FILE_NAME
        path.write_text(json.dumps({"format": 1, "qa_mark": "NLM.TEST",
                                    "admin_password": admin_password._encode("pw", b"salt", 1000)}),
                        encoding="utf-8")
        snap = shared_settings.read()
        self.assertTrue(snap.get("admin_password"))
        path.unlink()                                                 # 置き換え中で一瞬見えない
        with self.assertRaises(shared_settings.SharedError):
            shared_settings.update({"qa_mark": "NLM.OTHER"})
        self.assertFalse(path.exists(), "空から作り直していない(パスワードが消えない)")


class SlipHistoryTest(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)

    def test_確かめられないのを無いとみなして作らない(self):
        folder = shared_settings.history_dir()
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / config.HISTORY_DB_NAME
        with mock.patch.object(Path, "stat", side_effect=PermissionError("確かめられません")):
            with self.assertRaises(slip_history.HistoryError):
                slip_history.ensure_shared()
        self.assertFalse(path.exists())

    def test_無いと確かめたときだけ作る_ふだんは作らずに開く(self):
        folder = shared_settings.history_dir()
        folder.mkdir(parents=True, exist_ok=True)
        path = folder / config.HISTORY_DB_NAME
        with self.assertRaises(sqlite3.OperationalError):
            slip_history._open_shared(path)                           # 作らない
        self.assertFalse(path.exists())
        self.assertTrue(slip_history.ensure_shared())                 # 無いと確かめて作る
        self.assertTrue(path.exists())


class LocalReadTest(unittest.TestCase):
    def test_手元のDBを読めなければ二重出力の確かめを素通りしない(self):
        conn = _db.case_db(self)
        with mock.patch.object(db, "fetch_all", return_value=None):
            with self.assertRaises(sqlite3.OperationalError):
                history_repo.is_duplicate(conn, "A000001-1-1")
            with self.assertRaises(sqlite3.OperationalError):
                history_repo.duplicates(conn, "A000001", ["1-1"])


class LotLookupTest(unittest.TestCase):
    def setUp(self):
        lot_lookup.forget()
        self.addCleanup(lot_lookup.forget)
        root = Path(tempfile.mkdtemp())
        self.addCleanup(shutil.rmtree, root, True)
        self.path = root / "SIKALOT.sqlite3"
        c = sqlite3.connect(self.path)
        c.execute('CREATE TABLE "仕掛" ("ﾛｯﾄ番号" TEXT, "検査番号" TEXT)')
        c.execute('INSERT INTO "仕掛" VALUES (?, ?)', ("A000001", "N6X0945"))
        c.commit()
        c.close()
        p = mock.patch.object(lot_lookup, "sources",
                              lambda table="仕掛当工程": [self.path] if table == "仕掛ロット" else [])
        p.start()
        self.addCleanup(p.stop)

    def test_読めないのを表が無いと覚えない(self):
        with mock.patch.object(source_db, "list_tables_strict",
                               side_effect=source_db.SourceError("写せません")):
            r = lot_lookup.lookup("A000001")
        self.assertFalse(r["found"])
        self.assertIn("読めません", r["reason"])
        # 読めるようになったら(ファイルは変わっていなくても)引ける
        self.assertEqual(lot_lookup.lookup("A000001")["kensa"], "N6X0945")


if __name__ == "__main__":
    unittest.main()
