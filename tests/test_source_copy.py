"""取り込み元は、読む前に必ず手元へ写す(統合 1.2.28・現場の指示)

梱包明細・資材計算の `source_db` が読み取り専用で開くとき、共有の上のファイルを
直接開かないことを確かめる。写し方は `common/source_copy.py` の1か所。
"""
import os
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from common import source_copy
from modules.packing_details.meisai import source_db as details_db
from modules.packing_material_calculation.coil_tool import source_db as material_db


def _make(path: Path, rows=(("A",),), *, wal=False):
    conn = sqlite3.connect(str(path))
    if wal:
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA wal_autocheckpoint=0")
    conn.execute("CREATE TABLE t (v TEXT)")
    conn.executemany("INSERT INTO t VALUES (?)", rows)
    conn.commit()
    return conn


class LocalCopyTest(unittest.TestCase):
    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.share = self.tmp / "share"
        self.share.mkdir()
        source_copy.forget()

    def tearDown(self):
        source_copy.forget()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_写しは手元の一時フォルダに置き_中身が同じなら写し直さない(self):
        src = self.share / "SIKALOT.sqlite3"
        _make(src).close()
        first = source_copy.local_copy(src)
        self.assertNotEqual(first.parent, self.share)
        self.assertEqual(first.parent, source_copy.copy_dir())
        self.assertEqual(source_copy.local_copy(src), first)
        conn = sqlite3.connect(str(src))
        conn.execute("INSERT INTO t VALUES ('B')")
        conn.commit()
        conn.close()
        os.utime(src, ns=(1, 10**18))                         # 確実に更新時刻を変える
        self.assertNotEqual(source_copy.local_copy(src), first)

    def test_写している間に書き換わったら写し直す(self):
        src = self.share / "SIKAHIKI.sqlite3"
        _make(src).close()
        real = shutil.copyfile
        calls = []

        def copy_and_touch(a, b, *args, **kw):
            out = real(a, b, *args, **kw)
            if not calls:
                os.utime(src, ns=(2, 2 * 10**18))             # 1回目の写しの最中に書き換わった
            calls.append(a)
            return out

        with mock.patch.object(source_copy.shutil, "copyfile", side_effect=copy_and_touch), \
             mock.patch.object(source_copy, "RETRY_WAIT_SEC", 0):
            dest = source_copy.local_copy(src)
        self.assertEqual(len(calls), 2)
        self.assertTrue(dest.exists())

    def test_ずっと書き換わっていれば写せないと言う(self):
        src = self.share / "SIKAODR.sqlite3"
        _make(src).close()
        n = iter(range(3, 100))
        real = shutil.copyfile

        def copy_and_touch(a, b, *args, **kw):
            out = real(a, b, *args, **kw)
            os.utime(src, ns=(1, next(n) * 10**18))
            return out

        with mock.patch.object(source_copy.shutil, "copyfile", side_effect=copy_and_touch), \
             mock.patch.object(source_copy, "RETRY_WAIT_SEC", 0):
            with self.assertRaises(source_copy.CopyError):
                source_copy.local_copy(src)


class ReadOnlyGoesThroughCopyTest(unittest.TestCase):
    """両機能の source_db: 読むときは共有の上のファイルを sqlite3 で直接開かない。"""

    def setUp(self):
        self.tmp = Path(tempfile.mkdtemp())
        self.share = self.tmp / "share"
        self.share.mkdir()
        source_copy.forget()

    def tearDown(self):
        source_copy.forget()
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _opened_paths(self, fn):
        seen = []
        real = sqlite3.connect

        def spy(target, *args, **kw):
            seen.append(str(target))
            return real(target, *args, **kw)

        with mock.patch.object(sqlite3, "connect", side_effect=spy):
            result = fn()
        return result, seen

    def test_両機能とも_読む関数はみな写しを開く(self):
        src = self.share / "LS4LOT.sqlite3"
        _make(src, rows=(("X",), ("Y",))).close()
        for mod in (details_db, material_db):
            with self.subTest(mod=mod.__name__):
                for fn in (lambda: mod.read_table(src, "t"),
                           lambda: mod.list_tables(src),
                           lambda: mod.columns(src, "t"),
                           lambda: mod.read_query(src, "SELECT v FROM t"),
                           lambda: mod.is_readable(src),
                           lambda: mod.probe(src)):
                    _result, seen = self._opened_paths(fn)
                    self.assertTrue(seen, "何も開いていない")
                    for target in seen:
                        self.assertNotIn(str(self.share), target.replace("file:", ""),
                                         f"共有の上のファイルを直接開いた: {target}")
                self.assertEqual([r["v"] for r in mod.read_table(src, "t")], ["X", "Y"])
                self.assertEqual(mod.probe(src).opened_by, mod.WAY_COPY)

    def test_WALで書きかけの分も写しから読める(self):
        src = self.share / "SIKALOT.sqlite3"
        writer = _make(src, rows=(("1",), ("2",)), wal=True)    # 閉じずに -wal に残したまま
        try:
            self.assertTrue(Path(str(src) + "-wal").exists())
            rows = details_db.read_table(src, "t")
            self.assertEqual([r["v"] for r in rows], ["1", "2"])
        finally:
            writer.close()

    def test_書き戻しは共有のファイルへ_そのあと読むと写し直す(self):
        src = self.share / "梱包資材マスタ.sqlite3"
        _make(src).close()
        self.assertEqual([r["v"] for r in material_db.read_table(src, "t")], ["A"])
        with material_db.connect(src) as conn:
            conn.execute("UPDATE t SET v = 'Z'")
        conn = sqlite3.connect(str(src))
        self.assertEqual(conn.execute("SELECT v FROM t").fetchone()[0], "Z")
        conn.close()
        self.assertEqual([r["v"] for r in material_db.read_table(src, "t")], ["Z"])


if __name__ == "__main__":
    unittest.main()
