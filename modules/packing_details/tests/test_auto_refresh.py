"""15 分おきに参照パスと手元の写しを比べ、変わっていれば写し直す(統合 1.2.77・現場の指示)

現場で起きたこと: 資材計算・ペナラベルは LOTNO を引けるのに、梱包明細だけ「該当するﾛｯﾄ情報が見つかりません」。
梱包明細は起動したときに取り込んだ写しを引き、そのあと台帳に載ったロットは「取り込み」を押すまで引けなかった
(上の帯に「新しい台帳があります」と出るだけ)。
"""
from __future__ import annotations

import os
import shutil
import tempfile
import threading
import unittest
from pathlib import Path
from unittest import mock

from common import auto_refresh
from modules.packing_details.meisai import config, data_sync, db, lot_lookup, lot_repo

from . import _db


def _lot_source(path: Path, lots) -> Path:
    return _db.make_source(path, _db.LOT_COLUMNS, [_db.lot_row(n) for n in lots])


class DetailsRefreshTest(unittest.TestCase):
    def setUp(self):
        self.dir = Path(tempfile.mkdtemp(prefix="auto_refresh_"))
        self.addCleanup(shutil.rmtree, self.dir, True)
        self.conn, db_path = _db.file_db(self)
        patcher = mock.patch.object(config, "DB_PATH", db_path)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.path = _lot_source(self.dir / "SIKALOT.sqlite3", ["AAA0001"])
        for name, value in (("find_sources", lambda *a, **k: {"仕掛ロット": self.path}),
                            ("find_second_lots", lambda: {})):
            p = mock.patch.object(data_sync, name, value)
            p.start()
            self.addCleanup(p.stop)

    def _add_lot(self):
        _lot_source(Path(str(self.path) + ".new"), ["AAA0001", "BBB0002"])
        os.replace(str(self.path) + ".new", self.path)

    def test_あとから台帳に載ったロットも見回りのあとは引ける(self):
        data_sync.import_all(self.conn)
        self._add_lot()
        self.assertFalse(lot_repo.search(self.conn, "BBB0002").found, "見回りの前は前の写しのまま")
        self.assertEqual(auto_refresh.refresh_details(), ["仕掛ロット"])
        self.assertTrue(lot_repo.search(self.conn, "BBB0002").found)
        self.assertEqual(data_sync.newer_sources(self.conn), [], "「新しい台帳があります」は消える")

    def test_変わっていなければ取り込まない(self):
        data_sync.import_all(self.conn)
        with mock.patch.object(data_sync, "import_table") as spy:
            self.assertEqual(auto_refresh.refresh_details(), [])
        spy.assert_not_called()

    def test_取り込み中なら見送る(self):
        data_sync.import_all(self.conn)
        self._add_lot()
        with data_sync._import_lock:
            self.assertEqual(auto_refresh.refresh_details(), [])
        self.assertFalse(lot_repo.search(self.conn, "BBB0002").found)

    def test_取り込みは同時に1つ(self):
        """ボタンの取り込みと見回りが重ならない(片方が終わるまで待つ)。"""
        data_sync.import_all(self.conn)
        entered = threading.Event()
        release = threading.Event()
        real = data_sync._import_all

        def slow(*a, **k):
            entered.set()
            release.wait(5)
            return real(*a, **k)

        with mock.patch.object(data_sync, "_import_all", slow):
            def other():
                with db.connect() as conn:                 # 接続はスレッドごと
                    data_sync.import_all(conn, force=True)

            t = threading.Thread(target=other)
            t.start()
            self.assertTrue(entered.wait(5))
            self.assertTrue(data_sync.is_importing())
            release.set()
            t.join(5)
        self.assertFalse(data_sync.is_importing())


class RunOnceTest(unittest.TestCase):
    def test_1つが失敗してもほかは続ける(self):
        calls = []

        def bad():
            raise OSError("届かない")

        steps = (("A", bad), ("B", lambda: calls.append("B") or 1))
        with mock.patch.object(auto_refresh, "STEPS", steps):
            out = auto_refresh.run_once()
        self.assertEqual(calls, ["B"])
        self.assertIn("失敗", str(out["A"]))
        self.assertEqual(out["B"], 1)

    def test_間隔は15分(self):
        self.assertEqual(auto_refresh.INTERVAL_SEC, 15 * 60)


class PenaCopyTest(unittest.TestCase):
    def test_ロット引きの写しを取り直す(self):
        d = Path(tempfile.mkdtemp(prefix="auto_refresh_pena_"))
        self.addCleanup(shutil.rmtree, d, True)
        path = _lot_source(d / "SIKALOT.sqlite3", ["AAA0001"])
        with mock.patch.object(lot_lookup, "sources", lambda table=None: [path]), \
                mock.patch.object(lot_lookup.source_copy, "local_copy") as copy:
            self.assertEqual(lot_lookup.refresh_copies(), 1)
        copy.assert_called_once_with(path)


if __name__ == "__main__":
    unittest.main()
