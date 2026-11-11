"""仕掛ロット(SIKALOT)に無いロットも、実績の形式のファイル(実績_LS4)から引ける(統合 1.2.78・現場の指摘)

現場で起きたこと: 資材計算・ペナラベルは LOTNO を引けるのに、梱包明細だけ「該当するﾛｯﾄ情報が見つかりません」
(取り込み直しても同じ)。LS4 を通り終えたコイルは SIKALOT から外れるが、仕掛引当(SIKAHIKI)と実績_LS4 には
残る(2026-10 の実ファイルで、実績_LS4 の 912 ロットのうち 899 ロットが SIKALOT に無かった)。
"""
from __future__ import annotations

import os
import tempfile
import unittest
from pathlib import Path

from modules.packing_details.meisai import config, data_sync, lot_repo, user_settings

from . import _db
from .test_ls4_results import _ym

LOT_COLS = ("ﾛｯﾄ番号", "用途ｺｰﾄﾞ", "用途名", "製造材質", "製造調質", "製造板厚", "製造板幅",
            "ｵｰﾀﾞｰ板厚", "ｵｰﾀﾞｰ板幅", "ｵｰﾀﾞｰ板丈", "設計_設備ｺｰｽ", "実績_設備ｺｰｽ")
RESULT_COLUMNS = LOT_COLS + ("BOX設計_横割数", "BOX設計_縦割数", "BOX実績_作業開始_日付")


def _result(lot, *, name="ｼﾞｸｳｹﾖｳ", length="0.0", started=None):
    return {"ﾛｯﾄ番号": lot, "用途ｺｰﾄﾞ": "R100", "用途名": name, "製造材質": "52S", "製造調質": "H34",
            "製造板厚": "1.300", "製造板幅": "279.80", "ｵｰﾀﾞｰ板厚": "1.3", "ｵｰﾀﾞｰ板幅": "280.0",
            "ｵｰﾀﾞｰ板丈": length, "設計_設備ｺｰｽ": "HOT L-1 LS4", "実績_設備ｺｰｽ": "HOT L-1",
            "BOX設計_縦割数": "3", "BOX設計_横割数": "4",
            "BOX実績_作業開始_日付": started or f"{_ym()}-05 08:00:00"}


class ResultsLotsTest(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)
        self.conn = _db.case_db(self)
        root = Path(tempfile.mkdtemp())
        self.lot, self.second = root / "lot", root / "second"
        self.lot.mkdir()
        self.second.mkdir()
        user_settings.save(config.KEY_LOT_DB_DIR, str(self.lot))
        user_settings.save(config.KEY_KONPO_DB_DIR, str(self.lot))
        user_settings.save(config.KEY_LOT_DB_DIR2, str(self.second))
        user_settings.save(config.KEY_LOT_DB2_FILES["仕掛当工程"], "実績_LS4")
        _db.make_source(self.lot / "SIKALOT.sqlite3", _db.LOT_COLUMNS, [_db.lot_row("A000001")])
        _db.make_source(self.lot / "LS4LOT.sqlite3", _db.TOKOTEI_COLUMNS, [_db.tokotei_row("A000001")])
        self.results = self.second / "実績_LS4.sqlite3"

    def make_results(self, *rows):
        tmp = self.results.with_name("tmp.sqlite3")
        _db.make_source(tmp, RESULT_COLUMNS, list(rows))
        os.replace(tmp, self.results)

    def test_SIKALOTに無いロットも実績から引ける(self):
        self.make_results(_result("L8101L0"))
        result = data_sync.import_all(self.conn)
        self.assertTrue(any("実績_LS4.sqlite3" in n and "1ロット" in n for n in result.notes), result.notes)
        found = lot_repo.search(self.conn, "L8101L0")
        self.assertTrue(found.found, found.message)
        self.assertEqual(found.lot.yoto_name, "ｼﾞｸｳｹﾖｳ")
        self.assertEqual((found.lot.seizou_thickness, found.lot.seizou_width), (1.3, 279.8))

    def test_SIKALOTにあればそちらが正(self):
        self.make_results(_result("A000001", name="実績のほう"))
        data_sync.import_all(self.conn)
        self.assertNotEqual(lot_repo.search(self.conn, "A000001").lot.yoto_name, "実績のほう")

    def test_オーダー板丈が0でなければ引かない_VBAと同じ条件(self):
        self.make_results(_result("L8061F0", length="533.0"))
        data_sync.import_all(self.conn)
        self.assertFalse(lot_repo.search(self.conn, "L8061F0").found)

    def test_月は絞らない_ロットごとに新しい行(self):
        self.make_results(_result("B000001", name="古い", started=f"{_ym(2)}-01 08:00:00"),
                          _result("B000001", name="新しい", started=f"{_ym(1)}-28 08:00:00"))
        data_sync.import_all(self.conn)
        self.assertEqual(lot_repo.search(self.conn, "B000001").lot.yoto_name, "新しい")

    def test_変わっていなければ読み直さない_変われば入れ替える(self):
        self.make_results(_result("L8101L0"))
        data_sync.import_all(self.conn)
        again = data_sync.import_all(self.conn)
        self.assertFalse(any("実績" in n for n in again.notes), again.notes)
        self.make_results(_result("L9999X0"))
        later = os.path.getmtime(self.results) + 5
        os.utime(self.results, (later, later))
        data_sync.import_all(self.conn)
        self.assertTrue(lot_repo.search(self.conn, "L9999X0").found)
        self.assertFalse(lot_repo.search(self.conn, "L8101L0").found)

    def test_読めなければ前の中身のまま(self):
        self.make_results(_result("L8101L0"))
        data_sync.import_all(self.conn)
        self.results.write_bytes(b"not a database")
        data_sync.import_all(self.conn, force=True)
        self.assertTrue(lot_repo.search(self.conn, "L8101L0").found, "読めないのを「無い」と取り違えて消さない")


if __name__ == "__main__":
    unittest.main()
