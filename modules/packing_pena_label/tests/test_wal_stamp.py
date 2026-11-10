# -*- coding: utf-8 -*-
"""共有の資材マスタが WAL のとき、書き込みを見落とさない(統合 1.2.69・構造点検)

WAL の形のファイルでは、書き込みはまず `-wal` に入り、本体の大きさ・更新日時は変わらない(ほかのプログラムが
開いたままのあいだは本体へ移らない)。「変わったか」を本体だけで見ると、直したのに前の中身のまま動く。
統合 1.2.49 で3機能の目印に `-wal` を入れたが、ペナラベルの「足したサイズ」と「ラベル台紙一覧」が漏れていた
(全サイズの型番 `all_size` は入っていた)。
"""
from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
import unittest

from modules.packing_pena_label.app.services import size_extra
from modules.packing_pena_label.app.services.label_sheet_master import LabelSheetMaster
from modules.packing_pena_label.app.services.size_extra import SizeExtra


class WalStampTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="ppl-wal-")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.db = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        conn = sqlite3.connect(self.db)
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute('CREATE TABLE "ラベル台紙一覧" (管理番号 INTEGER, シート名 TEXT, 型番 TEXT)')
        conn.execute('INSERT INTO "ラベル台紙一覧" VALUES (1, "台紙A", "BJB0000001")')
        conn.execute('CREATE TABLE "%s" ("管理番号" INTEGER, "サイズ名" TEXT, "巾" REAL, "型番" TEXT, '
                     '"記号" TEXT, "備考" TEXT, "登録日時" TEXT, "%s" TEXT)'
                     % (size_extra.TABLE_NAME, size_extra.DELETED_COLUMN))
        conn.execute('INSERT INTO "%s" VALUES (?, "1.0×50×Coil", 50, "BJB1", "a", "", "", "")'
                     % size_extra.TABLE_NAME, (size_extra.FIRST_OB,))
        conn.commit()
        conn.close()
        # ほかのプログラム(マスタ管理・Access など)が開いたまま。書いても本体へは移らない
        self.holder = sqlite3.connect(self.db)
        self.holder.execute("PRAGMA wal_autocheckpoint=0")
        self.holder.execute("SELECT 1").fetchone()
        self.addCleanup(self.holder.close)

    def write(self, query: str) -> None:
        before = os.stat(self.db)
        self.holder.execute(query)
        self.holder.commit()
        after = os.stat(self.db)
        # 前提: 本体の大きさ・更新日時は変わらず、-wal だけが変わる
        self.assertEqual((before.st_size, before.st_mtime_ns), (after.st_size, after.st_mtime_ns))
        self.assertTrue(os.path.exists(self.db + "-wal"))

    def test_ラベル台紙一覧の書き込みを見落とさない(self):
        master = LabelSheetMaster(refresh_sec=0)
        self.assertEqual(master.load(self.db)[1], ("台紙A", "BJB0000001"))
        self.write('UPDATE "ラベル台紙一覧" SET 型番 = "BJB0000002" WHERE 管理番号 = 1')
        self.assertEqual(master.load(self.db)[1], ("台紙A", "BJB0000002"), "前の型番のまま")

    def test_足したサイズの書き込みを見落とさない(self):
        extra = SizeExtra(refresh_sec=0)
        first = [r["obIdx"] for r in extra.load(self.db)]
        self.assertEqual(first, [size_extra.FIRST_OB])
        self.write('INSERT INTO "%s" VALUES (%d, "1.2×60×Coil", 60, "BJB2", "b", "", "", "")'
                   % (size_extra.TABLE_NAME, size_extra.FIRST_OB + 2))
        self.assertEqual([r["obIdx"] for r in extra.load(self.db)],
                         [size_extra.FIRST_OB, size_extra.FIRST_OB + 2], "足したサイズが出ない")

    def test_変わっていなければ読み直さない(self):
        master = LabelSheetMaster(refresh_sec=0)
        master.load(self.db)
        master.load(self.db)                 # 初めて開くと空の -wal ができるので、1度だけ読み直す
        stamp = master._stamp
        master.load(self.db)
        self.assertEqual(master._stamp, stamp)


if __name__ == "__main__":
    unittest.main()
