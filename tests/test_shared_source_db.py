"""取り込み元の読み書き(source_db)は common に1つ(統合 1.2.71・構造点検の改善案2)

梱包明細と資材計算にほぼ同じ写しが1本ずつあり、直しが片方にしか入らないことが続いた(1.2.61 の `mode=rw` は
資材計算だけ、`find_strict` は梱包明細だけ、1.2.69 でやっと揃えた)。2つの機能の `source_db` が common の
1つを指していることを確かめる(写しに戻すと、この試験が落ちる)。
"""
from __future__ import annotations

import unittest

import common.source_db as shared
from modules.packing_details.meisai import source_db as details
from modules.packing_material_calculation.coil_tool import source_db as material


class SharedSourceDbTest(unittest.TestCase):
    def test_2つの機能が同じ1つを使う(self):
        self.assertIs(details, shared)
        self.assertIs(material, shared)

    def test_どちらの機能の呼び名も使える(self):
        for name in ("find_strict", "folder_reachable", "file_stamp", "open_read_only",
                     "list_tables_strict", "columns_strict", "table_names", "table_columns"):
            self.assertTrue(callable(getattr(shared, name)), name)
        self.assertTrue(callable(shared.SourceConnection.execute_all))
        self.assertIs(shared.table_names, shared.list_tables_strict)
        self.assertIs(shared.table_columns, shared.columns_strict)

    def test_写しの中身を持たない(self):
        from pathlib import Path
        root = Path(__file__).resolve().parent.parent
        for rel in ("modules/packing_details/meisai/source_db.py",
                    "modules/packing_material_calculation/coil_tool/source_db.py"):
            self.assertLess(len((root / rel).read_text(encoding="utf-8").splitlines()), 20, rel)


if __name__ == "__main__":
    unittest.main()
