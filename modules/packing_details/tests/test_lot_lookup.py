"""ロット番号から、仕掛(LS4LOT)の同じ行の検査番号・コイル外径を引く(統合 1.2.57・現場の指示)

ペナラベルの LOTNO 欄・資材計算のコイル外径が使う(`meisai/lot_lookup.py`)。
"""
from __future__ import annotations

import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from modules.packing_details.meisai import lot_lookup


def _file(path: Path, columns: tuple[str, ...], rows: list[tuple], table: str = "仕掛") -> Path:
    c = sqlite3.connect(path)
    cols = ", ".join(f'"{n}" TEXT' for n in columns)
    c.execute(f'CREATE TABLE "{table}" ({cols})')
    c.executemany(f'INSERT INTO "{table}" VALUES ({", ".join("?" * len(columns))})', rows)
    c.commit()
    c.close()
    return path


FULL = ("ﾛｯﾄ番号", "検査番号", "BOX実績_作業開始_日付", "BOX実績_ｺｲﾙ内径", "BOX実績_ｺｲﾙ肉厚")


class LotLookupTest(unittest.TestCase):
    def setUp(self):
        self.root = Path(tempfile.mkdtemp())
        lot_lookup.forget()
        self.addCleanup(lot_lookup.forget)

    def _use(self, *files: Path):
        p = mock.patch.object(lot_lookup, "sources", lambda: list(files))
        p.start()
        self.addCleanup(p.stop)

    def test_同じ行の検査番号と外径を引く(self):
        self._use(_file(self.root / "LS4LOT.sqlite3", FULL, [
            ("A000001", "N6X0001", "2026-10-01", "508", "200.5"),
            ("A000002", "N6X0002", "2026-10-01", "508", "251.5")]))
        r = lot_lookup.lookup(" a000002 ")
        self.assertTrue(r["found"])
        self.assertEqual(r["kensa"], "N6X0002")
        self.assertEqual(r["outer"], 1011.0)
        self.assertEqual(r["notes"], [])

    def test_同じロットが何行もあれば作業開始のいちばん新しい行(self):
        self._use(_file(self.root / "LS4LOT.sqlite3", FULL, [
            ("A000001", "N6X0009", "2026-10-05", "508", "300"),
            ("A000001", "N6X0001", "2026-09-01", "508", "200")]))
        r = lot_lookup.lookup("A000001")
        self.assertEqual((r["kensa"], r["outer"]), ("N6X0009", 1108.0))

    def test_無いロットは理由を返す(self):
        self._use(_file(self.root / "LS4LOT.sqlite3", FULL, [("A000001", "N", "", "508", "1")]))
        r = lot_lookup.lookup("Z999999")
        self.assertFalse(r["found"])
        self.assertIn("Z999999 がありません", r["reason"])
        self.assertIn("空", lot_lookup.lookup("")["reason"])

    def test_ファイルが無ければそう言う(self):
        self._use()
        self.assertIn("見つかりません", lot_lookup.lookup("A000001")["reason"])

    def test_第1パスに無い項目だけ第2パスで埋める(self):
        first = _file(self.root / "LS4LOT.sqlite3", ("ﾛｯﾄ番号", "検査番号"), [("A000001", "N6X0001")])
        second_dir = self.root / "second"
        second_dir.mkdir()
        second = _file(second_dir / "LS4LOT.sqlite3", FULL, [("A000001", "OTHER", "", "508", "100")])
        self._use(first, second)
        r = lot_lookup.lookup("A000001")
        self.assertEqual(r["kensa"], "N6X0001", "第1パスの検査番号を残す")
        self.assertEqual(r["outer"], 708.0, "外径は第2パスから")
        self.assertEqual(r["file"], str(first))

    def test_検査番号の列が無ければ列を添えて知らせる(self):
        self._use(_file(self.root / "LS4LOT.sqlite3", ("ﾛｯﾄ番号", "品名", "BOX実績_ｺｲﾙ内径", "BOX実績_ｺｲﾙ肉厚"),
                        [("A000001", "X", "508", "100")]))
        r = lot_lookup.lookup("A000001")
        self.assertTrue(r["found"])
        self.assertEqual(r["kensa"], "")
        self.assertEqual(r["outer"], 708.0)
        self.assertTrue(any("検査番号の列が見つかりません" in n and "品名" in n for n in r["notes"]))

    def test_内径と肉厚が無ければ外径を出せないと知らせる(self):
        self._use(_file(self.root / "LS4LOT.sqlite3", ("ﾛｯﾄ番号", "検査番号"), [("A000001", "N6X0001")]))
        r = lot_lookup.lookup("A000001")
        self.assertEqual((r["kensa"], r["outer"]), ("N6X0001", None))
        self.assertTrue(any("コイル外径を出せません" in n for n in r["notes"]))

    def test_検査番号の列は名前で探す(self):
        self.assertEqual(lot_lookup.kensa_column(["ﾛｯﾄ番号", "検査番号"]), "検査番号")
        self.assertEqual(lot_lookup.kensa_column(["ﾛｯﾄ番号", "新検査番号"]), "新検査番号")
        self.assertEqual(lot_lookup.kensa_column(["ﾛｯﾄ番号", "新検番"]), "新検番")
        self.assertIsNone(lot_lookup.kensa_column(["ﾛｯﾄ番号", "品名"]))

    def test_ﾛｯﾄ番号の列を持つ表を選ぶ(self):
        path = self.root / "LS4LOT.sqlite3"
        c = sqlite3.connect(path)
        c.execute('CREATE TABLE "別表" ("x" TEXT)')
        c.execute('CREATE TABLE "当工程" ("ﾛｯﾄ番号" TEXT, "検査番号" TEXT)')
        c.execute('INSERT INTO "当工程" VALUES (?, ?)', ("A000001", "N1"))
        c.commit()
        c.close()
        self._use(path)
        self.assertEqual(lot_lookup.lookup("A000001")["kensa"], "N1")


if __name__ == "__main__":
    unittest.main()
