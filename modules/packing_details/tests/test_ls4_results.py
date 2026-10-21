"""LS4 の仕掛ファイルの2つの形式(統合 1.2.29・現場の指示。仕様は docs/仕掛ファイルの仕様.md)

    名前が LS4LOT     … 当工程設計_縦割数 / 当工程設計_横割数(いままでどおり)
    名前が LS4LOT 以外 … BOX設計_縦割数 / BOX設計_横割数。BOX実績_作業開始_日付 が**いまの年月**の行だけ
"""
from __future__ import annotations

import tempfile
import unittest
from datetime import date
from pathlib import Path

from common import ls4_source
from modules.packing_details.meisai import config, data_sync, user_settings

from . import _db

RESULT_COLUMNS = ("ﾛｯﾄ番号", "BOX設計_横割数", "BOX設計_縦割数", "BOX実績_作業開始_日付",
                  "BOX実績_ｺｲﾙ内径", "BOX実績_ｺｲﾙ肉厚")


def _ym(months_ago: int = 0) -> str:
    today = date.today()
    y, m = today.year, today.month - months_ago
    while m < 1:
        y, m = y - 1, m + 12
    return f"{y}-{m:02d}"


def _row(lot, tate, yoko, started):
    return {"ﾛｯﾄ番号": lot, "BOX設計_縦割数": tate, "BOX設計_横割数": yoko,
            "BOX実績_作業開始_日付": started, "BOX実績_ｺｲﾙ内径": "508", "BOX実績_ｺｲﾙ肉厚": "200.5"}


def _warisu(conn) -> dict[str, tuple[int, int]]:
    return {r[0]: (r[1], r[2]) for r in conn.execute(
        'SELECT "ロット番号", "当工程設計_縦割数", "当工程設計_横割数" FROM "仕掛当工程"')}


class Ls4ResultsTest(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)
        self.conn = _db.case_db(self)
        root = Path(tempfile.mkdtemp())
        self.konpo, self.second = root / "konpo", root / "second"
        self.konpo.mkdir()
        self.second.mkdir()
        user_settings.save(config.KEY_LOT_DB_DIR, str(root / "lot"))
        user_settings.save(config.KEY_KONPO_DB_DIR, str(self.konpo))

    def import_konpo(self):
        return data_sync.import_all(self.conn, force=True, only=config.KONPO_DB_FILES)

    def test_名前がLS4LOT以外なら実績の形式_いまの年月の行だけ(self):
        now = _ym()
        _db.make_source(self.konpo / "実績_LS4.sqlite3", RESULT_COLUMNS, [
            _row("A000001", "2", "6", f"{now}-21 00:01:56"),
            _row("B000001", "3", "4", f"{_ym(1)}-28 23:59:00"),          # 先月は使わない
            _row("C000001", "1", "12", f"{now.replace('-', '/')}/02 08:00:00"),   # / 区切りも読む
        ])
        user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
        result = self.import_konpo()
        self.assertTrue(result.ok, result.errors)
        self.assertEqual(_warisu(self.conn), {"A000001": (2, 6), "C000001": (1, 12)})
        self.assertTrue(any("実績の形式" in n and now in n for n in result.notes), result.notes)

    def test_同じロットが何行もあれば作業開始がいちばん新しい行(self):
        now = _ym()
        _db.make_source(self.konpo / "実績_LS4.sqlite3", RESULT_COLUMNS, [
            _row("A000001", "2", "6", f"{now}-01 08:00:00"),
            _row("A000001", "2", "12", f"{now}-15 08:00:00"),
            _row("A000001", "1", "1", f"{now}-03 08:00:00"),
        ])
        user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
        self.import_konpo()
        self.assertEqual(_warisu(self.conn), {"A000001": (2, 12)})

    def test_今月の行が無ければ空にする_先月の割数を使わない(self):
        _db.make_source(self.konpo / "LS4LOT.sqlite3", _db.TOKOTEI_COLUMNS, [_db.tokotei_row("Z000001")])
        self.import_konpo()
        self.assertIn("Z000001", _warisu(self.conn))
        _db.make_source(self.konpo / "実績_LS4.sqlite3", RESULT_COLUMNS,
                        [_row("B000001", "3", "4", f"{_ym(1)}-28 23:59:00")])
        user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
        result = self.import_konpo()
        self.assertEqual(_warisu(self.conn), {})
        self.assertTrue(any("0ロット" in n for n in result.notes), result.notes)

    def test_実績の形式の列が無ければ取り込まない(self):
        _db.make_source(self.konpo / "実績_LS4.sqlite3", _db.TOKOTEI_COLUMNS, [_db.tokotei_row("A000001")])
        user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
        result = self.import_konpo()
        self.assertEqual(_warisu(self.conn), {})
        self.assertTrue(any("BOX設計_横割数" in e for e in result.errors), result.errors)
        self.assertIn("BOX実績_作業開始_日付",
                      data_sync.second_missing_columns(self.konpo / "実績_LS4.sqlite3", "仕掛当工程"))

    def test_名前がLS4LOTなら当工程設計の列(self):
        _db.make_source(self.konpo / "LS4LOT.sqlite3", _db.TOKOTEI_COLUMNS,
                        [_db.tokotei_row("A000001", tate="2", yoko="12")])
        self.import_konpo()
        self.assertEqual(_warisu(self.conn), {"A000001": (2, 12)})

    def test_第1パスがLS4LOT_第2パスが実績なら_第1パスに無いロットを実績から足す(self):
        _db.make_source(self.konpo / "LS4LOT.sqlite3", _db.TOKOTEI_COLUMNS, [_db.tokotei_row("A000001")])
        _db.make_source(self.second / "実績_LS4.sqlite3", RESULT_COLUMNS, [
            _row("A000001", "9", "9", f"{_ym()}-05 08:00:00"),
            _row("D000001", "2", "6", f"{_ym()}-05 08:00:00")])
        user_settings.save(config.KEY_LOT_DB_DIR2, str(self.second))
        user_settings.save(config.KEY_LOT_DB2_FILES["仕掛当工程"], "実績_LS4")
        self.import_konpo()
        self.assertEqual(_warisu(self.conn), {"A000001": (2, 12), "D000001": (2, 6)})


class Ls4SourceRulesTest(unittest.TestCase):
    def test_名前で形式を決める(self):
        for name in ("LS4LOT.sqlite3", "ls4lot.db", "LS4LOT"):
            self.assertFalse(ls4_source.is_results_name(name), name)
        for name in ("実績_LS4.sqlite3", "LS4LOT2.sqlite3", "当工程"):
            self.assertTrue(ls4_source.is_results_name(name), name)

    def test_年月の読み方(self):
        self.assertEqual(ls4_source.year_month("2026-10-21 00:01:56"), (2026, 10))
        self.assertEqual(ls4_source.year_month("2026/1/5"), (2026, 1))
        self.assertIsNone(ls4_source.year_month(""))
        self.assertIsNone(ls4_source.year_month("2026-13-01"))

    def test_検入数は枚本数_無ければ横割数と縦割数の積(self):
        self.assertEqual(ls4_source.coil_count({"BOX実績_枚本数": "18", "BOX設計_横割数": "12",
                                                "BOX設計_縦割数": "2"}), 18)
        self.assertEqual(ls4_source.coil_count({"BOX実績_枚本数": "", "BOX設計_横割数": "12",
                                                "BOX設計_縦割数": "2"}), 24)
        self.assertIsNone(ls4_source.coil_count({"BOX設計_横割数": "", "BOX設計_縦割数": "2"}))

    def test_外径は内径と2倍の肉厚_無ければNone(self):
        self.assertEqual(ls4_source.outer_diameter({"BOX実績_ｺｲﾙ内径": "508", "BOX実績_ｺｲﾙ肉厚": "200.5"}), 909.0)
        self.assertIsNone(ls4_source.outer_diameter({"BOX実績_ｺｲﾙ内径": "508", "BOX実績_ｺｲﾙ肉厚": ""}))
        self.assertIsNone(ls4_source.outer_diameter({"BOX実績_ｺｲﾙ内径": "0", "BOX実績_ｺｲﾙ肉厚": "10"}))


if __name__ == "__main__":
    unittest.main()
