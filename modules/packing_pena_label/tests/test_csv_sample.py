# -*- coding: utf-8 -*-
"""資材マスタ CSV・型番マスタ CSV の「手本を出す」と、読む順 DB > CSV（統合 1.2.8・現場の指定）

- 手本は欄に入れたパスへ出す。すでにあるファイルは上書きしない。フォルダーを入れたら同梱と同じ名前
- 手本は Excel でそのまま開ける UTF-8（BOM）・CRLF。Excel で保存し直した Shift_JIS も読める
- 資材マスタ: SQLite → (旧 Access) → CSV（前から）
- 型番マスタ: 梱包資材マスタ（SQLite）の表「型番」→ CSV（表が無い・空なら CSV）
"""
import os
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path

from modules.packing_pena_label.app.config import Config
from modules.packing_pena_label.app.repositories import csv_text
from modules.packing_pena_label.app.repositories.material_repo import MaterialRepository
from modules.packing_pena_label.app.services import csv_sample
from modules.packing_pena_label.app.services.all_size import KatabanMaster

from .test_http import HttpTestCase


class SampleTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.cfg = Config()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_出す_中身は同梱と同じ見出しと例(self):
        for key, name, head in (("material_csv_file", "資材.csv", "管理番号,梱包資材名,単位質量,係数"),
                                ("kataban_csv_file", "型番手本.csv", "サイズ名,記号,型番,備考")):
            with self.subTest(key=key):
                path = os.path.join(self.tmp, name)
                r = csv_sample.write_sample(self.cfg, key, '"%s"' % path)   # パスのコピーの " も取る
                self.assertTrue(r["ok"], r)
                raw = Path(path).read_bytes()
                self.assertTrue(raw.startswith(b"\xef\xbb\xbf"), "BOM 付き(Excel で化けない)")
                self.assertNotIn(b"\n", raw.replace(b"\r\n", b""), "改行は CRLF")
                self.assertEqual(csv_text.read_rows(path)[0], head.split(","))
                self.assertEqual(csv_text.read_rows(path),
                                 csv_text.read_rows(csv_sample.SAMPLES[key].bundled(self.cfg)))
                self.assertGreater(r["rows"], 0)

    def test_拡張子が無ければcsvを付ける(self):
        """「ペナ型番」と入れたら、拡張子の無いファイルではなく ペナ型番.csv を出す(現場の指摘)。"""
        r = csv_sample.write_sample(self.cfg, "kataban_csv_file", os.path.join(self.tmp, "ペナ型番"))
        self.assertTrue(r["ok"], r)
        self.assertEqual(r["path"], os.path.join(self.tmp, "ペナ型番.csv"))
        self.assertTrue(r["addedExt"])
        self.assertEqual(os.listdir(self.tmp), ["ペナ型番.csv"])
        r = csv_sample.write_sample(self.cfg, "kataban_csv_file", os.path.join(self.tmp, "型番.CSV"))
        self.assertEqual(r["path"], os.path.join(self.tmp, "型番.CSV"), "大文字の .CSV はそのまま")
        self.assertFalse(r["addedExt"])

    def test_あるファイルは上書きしない(self):
        path = os.path.join(self.tmp, "資材重量.csv")
        with open(path, "w", encoding="utf-8") as f:
            f.write("使っているマスタ")
        r = csv_sample.write_sample(self.cfg, "material_csv_file", path)
        self.assertFalse(r["ok"])
        self.assertIn("上書きしません", r["message"])
        self.assertEqual(Path(path).read_text(encoding="utf-8"), "使っているマスタ")

    def test_フォルダーなら同梱と同じ名前で出す(self):
        r = csv_sample.write_sample(self.cfg, "kataban_csv_file", self.tmp)
        self.assertTrue(r["ok"], r)
        self.assertEqual(r["path"], os.path.join(self.tmp, "型番.csv"))

    def test_空_フォルダーが無い_ほかの欄は断る(self):
        self.assertFalse(csv_sample.write_sample(self.cfg, "kataban_csv_file", " ")["ok"])
        r = csv_sample.write_sample(self.cfg, "kataban_csv_file",
                                    os.path.join(self.tmp, "無い", "型番.csv"))
        self.assertFalse(r["ok"])
        self.assertIn("フォルダーがありません", r["message"])
        self.assertFalse(os.path.exists(os.path.join(self.tmp, "無い")), "フォルダーを作らない")
        self.assertFalse(csv_sample.write_sample(self.cfg, "size_master_file", self.tmp)["ok"])

    def test_Excelで保存したShift_JISも読む(self):
        path = os.path.join(self.tmp, "型番.csv")
        with open(path, "w", encoding="cp932", newline="") as f:
            f.write("サイズ名,記号,型番,備考\r\n1.0×33×Coil,a,BJB7603900QR,\r\n")
        m = KatabanMaster.load(path)
        self.assertEqual((m.kind, m.names()), ("csv", ["1.0×33×Coil"]))
        mat = os.path.join(self.tmp, "資材重量.csv")
        with open(mat, "w", encoding="cp932", newline="") as f:
            f.write("管理番号,梱包資材名,単位質量,係数\r\n1,テスラピン,0.35,1\r\n")
        table = MaterialRepository("", mat, prefer_access=False).load(use_cache=False)
        self.assertTrue(table.has("テスラピン"))


def _db(path, kataban_rows=None):
    conn = sqlite3.connect(path)
    conn.execute('CREATE TABLE "資材重量" ("管理番号" TEXT, "梱包資材名" TEXT, "単位質量" TEXT, "係数" TEXT)')
    conn.execute('INSERT INTO "資材重量" VALUES (\'1\', \'テスラピン\', \'0.5\', \'1\')')
    if kataban_rows is not None:
        conn.execute('CREATE TABLE "型番" ("サイズ名" TEXT, "記号" TEXT, "型番" TEXT, "備考" TEXT)')
        conn.executemany('INSERT INTO "型番" VALUES (?, ?, ?, ?)', kataban_rows)
    conn.commit()
    conn.close()
    return path


class ReadOrderTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.csv = Config().default_kataban_csv_path

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_型番はDBに表があればDB(self):
        db = _db(os.path.join(self.tmp, "m.sqlite3"), [("DB×1×Coil", "z", "DB0001", "")])
        m = KatabanMaster.load(self.csv, db)
        self.assertEqual((m.kind, m.source, m.names()), ("sqlite", db, ["DB×1×Coil"]))
        self.assertEqual(m.by_name("DB×1×Coil")["kataban"], "DB0001")

    def test_型番はDBに表が無い_空なら_CSV(self):
        for rows in (None, []):
            with self.subTest(rows=rows):
                db = _db(os.path.join(self.tmp, "m%s.sqlite3" % (rows is None)), rows)
                m = KatabanMaster.load(self.csv, db)
                self.assertEqual((m.kind, m.source), ("csv", self.csv))
                self.assertGreater(m.count(), 0)
        self.assertEqual(KatabanMaster.load(self.csv, os.path.join(self.tmp, "無い.sqlite3")).kind, "csv")

    def test_資材マスタもDBが先(self):
        db = _db(os.path.join(self.tmp, "m.sqlite3"))
        repo = MaterialRepository("", Config().default_material_csv_path,
                                  prefer_access=False, db_path=db)
        repo.load(use_cache=False)
        self.assertEqual(repo.last_source, "sqlite")
        os.remove(db)
        repo.load(use_cache=False)
        self.assertEqual(repo.last_source, "csv")


class ScreenTest(HttpTestCase):
    @classmethod
    def setUpClass(cls):
        # HttpTestCase は Config.local_dir を差し替える。あとの試験へ持ち越さないよう戻す
        cls._local_dir = Config.__dict__["local_dir"]
        super().setUpClass()

    @classmethod
    def tearDownClass(cls):
        super().tearDownClass()
        Config.local_dir = cls._local_dir

    def test_2つの欄に手本を出すと読む順(self):
        _, html = self.get("/settings")
        for key in ("material_csv_file", "kataban_csv_file"):
            self.assertIn('data-sample="%s"' % key, html)
            self.assertIn('id="sample_%s"' % key, html)
        self.assertEqual(html.count("data-sample="), 2, "ほかの欄には出さない")
        self.assertIn("読む順: 梱包資材マスタ（SQLite）の表「型番」→ この CSV", html)
        self.assertIn("いま読んでいる先", html)

    def test_APIで出す_合言葉は要らない_設定は変えない(self):
        d = tempfile.mkdtemp()
        try:
            path = os.path.join(d, "型番.csv")
            _, j = self.post("/api/settings/csv-sample", {"key": "kataban_csv_file", "path": path})
            self.assertTrue(j["ok"], j)
            self.assertTrue(os.path.exists(path))
            self.assertEqual(self.ctx.cfg.kataban_csv_file, "", "欄の値は保存しない")
            _, j = self.post("/api/settings/csv-sample", {"key": "kataban_csv_file", "path": path})
            self.assertFalse(j["ok"])
        finally:
            shutil.rmtree(d, ignore_errors=True)


if __name__ == "__main__":
    unittest.main()
