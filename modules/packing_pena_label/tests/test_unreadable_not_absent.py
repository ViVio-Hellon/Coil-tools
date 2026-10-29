# -*- coding: utf-8 -*-
"""「読めない」を「無い」と取り違えない(統合 1.2.61)

別のツールで、共有ファイルが読めなかったのを「実績の表が無い」と取り違え、手元の実績を送り直して
重複データを大量に作った事案があった。ペナラベルで同じ形になりうる所を確かめる:

- この端末の設定(local.json)が**あるのに読めない**とき、空とみなして保存・配布設定の読み込みをしない
- 足したサイズ(共有の資材マスタの表)が読めないとき、前に読めた行のままにして「無い」と覚えない
- サイズ構成マスタ(JSON)・型番マスタ(DB)が読めないとき、同梱の定義へ黙って戻さない
- ラベル台紙一覧が読めないとき、「無い」と覚えない(次にまた読む)
- 資材マスタへ書くときは、無いファイルを作らない・確かめから書くまでを1つのまとまりにする
"""
import json
import os
import shutil
import sqlite3
import tempfile
import unittest
from pathlib import Path
from unittest import mock

from modules.packing_pena_label.app import config as C
from modules.packing_pena_label.app.config import Config
from modules.packing_pena_label.app.services import all_size as AS
from modules.packing_pena_label.app.services import distribution as D
from modules.packing_pena_label.app.services import master_admin as MA
from modules.packing_pena_label.app.services import settings as S
from modules.packing_pena_label.app.services import size_extra as SX
from modules.packing_pena_label.app.services.label_sheet_master import LabelSheetMaster
from modules.packing_pena_label.app.services.size_master import SizeMaster

from .test_distribution import DistributionTestCase
from .test_master_admin import make_master

BROKEN = "{ 書きかけ"


def _break(path):
    """sqlite として読めないファイルにする(使用中・壊れていると同じく「読めない」)。"""
    with open(path, "wb") as f:
        f.write(b"this is not a database" * 64)


class LocalSettingsTest(DistributionTestCase):
    """この端末の設定(local.json)。"""

    def _corrupt(self, cfg):
        os.makedirs(cfg.local_config_dir, exist_ok=True)
        Path(cfg.local_config_path).write_text(BROKEN, encoding="utf-8")

    def test_無いときは空_あるのに読めないときは例外(self):
        self.assertEqual(C.load_local_overrides(self.dest, strict=True), {})
        self._corrupt(self.dest)
        self.assertEqual(C.load_local_overrides(self.dest), {}, "表示用は空で続ける")
        with self.assertRaises(C.LocalConfigUnreadable):
            C.load_local_overrides(self.dest, strict=True)

    def test_読めないときは保存しない(self):
        self._corrupt(self.dest)
        with self.assertRaises(C.LocalConfigUnreadable):
            C.save_local_overrides({"label_scale_pct": 100.4}, self.dest)
        self.assertEqual(Path(self.dest.local_config_path).read_text(encoding="utf-8"), BROKEN)

    def test_読めないときは設定画面の保存を断る(self):
        self._corrupt(self.dest)
        r = S.save({"label_scale_pct": "100.4"}, self.dest)
        self.assertFalse(r["ok"])
        self.assertIn("読めない", r["message"])
        self.assertEqual(Path(self.dest.local_config_path).read_text(encoding="utf-8"), BROKEN)

    def test_読めないときは配布設定で上書きしない(self):
        """以前は読めないと「この端末にはまだ無い」とみなし、配布の値で全部上書きしていた。"""
        self.configure_source()
        self.assertTrue(self.export().ok)
        self._corrupt(self.dest)
        r = D.apply_on_start(self.dest)
        self.assertFalse(r.ok)
        self.assertEqual(r.applied, [])
        self.assertEqual(Path(self.dest.local_config_path).read_text(encoding="utf-8"), BROKEN)
        # 読めるようになれば読み込む(無いものだけ)
        Path(self.dest.local_config_path).write_text(json.dumps({"label_scale_pct": 99.0}),
                                                     encoding="utf-8")
        r = D.apply_on_start(self.dest)
        self.assertTrue(r.ok, r.message)
        self.assertEqual(C.load_local_overrides(self.dest)["label_scale_pct"], 99.0, "この端末の値は残す")

    def test_読めないときは配布設定を書き出さない(self):
        self.configure_source()
        self._corrupt(self.source)
        r = self.export()
        self.assertFalse(r.ok)
        self.assertIn("読めない", r.message)


class SizeExtraTest(unittest.TestCase):
    """足したサイズ(共有の資材マスタの表「ペナラベル追加サイズ」)。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.db = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        make_master(self.db)
        c = sqlite3.connect(self.db)
        c.execute('CREATE TABLE "%s" ("管理番号" INTEGER, "サイズ名" TEXT, "巾" REAL, "型番" TEXT,'
                  ' "記号" TEXT, "備考" TEXT, "登録日時" TEXT, "削除日時" TEXT)' % SX.TABLE_NAME)
        c.execute('INSERT INTO "%s" VALUES (15, "1.2mm×45.0mm", 45, "BJB7699900QR", "", "", "", "")'
                  % SX.TABLE_NAME)
        c.commit()
        c.close()
        self.good = open(self.db, "rb").read()

    def test_読めないときは前の行のまま_次にまた読む(self):
        sx = SX.SizeExtra(refresh_sec=0)
        self.assertEqual([r["obIdx"] for r in sx.load(self.db)], [15])
        _break(self.db)
        rows = sx.load(self.db)
        self.assertEqual([r["obIdx"] for r in rows], [15], "読めないのを「無い」とみなさない")
        self.assertFalse(sx.ok)
        self.assertIn("読めません", sx.last_error)
        self.assertIsNone(sx._stamp, "印を覚えない(次にまた読む)")
        with open(self.db, "wb") as f:
            f.write(self.good)
        sx.load(self.db)
        self.assertTrue(sx.ok)

    def test_見えないときも前の行のまま(self):
        sx = SX.SizeExtra(refresh_sec=0)
        sx.load(self.db)
        moved = self.db + ".moved"
        os.replace(self.db, moved)
        self.assertEqual([r["obIdx"] for r in sx.load(self.db)], [15])
        self.assertFalse(sx.ok)

    def test_表が無いのは無い(self):
        c = sqlite3.connect(self.db)
        c.execute('DROP TABLE "%s"' % SX.TABLE_NAME)
        c.commit()
        c.close()
        sx = SX.SizeExtra(refresh_sec=0)
        self.assertEqual(sx.load(self.db), [])
        self.assertTrue(sx.ok)
        self.assertEqual(SX.SizeExtra(refresh_sec=0).load(""), [], "資材マスタを設定していない")

    def test_読めないときに例外を選べる(self):
        _break(self.db)
        self.assertEqual(SX.read_rows(self.db), [])
        with self.assertRaises(sqlite3.Error):
            SX.read_rows(self.db, strict=True)


class MasterFallbackTest(unittest.TestCase):
    """サイズ構成マスタ・型番マスタ・ラベル台紙一覧。同梱の定義へ黙って戻さない。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_サイズ構成マスタが見えなければ今の内容のまま(self):
        m = SizeMaster()
        before = [(c.ob_idx, c.kataban) for c in m.all()]
        self.assertFalse(m.reload(os.path.join(self.tmp, "無い.json")))
        self.assertEqual([(c.ob_idx, c.kataban) for c in m.all()], before)
        self.assertTrue(m.reload(None), "設定していないときは同梱の定義")

    def test_型番マスタのDBが読めなければ同梱のCSVへ戻さない(self):
        db = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        c = sqlite3.connect(db)
        c.execute('CREATE TABLE "%s" ("サイズ名" TEXT, "記号" TEXT, "型番" TEXT, "備考" TEXT)' % AS.KATABAN_TABLE)
        c.execute('INSERT INTO "%s" VALUES ("DBのサイズ", "D", "BJB-DB", "")' % AS.KATABAN_TABLE)
        c.commit()
        c.close()
        csv_path = os.path.join(self.tmp, "型番.csv")
        Path(csv_path).write_text("サイズ名,記号,型番,備考\nCSVのサイズ,C,BJB-CSV,\n", encoding="utf-8")
        svc = AS.AllSizeService(mock.MagicMock(), csv_path, db)
        self.assertEqual(svc.master.names(), ["DBのサイズ"])
        _break(db)
        self.assertFalse(svc.reload_master(csv_path, db))
        self.assertEqual(svc.master.names(), ["DBのサイズ"], "CSV に戻さない")
        self.assertIsNone(svc._stamp, "次にまた読む")
        with self.assertRaises(sqlite3.Error):
            AS.KatabanMaster.load(csv_path, db, strict=True)
        self.assertEqual(AS.KatabanMaster.load(csv_path, db).names(), ["CSVのサイズ"], "起動時は CSV で続ける")

    def test_ラベル台紙一覧は表が無いときだけ無いと覚える(self):
        db = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        _break(db)
        lm = LabelSheetMaster(refresh_sec=0)
        self.assertEqual(lm.load(db), {})
        self.assertIsNone(lm._stamp, "読めなかっただけなら覚えない")
        os.remove(db)
        make_master(db)
        self.assertEqual(lm.load(db), {})
        self.assertIsNotNone(lm._stamp, "表が無いと確かめられたので覚える")


class WriteConnectionTest(unittest.TestCase):
    """資材マスタへ書くとき。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def test_無いファイルは作らない(self):
        path = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        with self.assertRaises(sqlite3.Error):
            MA._connect(path, write=True)
        self.assertFalse(os.path.exists(path), "差し替え中などで消えていたマスタの場所に空のファイルを作らない")

    def test_開いたらすぐ書くまとまりを始める(self):
        path = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        make_master(path)
        conn = MA._connect(path, write=True)
        try:
            self.assertTrue(conn.in_transaction)
            other = sqlite3.connect(path, timeout=0.1)
            with self.assertRaises(sqlite3.OperationalError, msg="2台目は確かめの前で待つ(重ならない)"):
                other.execute("BEGIN IMMEDIATE")
            other.close()
        finally:
            conn.rollback()
            conn.close()

    def test_サイズを足すのは確かめから書くまで1つのまとまり(self):
        """2台が同時に同じサイズを足しても、確かめと番号の割り当てが重ならない。"""
        from modules.packing_pena_label.app.repositories.material_repo import MaterialRepository
        path = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        make_master(path)
        repo = MaterialRepository(accdb_dir="", csv_path="", prefer_access=False, db_path=path, refresh_sec=0)
        holder = sqlite3.connect(path, timeout=0.1)
        holder.execute("BEGIN IMMEDIATE")             # ほかの PC が書いている途中
        try:
            orig = sqlite3.connect
            # 待つ時間だけ短くする(本番は 15 秒待つ)
            with mock.patch.object(MA.sqlite3, "connect",
                                   lambda *a, **k: orig(*a, **dict(k, timeout=0.1))):
                with self.assertRaises(MA.Refused) as cm:
                    SX.add(repo, SizeMaster(), "1.2", "45", "BJB7699900QR")
            self.assertEqual(cm.exception.kind, MA.REFUSE_WRITE_FAILED)
        finally:
            holder.rollback()
            holder.close()
        self.assertEqual(SX.read_rows(path), [], "待てなければ書かない(確かめないまま書かない)")
        r = SX.add(repo, SizeMaster(), "1.2", "45", "BJB7699900QR")
        self.assertEqual(r["obIdx"], [15, 16])


if __name__ == "__main__":
    unittest.main()
