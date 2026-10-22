# -*- coding: utf-8 -*-
"""足したサイズを消す・どの PC にも出す・見出しが見える(統合 1.2.38・現場の指摘)

- マスタ管理の「ペナラベル追加サイズ」から消せる(論理削除。行は残して削除日時を入れる)
- 消したサイズの番号は新しいサイズで使い回さない(各 PC の手元の台紙の中身が新しいサイズに出ないように)
- 消したサイズと同じ寸法は、もう一度足せる(番号は新しくなる)
- 足したサイズは、どの PC の指定サイズの画面にも最初から出る(この PC で外したものだけ出ない)
- 足したサイズの見出しにも色の帯を付ける(暗い配色で文字が見えなかった)
"""
import json
import os
import shutil
import sqlite3
import tempfile
import threading
import unittest
import urllib.error
import urllib.request

from modules.packing_pena_label import server
from modules.packing_pena_label.app.config import Config
from modules.packing_pena_label.app.repositories.material_repo import MaterialRepository
from modules.packing_pena_label.app.services import master_admin as MA
from modules.packing_pena_label.app.services import settings as S
from modules.packing_pena_label.app.services import size_extra as SX
from modules.packing_pena_label.app.services.size_master import SizeMaster

from ._material import business_csv
from .test_master_admin import free_port, make_master


def _rowid(db, name):
    with sqlite3.connect(db) as c:
        return c.execute('SELECT rowid FROM "%s" WHERE "サイズ名"=?' % SX.TABLE_NAME,
                         (name,)).fetchone()[0]


class SoftDeleteTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.db = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        make_master(self.db)
        self.repo = MaterialRepository(accdb_dir="", csv_path="", prefer_access=False,
                                       db_path=self.db, refresh_sec=0)
        self.master = SizeMaster()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_消すと読まない_番号は使い回さない_同じ寸法は足し直せる(self):
        SX.add(self.repo, self.master, "1.2", "45", "BJB7699900QR")
        self.master.apply_extra(SX.read_rows(self.db))
        self.assertTrue(MA.can_delete(SX.TABLE_NAME))
        r = MA.delete_row(self.repo, SX.TABLE_NAME, _rowid(self.db, "1.2mm×45.0mm"))
        self.assertIn("削除日時", r["message"])
        self.assertEqual(SX.read_rows(self.db), [], "消したサイズは読まない")
        with sqlite3.connect(self.db) as c:
            left = c.execute('SELECT "管理番号", "削除日時" FROM "%s"' % SX.TABLE_NAME).fetchall()
        self.assertEqual(len(left), 1, "行は残す")
        self.assertTrue(left[0][1])
        self.master.apply_extra(SX.read_rows(self.db))
        again = SX.add(self.repo, self.master, "1.2", "45", "BJB7699901QR")
        self.assertEqual(again["obIdx"], [17, 18], "消した 15 は使い回さない")
        with self.assertRaises(MA.Refused):
            MA.delete_row(self.repo, SX.TABLE_NAME, _rowid(self.db, "1.2mm×45.0mm"))

    def test_2度消しは断る_ほかの決まった表は消せないまま(self):
        SX.add(self.repo, self.master, "1.2", "45", "BJB7699900QR")
        rid = _rowid(self.db, "1.2mm×45.0mm")
        MA.delete_row(self.repo, SX.TABLE_NAME, rid)
        with self.assertRaises(MA.Refused) as cm:
            MA.delete_row(self.repo, SX.TABLE_NAME, rid)
        self.assertIn("もう消してあります", cm.exception.message)
        self.assertFalse(MA.can_delete("ラベル台紙一覧"))
        self.assertTrue(MA.can_delete("資材重量"))

    def test_前の版で作った表_削除日時の列が無い_も消せる(self):
        with sqlite3.connect(self.db) as c:
            c.execute('CREATE TABLE "%s" ("管理番号" INTEGER, "サイズ名" TEXT, "巾" REAL, '
                      '"型番" TEXT, "記号" TEXT, "備考" TEXT, "登録日時" TEXT)' % SX.TABLE_NAME)
            c.execute('INSERT INTO "%s" VALUES (15, "1.2mm×45.0mm", 45, "BJB7699900QR", "", "", "")'
                      % SX.TABLE_NAME)
        self.assertEqual(len(SX.read_rows(self.db)), 1)
        MA.delete_row(self.repo, SX.TABLE_NAME, _rowid(self.db, "1.2mm×45.0mm"))
        self.assertEqual(SX.read_rows(self.db), [])
        page = MA.page(self.repo, SX.TABLE_NAME)
        self.assertTrue(page["canDelete"])
        self.assertIn("削除日時", page["columns"])


class SizeDeleteHttpTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._local_dir = Config.__dict__["local_dir"]
        cls.tmp = tempfile.mkdtemp()
        cls.db = os.path.join(cls.tmp, "梱包資材マスタ.sqlite3")
        make_master(cls.db)
        cfg = Config()
        cfg.port = free_port()
        cfg.prefer_access = False
        cfg.material_csv_file = business_csv()
        cfg.material_db_file = cls.db
        cfg.__class__.local_dir = property(lambda self, d=cls.tmp: d)
        cls.cfg = cfg
        cls.httpd, cls.ctx = server.create_server(cfg)
        cls.thread = threading.Thread(target=cls.httpd.serve_forever,
                                      kwargs={"poll_interval": 0.1}, daemon=True)
        cls.thread.start()
        cls.base = "http://127.0.0.1:%d" % cfg.port

    @classmethod
    def tearDownClass(cls):
        cls.httpd.shutdown()
        cls.httpd.server_close()
        cls.thread.join(timeout=5)
        cls.ctx.store.close()
        Config.local_dir = cls._local_dir
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def post(self, path, payload):
        req = urllib.request.Request(self.base + path, data=json.dumps(payload).encode("utf-8"),
                                     method="POST", headers={"Content-Type": "application/json"})
        try:
            with urllib.request.urlopen(req, timeout=20) as r:
                return r.status, json.loads(r.read().decode("utf-8"))
        except urllib.error.HTTPError as exc:
            return exc.code, json.loads(exc.read().decode("utf-8"))

    def get(self, path):
        with urllib.request.urlopen(self.base + path, timeout=20) as r:
            return r.read().decode("utf-8")

    def test_足す_ほかのPCにも出る_外せる_消せる(self):
        pw = S.PATH_PASSWORD
        # 別の PC が足した(この PC の設定には何も書かれない)
        SX.add(self.ctx.wf.materials, self.ctx.wf.master, "1.2", "45", "BJB7699900QR")
        self.ctx.wf.size_extra.invalidate()
        self.ctx.wf.refresh_label_sheets(force=True)
        self.assertEqual(self.cfg.visible_sizes, ["1.0mm×53.5mm", "0.8mm×53.5mm"])
        home = self.get("/")
        self.assertIn("1.2mm×45.0mm", home, "ほかの PC が足したサイズも最初から出る")
        self.assertIn('class="size-row sz-extra"', home, "足したサイズにも色の帯")
        self.assertIn('<span class="sz-tag">追加</span>', home)
        # この PC で外す → 出ない。覚えるのは「外したもの」
        _, j = self.post("/api/settings/visible-sizes",
                         {"sizes": ["1.0mm×53.5mm", "0.8mm×53.5mm"], "password": pw})
        self.assertTrue(j["ok"], j)
        self.assertEqual(self.cfg.hidden_extra_sizes, ["1.2mm×45.0mm"])
        self.assertNotIn("1.2mm×45.0mm", self.get("/"))
        # 戻す
        _, j = self.post("/api/settings/visible-sizes",
                         {"sizes": ["1.0mm×53.5mm", "0.8mm×53.5mm", "1.2mm×45.0mm"], "password": pw})
        self.assertEqual(self.cfg.hidden_extra_sizes, [])
        self.assertIn("1.2mm×45.0mm", self.get("/"))
        # 選んでいたサイズを消すと、選択が外れる
        _, j = self.post("/api/select-size", {"obIdx": 15})
        self.assertTrue(j["ok"], j)
        _, j = self.post("/api/master/delete", {"table": SX.TABLE_NAME, "password": pw,
                                                "rowId": _rowid(self.db, "1.2mm×45.0mm")})
        self.assertTrue(j["ok"], j)
        self.assertNotIn("1.2mm×45.0mm", self.get("/"), "消したサイズは出ない")
        self.assertEqual(self.ctx.wf.load_state().selected_ob, 0)
        self.assertNotIn("1.2mm×45.0mm", self.get("/settings"))

    def test_マスタ管理の画面に消すボタンの決まり(self):
        html = self.get("/settings")
        self.assertIn("M.canDelete = !!j.canDelete", html)
        self.assertIn('mEl("mDelete").hidden = !row || !M.canDelete', html)


if __name__ == "__main__":
    unittest.main()
