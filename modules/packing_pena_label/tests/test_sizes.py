# -*- coding: utf-8 -*-
"""設定の「サイズ」(統合 1.2.16・現場の指定)

- 指定サイズの画面に出すサイズを、チェックで選べる(この PC。合言葉が要る)
- 新しいサイズを足せる(共有の資材マスタの表「ペナラベル追加サイズ」。全ライン。合言葉が要る)
- 足したサイズは丈1・丈2・同時のどれでも選べ、巾が計算に使われる
"""
import json
import os
import re
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

from .test_http import HttpTestCase
from .test_master_admin import free_port, make_master
from ._material import business_csv  # noqa: E402


class AddSizeTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.db = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        make_master(self.db)
        self.repo = MaterialRepository(accdb_dir="", csv_path="", prefer_access=False,
                                       db_path=self.db, refresh_sec=0)
        self.master = SizeMaster()

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def test_足すと共有マスタに書き_番号は15から(self):
        r = SX.add(self.repo, self.master, "1.2", "45", "bjb7699900qr", "h", "2026.10.14追加")
        self.assertEqual((r["baseName"], r["obIdx"], r["kataban"]),
                         ("1.2mm×45.0mm", [15, 16], "BJB7699900QR"))
        rows = SX.read_rows(self.db)
        self.assertEqual(rows[0]["widthMm"], 45.0)
        self.assertEqual(rows[0]["note"], "2026.10.14追加")
        self.assertTrue(r["backup"] and os.path.exists(r["backup"]), "書く前に控えを取る")
        # 2 つ目は 17
        self.master.apply_extra(rows)
        r2 = SX.add(self.repo, self.master, "０．８", "６３．５", "BJB7699800QR")   # 全角でも読む
        self.assertEqual((r2["baseName"], r2["obIdx"]), ("0.8mm×63.5mm", [17, 18]))

    def test_同じ寸法_形の違う数_型番なしは断る(self):
        cases = (("1.0", "53.5", "BJB1"),            # 同梱にある(表記が違っても寸法で見る)
                 ("1", "53.50", "BJB7600000QR"),
                 ("1.2", "abc", "BJB7600000QR"),
                 ("0", "45", "BJB7600000QR"),
                 ("1.2345", "45", "BJB7600000QR"),
                 ("1.2", "45", ""),
                 ("1.2", "45", "型番"))
        for t, w, k in cases:
            with self.subTest(t=t, w=w, k=k):
                with self.assertRaises(MA.Refused):
                    SX.add(self.repo, self.master, t, w, k)
        SX.add(self.repo, self.master, "1.2", "45", "BJB7699900QR")
        with self.assertRaises(MA.Refused, msg="表の中の重なりも断る"):
            SX.add(self.repo, self.master, "1.20", "45.0", "BJB7699901QR")
        self.assertEqual(len(SX.read_rows(self.db)), 1)

    def test_書き先が無ければ理由を出して断る(self):
        repo = MaterialRepository(accdb_dir="", csv_path="", prefer_access=False, db_path="")
        with self.assertRaises(MA.Refused) as cm:
            SX.add(repo, self.master, "1.2", "45", "BJB7699900QR")
        self.assertIn("SQLite", cm.exception.message)

    def test_足したサイズは丈1丈2同時とも選べる(self):
        SX.add(self.repo, self.master, "1.2", "45", "BJB7699900QR")
        self.master.apply_extra(SX.read_rows(self.db))
        c1, c2 = self.master.get(15), self.master.get(16)
        self.assertEqual((c1.take_type, c2.take_type), (1, 2))
        self.assertEqual(c1.width_mm, 45.0)
        self.assertEqual(self.master.width_from_sheet_name(15), 45.0)
        cb = int(c1.cb_key.replace("CheckBox", ""))
        self.assertEqual(self.master.cb_to_pair(cb), (15, 16))
        self.assertEqual(self.master.cb_to_pair(5), (5, 6), "同梱の組は今までどおり")
        # 読み直しても残る・表から消えたら外れる
        self.master.reload(None)
        self.assertIsNotNone(self.master.get(15))
        self.master.apply_extra([])
        self.assertIsNone(self.master.get(15))

    def test_表が無い資材マスタでも今までどおり(self):
        self.assertEqual(SX.read_rows(self.db), [])
        self.assertEqual(self.master.apply_extra(SX.read_rows(self.db)), 0)
        self.assertEqual(len(self.master.all()), 14)

    def test_マスタ管理では型番記号備考だけ直せる(self):
        self.assertEqual(MA.editable_columns(SX.TABLE_NAME, list(SX.COLUMNS)), ["型番", "記号", "備考"])
        self.assertTrue(MA.fixed_rows(SX.TABLE_NAME), "行の追加・削除はマスタ管理からはしない")


class SizeSettingsHttpTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls._local_dir = Config.__dict__["local_dir"]
        cls.tmp = tempfile.mkdtemp()
        cls.db = os.path.join(cls.tmp, "梱包資材マスタ.sqlite3")
        make_master(cls.db)
        cfg = Config()
        cfg.port = free_port()
        cfg.prefer_access = False
        cfg.material_csv_file = business_csv()   # 同梱サンプルの写しを指定した CSV として
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

    def test_1_設定にサイズのタブと全サイズのチェック(self):
        html = self.get("/settings")
        self.assertIn('data-sub="sizes"', html)
        self.assertIn('id="sizesPane" data-live', html)
        self.assertEqual(html.count("data-vsize="), 7, "同梱の 7 サイズ")
        checked = re.findall(r'data-vsize="([^"]+)" checked>', html)
        self.assertEqual(checked, ["0.8mm×53.5mm", "1.0mm×53.5mm"], "出しているのは既定の 2 つ")

    def test_2_出すサイズの保存は合言葉が要る_この端末に残る(self):
        want = ["1.0mm×53.5mm", "1.0mm×40.0mm"]
        st, j = self.post("/api/settings/visible-sizes", {"sizes": want})
        self.assertTrue(j.get("needPassword"), j)
        self.assertNotIn("1.0mm×40.0mm", self.get("/"))
        _, j = self.post("/api/settings/visible-sizes", {"sizes": [], "password": S.PATH_PASSWORD})
        self.assertFalse(j["ok"], "1 つも出さないのは断る")
        _, j = self.post("/api/settings/visible-sizes", {"sizes": ["9.9mm×9.9mm"], "password": S.PATH_PASSWORD})
        self.assertFalse(j["ok"])
        _, j = self.post("/api/settings/visible-sizes", {"sizes": want, "password": S.PATH_PASSWORD})
        self.assertTrue(j["ok"], j)
        home = self.get("/")
        self.assertIn("1.0mm×40.0mm", home)
        self.assertNotIn("0.8mm×53.5mm", home)
        with open(self.cfg.local_config_path, encoding="utf-8") as f:
            self.assertEqual(json.load(f)["visible_sizes"], ["1.0mm×53.5mm", "1.0mm×40.0mm"])

    def test_3_サイズを足すと選べて計算できる(self):
        body = {"thickness": "1.2", "width": "45", "kataban": "BJB7699900QR",
                "note": "2026.10.14追加", "show": True}
        _, j = self.post("/api/settings/size-add", body)
        self.assertTrue(j.get("needPassword"), j)
        self.assertEqual(SX.read_rows(self.db), [], "合言葉なしでは書かない")
        _, j = self.post("/api/settings/size-add", dict(body, password=S.PATH_PASSWORD))
        self.assertTrue(j["ok"], j)
        home = self.get("/")
        self.assertIn("1.2mm×45.0mm", home, "足した PC の画面に出る")
        self.assertIn("2026.10.14追加", home, "備考は赤字の注記に")
        st, j = self.post("/api/settings/size-add", dict(body, password=S.PATH_PASSWORD))
        self.assertEqual(st, 400)
        self.assertFalse(j["ok"], "同じサイズは 2 度足さない")
        # 丈1 を選ぶ・同時を選ぶ
        _, j = self.post("/api/select-size", {"obIdx": 15})
        self.assertTrue(j["ok"], j)
        cb = int(self.ctx.wf.master.get(15).cb_key.replace("CheckBox", ""))
        _, j = self.post("/api/select-checkbox", {"cbIdx": cb})
        self.assertTrue(j["ok"], j)
        self.assertIn('data-cb="%d"' % cb, self.get("/"))
        # 設定の一覧にも「足したサイズ」として出る
        self.assertIn("足したサイズ(共有の資材マスタ)", self.get("/settings"))
        # ラベル台紙の一覧と、その台紙の印刷ビューも出る
        self.assertIn("1.2mm×45.0mm", self.get("/labels"))
        for ob in (15, 16):
            with self.subTest(ob=ob):
                self.assertIn("1.2mm×45.0mm", self.get("/labels/sheet?ob=%d" % ob))


    def test_4_ラベル台紙にサイズの足し方と設定へのボタン(self):
        """ラベル台紙の一覧に「新しいサイズを足すには」(統合 1.2.18・現場の求め)。"""
        html = self.get("/labels")
        self.assertIn('<details class="card howto" id="sizeHowto">', html, "ふだんは閉じておく")
        self.assertIn("新しいサイズを足すには", html)
        self.assertIn('href="/settings#sizes"', html)
        self.assertIn("このPCからサイズを足せます", html, "資材マスタが SQLite なら足せると出す")
        # 足したサイズには「追加」の印(test_3 で 1.2mm×45.0mm を足している)
        self.assertIn('1.2mm×45.0mm 　丈1 <span class="tag-extra"', html)
        self.assertNotIn('1.0mm×53.5mm 　丈1 <span class="tag-extra"', html)
        # 設定を #sizes で開くと「サイズ」のタブを開く
        self.assertIn('location.hash === "#sizes"', self.get("/settings"))
        # 中の面として読むとき(pane=1)も出る
        self.assertIn('id="sizeHowto"', self.get("/labels?pane=1"))


class SizeHowtoCsvTest(HttpTestCase):
    """資材マスタが CSV のときは、ラベル台紙で「いまは足せない」と理由を出す。"""

    @classmethod
    def setUpClass(cls):
        cls._local_dir2 = Config.__dict__["local_dir"]
        super().setUpClass()

    @classmethod
    def tearDownClass(cls):
        super().tearDownClass()
        Config.local_dir = cls._local_dir2

    def test_足せない理由(self):
        _, html = self.get("/labels")
        self.assertIn("いまこのPCからはサイズを足せません", html)
        self.assertIn("SQLite", html)
        self.assertIn("まだ足したサイズはありません", html)

if __name__ == "__main__":
    unittest.main()
