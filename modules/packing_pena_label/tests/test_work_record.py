# -*- coding: utf-8 -*-
"""計算・印刷の記録を強くする(統合 1.2.39・現場の指示)

- 1 件ごとにソフトの版(統合・ペナラベル・中身の指紋)と PC 名・ライン名を残す
- 消えない控え: この PC の月ごとのファイルと、共有ログの出力先フォルダー(PC ごと)へ 1 行ずつ足す
- 印刷の行は「印刷画面」(刷った証拠ではない)。入力が変わったまま止めた印刷は「止めた印刷」で残す
- ラベルに載らない梱包は分けて残す(NW 0 の梱包に見えないように)
- 資材マスタの編集履歴: 誰が・いつ・どの値を何から何に(資材マスタの中の表に)
"""
import glob
import json
import os
import shutil
import sqlite3
import tempfile
import unittest

from modules.packing_pena_label.app.config import APP_VERSION
from modules.packing_pena_label.app.repositories.material_repo import MaterialRepository
from modules.packing_pena_label.app.services import master_admin as MA
from modules.packing_pena_label.app.services import size_extra as SX
from modules.packing_pena_label.app.services import work_record as WR
from modules.packing_pena_label.app.services.size_master import SizeMaster
from modules.packing_pena_label.app.services.workflow import Workflow

from .test_http import HttpTestCase
from .test_master_admin import make_master


def _cur(**kw):
    base = {"selectedOb": 14, "selectedCb": 0, "namedCb": False, "tip": "TIP1000",
            "kensaNo": "Z987654", "weight1": "", "weight2": "22.0",
            "coilH": {"1": "", "2": "", "3": "11", "4": "11"}}
    base.update(kw)
    return base


class RecordTest(HttpTestCase):
    @classmethod
    def setUpClass(cls):
        super().setUpClass()
        cls.shared = tempfile.mkdtemp(prefix="ppl_shared_")
        cls.cfg.record_dir = cls.shared
        cls.cfg.line_name = "2ライン"

    @classmethod
    def tearDownClass(cls):
        cls.cfg.record_dir = ""
        cls.cfg.line_name = ""
        shutil.rmtree(cls.shared, ignore_errors=True)
        super().tearDownClass()

    def _calc(self, **kw):
        self.post("/api/clear", {"current": _cur(**kw)})
        self.post("/api/apply-weight", {"current": _cur(**kw)})
        _, j = self.post("/api/calc-tare", {"current": _cur(**kw)})
        self.assertTrue(j["ok"], j)

    def test_版とPC名_ライン名が残る(self):
        self._calc()
        row = self.ctx.wf.recent_log(1)[0]
        v, w = row["data"]["version"], row["data"]["who"]
        self.assertEqual(v["pena"], APP_VERSION)
        self.assertTrue(v.get("app"), "統合ツールの版")
        self.assertTrue(v.get("stamp"), "中身の指紋")
        self.assertEqual(w["pc"], WR.pc_name())
        self.assertEqual(w["line"], "2ライン")

    def test_消えない控えをこのPCと共有へ月ごとに足す(self):
        self._calc()
        local = glob.glob(os.path.join(self.cfg.local_dir, "records", "ペナラベル記録_*.jsonl"))
        shared = glob.glob(os.path.join(self.shared, WR.pc_name(), "ペナラベル記録_*.jsonl"))
        self.assertEqual(len(local), 1)
        self.assertEqual(len(shared), 1)
        for path in (local[0], shared[0]):
            with open(path, encoding="utf-8") as f:
                last = json.loads(f.read().splitlines()[-1])
            self.assertEqual((last["kind"], last["title"]), ("計算", "風袋計算"))
            self.assertEqual(last["kensaNo"], "Z987654")
            self.assertEqual(last["who"]["line"], "2ライン")
            self.assertEqual(last["version"]["pena"], APP_VERSION)
        _, html = self.get("/print-log")
        self.assertIn("消えない控え", html)
        self.assertIn(self.shared, html)
        self.assertIn("2ライン", html)

    def test_印刷画面と止めた印刷(self):
        self._calc()
        _, j = self.post("/api/print-log", {"path": "/labels/print", "search": "?ob=14"})
        self.assertEqual(self.ctx.wf.recent_log(1)[0]["kind"], "印刷画面")
        # 同じ検番で重量だけ変えて反映(計算し直さない)→ 印刷ビューは止めて、記録に残す
        self.post("/api/apply-weight", {"current": _cur(weight2="30.0")})
        _, page = self.get("/labels/print?ob=14")
        self.assertIn("印刷する（停止中）", page)
        row = self.ctx.wf.recent_log(1)[0]
        self.assertEqual((row["kind"], row["title"]), ("止めた印刷", "小ラベル"))
        self.assertIn("印刷を止めた", row["data"]["stopReason"])
        self.get("/tare/print")
        row = self.ctx.wf.recent_log(1)[0]
        self.assertEqual((row["kind"], row["title"]), ("止めた印刷", "風袋計算"))
        _, html = self.get("/print-log")
        self.assertIn("止めた印刷", html)
        self.assertIn("刷った証拠にはなりません", html)


class NotOnLabelTest(unittest.TestCase):
    def test_ラベルに載らない梱包は分ける(self):
        coils = {"1": {"本数": "11", "NW": "110.0", "GW": "147.0"},
                 "2": {"本数": "5", "NW": "", "GW": ""},          # NW が無い
                 "3": {"本数": "11", "NW": "120.0", "GW": "155.0"}}   # 隠れている丈
        on, off = Workflow._split_label_coils(coils, [1, 2])
        self.assertEqual(sorted(on), ["1"])
        self.assertEqual(sorted(off), ["2", "3"])


class MasterHistoryTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp()
        self.db = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        make_master(self.db)
        self.repo = MaterialRepository(accdb_dir="", csv_path="", prefer_access=False,
                                       db_path=self.db, refresh_sec=0)
        self.who = {"pc": "LINE2-PC", "line": "2ライン", "user": "tanaka"}

    def tearDown(self):
        shutil.rmtree(self.tmp, ignore_errors=True)

    def _rowid(self, name):
        with sqlite3.connect(self.db) as c:
            return c.execute("SELECT rowid FROM 資材重量 WHERE 梱包資材名=?", (name,)).fetchone()[0]

    def test_誰が_いつ_どの値を何から何に(self):
        MA.save_row(self.repo, "資材重量", self._rowid("テスラピン"), {"単位質量": "0.6"}, who=self.who)
        h = MA.history(self.repo)["rows"]
        self.assertEqual(len(h), 1, "変わった列だけ")
        r = h[0]
        self.assertEqual((r["表"], r["操作"], r["列"], r["前"], r["後"]),
                         ("資材重量", "更新", "単位質量", "0.5", "0.6"))
        self.assertEqual((r["PC"], r["ライン"], r["利用者"]), ("LINE2-PC", "2ライン", "tanaka"))
        self.assertIn("梱包資材名=テスラピン", r["行の名前"])
        self.assertRegex(r["日時"], r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}$")

    def test_足す_消す_サイズを足すも残る(self):
        new = MA.save_row(self.repo, "資材重量", None,
                          {"管理番号": "9", "梱包資材名": "テスト", "単位質量": "1", "係数": "1"},
                          who=self.who)
        MA.delete_row(self.repo, "資材重量", new["rowId"], who=self.who)
        SX.add(self.repo, SizeMaster(), "1.2", "45", "BJB7699900QR", who=self.who)
        rows = MA.history(self.repo)["rows"]
        ops = [(r["表"], r["操作"]) for r in rows]
        self.assertIn(("資材重量", "追加"), ops)
        self.assertIn(("資材重量", "削除"), ops)
        self.assertIn((SX.TABLE_NAME, "追加"), ops)
        gone = [r for r in rows if r["操作"] == "削除" and r["列"] == "単位質量"]
        self.assertEqual((gone[0]["前"], gone[0]["後"]), ("1", ""))

    def test_履歴の表が無ければ空(self):
        self.assertEqual(MA.history(self.repo)["rows"], [])


if __name__ == "__main__":
    unittest.main()
