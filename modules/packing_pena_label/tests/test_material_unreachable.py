"""資材マスタに届かないとき(統合 1.2.31・現場の指摘)

以前は SQLite / Access に届かないと、同梱の data/資材重量.csv(試験用のサンプル値)に黙って落ち、
間違った NW/GW のラベルが刷れた。作業の途中の瞬断でもサンプルに入れ替わった。

    朝から届かない        → 止める(「資材重量なし」と、何が読めなかったか・どこを見ればよいか)
    途中で届かなくなった  → 直前に読んだ内容で続ける(画面右上を赤く「古い内容」)。届けば戻る
    設定で CSV を指定した → その CSV を読む(業務の取得元として)
    同梱のサンプル        → どの場合も読まない
"""
from __future__ import annotations

import os
import shutil
import sqlite3
import tempfile
import unittest

from modules.packing_pena_label.app.repositories.access_bridge import AccessError
from modules.packing_pena_label.app.repositories.material_repo import MaterialRepository

from ._material import SAMPLE, business_csv


def _sqlite(path: str, mass: str) -> str:
    conn = sqlite3.connect(path)
    conn.execute('CREATE TABLE "資材重量" (管理番号, 梱包資材名, 単位質量, 係数)')
    conn.execute("INSERT INTO 資材重量 VALUES (1, 'テスラピン', ?, 1)", (mass,))
    conn.commit()
    conn.close()
    return path


def _mass(table) -> str:
    return table.as_rows()[0]["mass"]


class UnreachableTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.dir, True)
        self.db = os.path.join(self.dir, "梱包資材マスタ.sqlite3")

    def repo(self, csv_path: str = "") -> MaterialRepository:
        return MaterialRepository("", csv_path, prefer_access=False, refresh_sec=0, db_path=self.db)

    def test_朝から届かなければ止めて理由を言う_同梱のサンプルは読まない(self):
        self.assertTrue(os.path.exists(SAMPLE))           # サンプルはあっても読まない
        repo = self.repo()
        with self.assertRaises(AccessError) as ctx:
            repo.load()
        text = str(ctx.exception)
        self.assertIn("計算を止めました", text)
        self.assertIn(self.db, text)                       # 何が読めなかったか
        self.assertIn("確かめること", text)                 # どこを見ればよいか
        self.assertEqual(repo.last_source, "")

    def test_途中で届かなくなったら直前の内容で続け_届けば戻る(self):
        _sqlite(self.db, "9.99")
        repo = self.repo()
        self.assertEqual(_mass(repo.load()), "9.99")
        os.rename(self.db, self.db + ".x")                 # 共有フォルダに届かなくなった
        table = repo.load(check_now=True)
        self.assertEqual(_mass(table), "9.99", "サンプルの単重に入れ替わってはいけない")
        self.assertTrue(repo.serving_stale)
        self.assertEqual(repo.last_source, "sqlite")
        os.rename(self.db + ".x", self.db)                 # 届くようになった
        repo.load(check_now=True)
        self.assertFalse(repo.serving_stale, "届いたら「古い内容」を外す")

    def test_指定したCSVは届かないときに読む(self):
        repo = self.repo(business_csv())
        repo.load()
        self.assertEqual(repo.last_source, "csv")

    def test_計算は止まり_画面に理由が出る(self):
        from modules.packing_pena_label.app.config import Config
        from modules.packing_pena_label.app.services.size_master import SizeMaster
        from modules.packing_pena_label.app.repositories.sqlite_store import Store
        from modules.packing_pena_label.app.services.workflow import Workflow
        from modules.packing_pena_label.app.models import FormState
        from ._material import ROOT
        cfg = Config()
        cfg.prefer_access = False
        wf = Workflow(Store(os.path.join(self.dir, "state.sqlite3")),
                      SizeMaster(os.path.join(ROOT, "data", "size_master.json")), self.repo(), cfg)
        state = FormState()
        state.selected_ob = 14
        state.tip = "TIP1000"
        state.weight2 = "22.0"
        state.lbl_weight2 = "22.0kg"          # 重量反映で写した値
        state.coil_h = {1: "", 2: "", 3: "11", 4: "11"}
        res = wf.calc_tare(state)
        self.assertFalse(res.ok)
        self.assertIn("資材重量なし", res.message)
        self.assertIn("計算を止めました", res.message)


if __name__ == "__main__":
    unittest.main()
