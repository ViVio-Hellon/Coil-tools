# -*- coding: utf-8 -*-
"""マスタを直したら、前の中身で計算・印刷しない(統合 1.2.49・現場の指摘)

現場の依頼: 「マスタを直しても前の中身のまま動く・表示する事がないかのチェックをしてください
例: 資材重量 変更後 前の重量のまま計算」。点検で見つかったことを、直したあとの決まりで確かめる。

    1  資材重量を直したら、前の NW/GW のラベル・風袋計算を刷らない(どの資材が変わったかを出す)
    2  羅列計算も同じ
    3  全サイズは、決定したあとに型番マスタが変わったら刷らない
    4  ラベル台紙一覧の型番・シート名を空にした・行を消した → JSON の値に戻る(再起動した PC と同じ)。
       型番が空ならラベルを刷らない
    5  資材マスタの更新日時は「読む前」に取る(読んでいる途中の書き込みを取りこぼさない)
    6  「計算」のときは SQLite・CSV を毎回読み直す(大きさ・更新日時が変わらない書き込み・WAL も拾う)
    7  前の内容(届かない)・CSV に落ちて計算したら、計算の返事を注意にして理由を出す
    8  計算内容の画面に、計算したときの資材マスタの時点と、いまのマスタと違う資材を出す
    9  size_master.json を外で直したら、再起動しなくても読み直す
"""
from __future__ import annotations

import csv
import json
import os
import shutil
import sqlite3
import tempfile
import time
import unittest

from modules.packing_pena_label.app.config import Config
from modules.packing_pena_label.app.models import FormState
from modules.packing_pena_label.app.repositories.material_repo import MaterialRepository
from modules.packing_pena_label.app.repositories.sqlite_store import Store
from modules.packing_pena_label.app.routes.pages import PageRoutes
from modules.packing_pena_label.app.services.label_sheet_master import LabelSheetMaster
from modules.packing_pena_label.app.services.size_master import SizeMaster
from modules.packing_pena_label.app.services.workflow import Workflow

from ._material import ROOT, SAMPLE

JSON_SRC = os.path.join(ROOT, "data", "size_master.json")


def _sample_rows():
    with open(SAMPLE, encoding="utf-8-sig") as f:
        return list(csv.reader(f))[1:]


def make_db(path: str, wal: bool = False) -> None:
    """業務の形の梱包資材マスタ(資材重量・ラベル台紙一覧・型番)。"""
    conn = sqlite3.connect(path)
    if wal:
        conn.execute("PRAGMA journal_mode=WAL")
    conn.execute('CREATE TABLE "資材重量" (管理番号 INTEGER, 梱包資材名 TEXT, 単位質量 TEXT, 係数 TEXT)')
    conn.executemany('INSERT INTO "資材重量" VALUES (?,?,?,?)', _sample_rows())
    conn.execute('CREATE TABLE "ラベル台紙一覧" (管理番号 INTEGER, シート名 TEXT, 型番 TEXT)')
    conn.execute('CREATE TABLE "型番" (サイズ名 TEXT, 記号 TEXT, 型番 TEXT, 備考 TEXT)')
    conn.execute('INSERT INTO "型番" VALUES (?,?,?,?)', ("1.0×53.5×Coil", "d", "BJB7606500QR", ""))
    conn.commit()
    conn.close()


def sql(path: str, query: str, args=()) -> None:
    conn = sqlite3.connect(path)
    conn.execute(query, args)
    conn.commit()
    conn.close()


def set_mass(path: str, name: str, mass: str) -> None:
    sql(path, 'UPDATE "資材重量" SET 単位質量=? WHERE 梱包資材名=?', (mass, name))


class _Base(unittest.TestCase):
    """SQLite を設定した PC 1 台分(ツールは動いたまま)。"""

    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="ppl_fresh_")
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.db = os.path.join(self.tmp, "梱包資材マスタ.sqlite3")
        make_db(self.db)
        self.json = os.path.join(self.tmp, "size_master.json")
        shutil.copyfile(JSON_SRC, self.json)
        self.cfg = Config()
        self.cfg.prefer_access = False
        self.cfg.material_db_file = self.db
        self.cfg.size_master_file = self.json
        self.cfg.kataban_csv_file = ""
        self.cfg.record_dir = ""
        # 記録の控え(この PC の月ごとのファイル)を一時領域へ
        self.cfg.__class__ = type("Cfg", (Config,), {"local_dir": property(lambda s, d=self.tmp: d)})
        self.store = Store(os.path.join(self.tmp, "state.sqlite3"))
        self.addCleanup(self.store.close)
        self.master = SizeMaster(self.json)
        # 鮮度チェックの間隔は既定(300 秒)のまま。間隔の中でも「計算」「印刷」で確かめることを見る
        self.repo = MaterialRepository("", "", prefer_access=False, refresh_sec=300, db_path=self.db)
        self.wf = Workflow(self.store, self.master, self.repo, self.cfg)
        self.pages = PageRoutes(self.wf, self.cfg)

    def calc(self, kensa="W111111") -> FormState:
        st = FormState()
        self.wf.select_checkbox(st, 5)
        st.tip = "TIP1000"
        st.kensa_no = kensa
        st.weight1, st.weight2 = "10", "12"
        r = self.wf.apply_weight(st)
        self.assertTrue(r.ok, r.message)
        st.coil_h = {1: "11", 2: "11", 3: "10", 4: "10"}
        r = self.wf.calc_tare(st)
        self.assertTrue(r.ok, r.message)
        self.last = r
        return st

    def stopped_logs(self):
        return [r for r in self.wf.recent_log() if r.get("kind") == Workflow.KIND_STOPPED]


# ---------------------------------------------------------------- 1
class MaterialChangedAfterCalcTest(_Base):
    def test_資材重量を直したらラベルと風袋計算を刷らない(self):
        self.calc()
        self.assertFalse(self.wf.calc_is_stale())
        self.assertFalse(self.wf.tare_is_stale())
        # 大きさの変わらない直し方(0.012 → 0.013)。SQLite のファイルの大きさは変わらない
        set_mass(self.db, "PETバンド", "0.013")

        self.assertTrue(self.wf.calc_is_stale(), "資材重量が変わったのに、ラベルが古いと分からない")
        self.assertTrue(self.wf.tare_is_stale(), "資材重量が変わったのに、風袋計算が古いと分からない")
        # メイン画面・台紙一覧の「止めている」表示にも効く
        self.assertTrue(self.wf.label_view(5)["stale"])

        html = self.pages.label_print([5, 6])
        self.assertIn("資材重量が計算のあとに変わりました", html)
        self.assertIn("PETバンド", html)
        self.assertIn("0.012 → 0.013", html)
        self.assertIn("印刷する（停止中）", html)
        self.assertNotIn("window.print();});", html)       # 開いたときに自分で刷らない

        html = self.pages.tare(print_view=True)
        self.assertIn("資材重量が計算のあとに変わりました", html)
        self.assertIn("印刷する（停止中）", html)

        html = self.pages.labels()
        self.assertIn("資材重量が計算のあとに変わりました", html)

        reasons = [r["data"].get("stopReason", "") for r in self.stopped_logs()]
        self.assertTrue(reasons and all("資材重量" in x for x in reasons), reasons)
        self.assertTrue(any("PETバンド" in x for x in reasons), reasons)

        # 計算し直せば刷れる
        self.calc()
        self.assertFalse(self.wf.calc_is_stale())
        self.assertFalse(self.wf.tare_is_stale())
        self.assertNotIn("停止中", self.pages.label_print([5, 6]))

    def test_マスタが読めないときは止めない(self):
        """ハッシュが取れない(届かない)なら比べない(2026.09 決定 D-8 と同じ扱い)。"""
        self.calc()
        os.rename(self.db, self.db + ".x")
        self.repo.clear_cache()            # 前の内容も無い(読めない)
        self.assertFalse(self.wf.calc_is_stale())
        self.assertFalse(self.wf.tare_is_stale())

    def test_前の版で計算した控え_資材の控えが無い_は止めない(self):
        self.calc()
        saved = self.store.get_kv(Workflow._CALC_INPUT_KEY)
        saved.pop("material", None)
        self.store.set_kv(Workflow._CALC_INPUT_KEY, saved)
        set_mass(self.db, "PETバンド", "0.013")
        self.assertFalse(self.wf.calc_is_stale())


# ---------------------------------------------------------------- 2
class ListChangedAfterCalcTest(_Base):
    def test_羅列計算も資材重量が変わったら刷らない(self):
        st = self.calc()
        r = self.wf.calc_list(st)
        self.assertTrue(r.ok, r.message)
        self.assertNotIn("停止中", self.pages.list_page(print_view=True))
        set_mass(self.db, "樹脂パレット", "7.6")
        html = self.pages.list_page(print_view=True)
        self.assertIn("資材重量が計算のあとに変わりました", html)
        self.assertIn("樹脂パレット", html)
        self.assertIn("印刷する（停止中）", html)
        self.assertIn("資材重量が計算のあとに変わりました", self.pages.list_page())
        self.assertTrue(any("資材重量" in r["data"].get("stopReason", "")
                            and r["title"] == "羅列計算" for r in self.stopped_logs()))


# ---------------------------------------------------------------- 3
class AllSizeKatabanChangedTest(_Base):
    def test_決定のあとに型番マスタが変わったら刷らない(self):
        svc = self.wf.all_size
        r = svc.submit("1.0×53.5×Coil", "N6X0945", False, "10", "12", "", "")
        self.assertTrue(r["ok"], r)
        html = self.pages.all_size_print("1.0×53.5×Coil")
        self.assertNotIn("停止中", html)
        self.assertIn("BJB7606500QR", html)
        # 大きさの変わらない直し方(1 文字だけ)
        sql(self.db, 'UPDATE "型番" SET 型番=? WHERE サイズ名=?', ("BJB7606500QX", "1.0×53.5×Coil"))
        html = self.pages.all_size_print("1.0×53.5×Coil")
        self.assertIn("型番マスタが変わりました", html)
        self.assertIn("BJB7606500QR → BJB7606500QX", html)
        self.assertIn("決定し直してください", html)
        self.assertIn("印刷する（停止中）", html)
        self.assertTrue(any("型番" in r["data"].get("stopReason", "")
                            for r in self.stopped_logs()))
        # 決定し直せば刷れる(新しい型番で)
        svc.submit("1.0×53.5×Coil", "N6X0945", False, "10", "12", "", "")
        html = self.pages.all_size_print("1.0×53.5×Coil")
        self.assertNotIn("停止中", html)
        self.assertIn("BJB7606500QX", html)


# ---------------------------------------------------------------- 4
class LabelSheetOverlayTest(_Base):
    def _restarted(self) -> SizeMaster:
        """再起動した PC(JSON から読み、ラベル台紙一覧を重ねる)。"""
        m = SizeMaster(self.json)
        LabelSheetMaster().apply_to(m, self.db, force=True)
        return m

    def _same_as_restarted(self, ob: int):
        self.wf.refresh_label_sheets(force=True)
        fresh = self._restarted()
        self.assertEqual((self.master.get(ob).kataban, self.master.get(ob).size_cell_label),
                         (fresh.get(ob).kataban, fresh.get(ob).size_cell_label))
        return self.master.get(ob)

    def test_空にした_行を消した_は動いたままのPCもJSONの値に戻る(self):
        sql(self.db, 'INSERT INTO "ラベル台紙一覧" VALUES (1, "べつの名前 丈1", "AAA111")')
        cfg = self._same_as_restarted(1)
        self.assertEqual((cfg.kataban, cfg.size_cell_label), ("AAA111", "べつの名前 丈1"))

        sql(self.db, 'UPDATE "ラベル台紙一覧" SET 型番="", シート名="" WHERE 管理番号=1')
        cfg = self._same_as_restarted(1)
        self.assertEqual(cfg.kataban, "BJB7604000QR", "空にしたのに前の型番が残っている")
        self.assertEqual(cfg.size_cell_label, "1.0mm×73.0mm 　丈1")

        sql(self.db, 'UPDATE "ラベル台紙一覧" SET 型番="BBB222" WHERE 管理番号=1')
        self.assertEqual(self._same_as_restarted(1).kataban, "BBB222")
        sql(self.db, 'DELETE FROM "ラベル台紙一覧"')
        self.assertEqual(self._same_as_restarted(1).kataban, "BJB7604000QR",
                         "行を消したのに前の型番が残っている")

    def test_表が読めなくなったらJSONの値に戻る(self):
        sql(self.db, 'INSERT INTO "ラベル台紙一覧" VALUES (1, "", "AAA111")')
        self.assertEqual(self._same_as_restarted(1).kataban, "AAA111")
        sql(self.db, 'DROP TABLE "ラベル台紙一覧"')
        self.assertEqual(self._same_as_restarted(1).kataban, "BJB7604000QR")

    def test_共有に一時的に届かない_ロックのときは前に読めた内容で続ける(self):
        """表がもともと無いのとは違い、一時的に読めないだけなら JSON の値へ黙って戻さない
        (戻すと、上書きしていた型番と違う型番を刷る。資材マスタの D-8 と同じ考え)。"""
        from unittest import mock
        from modules.packing_pena_label.app.services import label_sheet_master as LSM
        sql(self.db, 'INSERT INTO "ラベル台紙一覧" VALUES (1, "", "AAA111")')
        self.wf.refresh_label_sheets(force=True)
        self.assertEqual(self.master.get(1).kataban, "AAA111")
        # ロック・I/O の誤り
        with mock.patch.object(LSM.sqlite3, "connect",
                               side_effect=sqlite3.OperationalError("database is locked")):
            self.wf.refresh_label_sheets(force=True)
        self.assertEqual(self.master.get(1).kataban, "AAA111", "一時的に読めないのに JSON の型番へ戻った")
        # 共有に届かない(ファイルが見えない)
        moved = self.db + ".away"
        os.replace(self.db, moved)
        try:
            self.wf.refresh_label_sheets(force=True)
            self.assertEqual(self.master.get(1).kataban, "AAA111")
        finally:
            os.replace(moved, self.db)
        # 届くようになったら読み直す
        sql(self.db, 'UPDATE "ラベル台紙一覧" SET 型番="CCC333" WHERE 管理番号=1')
        self.wf.refresh_label_sheets(force=True)
        self.assertEqual(self.master.get(1).kataban, "CCC333")

    def test_型番が空ならラベルを刷らない_重量反映のときの型番を使わない(self):
        with open(self.json, encoding="utf-8") as f:
            data = json.load(f)
        for row in data["sizes"]:
            if 5 in row["obIdx"]:
                row["kataban"] = ""
        sql(self.db, 'INSERT INTO "ラベル台紙一覧" VALUES (5, "", "CCC333")')
        sql(self.db, 'INSERT INTO "ラベル台紙一覧" VALUES (6, "", "CCC333")')
        with open(self.json, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        self.master.reload(self.json)
        self.wf.refresh_label_sheets(force=True)
        self.calc()                          # 重量反映のときは CCC333
        self.assertEqual(self.store.get_sheet_header(5)["kataban"], "CCC333")
        sql(self.db, 'UPDATE "ラベル台紙一覧" SET 型番="" WHERE 管理番号 IN (5, 6)')
        v = self.wf.label_view(5)
        self.assertEqual(v["printHeader"]["kataban"], "")
        self.assertNotIn("CCC333", json.dumps(v["labels"], ensure_ascii=False))
        html = self.pages.label_print([5, 6])
        self.assertIn("型番が空", html)
        self.assertIn("印刷する（停止中）", html)
        self.assertNotIn("CCC333", html.split("stopprint")[-1])


# ---------------------------------------------------------------- 5・6
class StampTest(unittest.TestCase):
    def setUp(self):
        self.tmp = tempfile.mkdtemp(prefix="ppl_stamp_")
        self.addCleanup(shutil.rmtree, self.tmp, True)

    def _pet(self, repo, **kw):
        return repo.load(**kw).unit_of("PETバンド")[0]

    def test_読んでいる途中の書き込みを取りこぼさない_SQLite(self):
        db = os.path.join(self.tmp, "a.sqlite3")
        make_db(db)
        repo = MaterialRepository("", "", prefer_access=False, refresh_sec=0, db_path=db)
        orig = MaterialRepository._load_sqlite

        def racing(path):
            table = orig(path)
            time.sleep(0.02)
            set_mass(path, "PETバンド", "3.33")        # 読んだあと、更新日時を取る前に書かれた
            return table
        MaterialRepository._load_sqlite = staticmethod(racing)
        try:
            self.assertEqual(self._pet(repo), "0.012")
        finally:
            MaterialRepository._load_sqlite = staticmethod(orig)
        self.assertEqual(self._pet(repo), "3.33", "読む前の更新日時を控えていない")

    def test_読んでいる途中の書き込みを取りこぼさない_CSV(self):
        path = os.path.join(self.tmp, "m.csv")
        shutil.copyfile(SAMPLE, path)
        repo = MaterialRepository("", path, prefer_access=False, refresh_sec=0)
        orig = MaterialRepository._load_csv

        def racing(p):
            table = orig(p)
            time.sleep(0.02)
            with open(p, encoding="utf-8-sig") as f:
                text = f.read()
            with open(p, "w", encoding="utf-8") as f:
                f.write(text.replace("PETバンド,0.012", "PETバンド,0.019"))
            return table
        MaterialRepository._load_csv = staticmethod(racing)
        try:
            self.assertEqual(self._pet(repo), "0.012")
        finally:
            MaterialRepository._load_csv = staticmethod(orig)
        self.assertEqual(self._pet(repo), "0.019")

    def test_計算のときは大きさも更新日時も変わらない書き込みを拾う(self):
        db = os.path.join(self.tmp, "b.sqlite3")
        make_db(db)
        repo = MaterialRepository("", "", prefer_access=False, refresh_sec=300, db_path=db)
        self.assertEqual(self._pet(repo), "0.012")
        st = os.stat(db)
        set_mass(db, "PETバンド", "7.77")
        os.utime(db, ns=(st.st_atime_ns, st.st_mtime_ns))      # 更新日時が進まない共有
        self.assertEqual(os.path.getsize(db), st.st_size)
        self.assertEqual(self._pet(repo), "0.012")             # 間隔の中は今までどおり
        self.assertEqual(self._pet(repo, check_now=True), "7.77")

    def test_WAL_にだけ書かれた変更も拾う(self):
        db = os.path.join(self.tmp, "c.sqlite3")
        make_db(db, wal=True)
        repo = MaterialRepository("", "", prefer_access=False, refresh_sec=0, db_path=db)
        self.assertEqual(self._pet(repo), "0.012")
        w = sqlite3.connect(db)                                 # DB の道具を開いたまま
        self.addCleanup(w.close)
        w.execute('UPDATE "資材重量" SET 単位質量=? WHERE 梱包資材名=?', ("5.55", "PETバンド"))
        w.commit()
        self.assertEqual(self._pet(repo, check_now=True), "5.55")
        # 間隔 0 の stamp の比べ方でも(本体と -wal を見る)
        w.execute('UPDATE "資材重量" SET 単位質量=? WHERE 梱包資材名=?', ("5.56", "PETバンド"))
        w.commit()
        self.assertEqual(self._pet(repo), "5.56")

    def test_CSV_も計算のときは毎回読み直す(self):
        path = os.path.join(self.tmp, "e.csv")
        shutil.copyfile(SAMPLE, path)
        repo = MaterialRepository("", path, prefer_access=False, refresh_sec=300)
        self.assertEqual(self._pet(repo), "0.012")
        st = os.stat(path)
        with open(path, encoding="utf-8-sig") as f:
            text = f.read()
        with open(path, "w", encoding="utf-8-sig") as f:
            f.write(text.replace("PETバンド,0.012", "PETバンド,0.018"))
        os.utime(path, ns=(st.st_atime_ns, st.st_mtime_ns))
        self.assertEqual(self._pet(repo, check_now=True), "0.018")


# ---------------------------------------------------------------- 7
class FallbackWarnsTest(_Base):
    def test_前の内容で計算したら注意にして理由を出す(self):
        self.calc()
        self.assertEqual(self.last.level, "info")
        os.rename(self.db, self.db + ".x")                     # 共有に届かなくなった
        self.calc()
        self.assertEqual(self.last.level, "warn")
        self.assertIn("資材マスタ(SQLite)に届かないため", self.last.message)
        self.assertIn("前の内容", self.last.message)
        html = self.pages.tare()
        self.assertIn("資材重量:", html)
        self.assertIn("古い内容", html)
        html = self.pages.label_print([5, 6])
        self.assertIn("資材重量:", html)

    def test_CSV_に落ちて計算したら注意にして理由を出す(self):
        csv_path = os.path.join(self.tmp, "資材重量.csv")
        shutil.copyfile(SAMPLE, csv_path)
        self.repo.reconfigure("", csv_path, False, 60, 300, db_path=self.db + ".none")
        self.calc()
        self.assertEqual(self.last.level, "warn")
        self.assertIn("資材マスタ(SQLite)に届かないため", self.last.message)
        self.assertIn("CSV", self.last.message)

    def test_資材重量の素性は画面だけに出す(self):
        self.calc()
        html = self.pages.tare(print_view=True)
        i = html.index("資材重量:")
        self.assertIn("no-print", html[max(0, i - 300):i])


# ---------------------------------------------------------------- 8
class BreakdownTest(_Base):
    def test_計算内容に計算したときの時点と_いまと違う資材を出す(self):
        self.calc()
        html = self.pages.breakdown()
        self.assertIn("計算したときの資材マスタ", html)
        self.assertNotIn("いまの資材マスタと違う資材", html)
        set_mass(self.db, "PETバンド", "0.013")
        html = self.pages.breakdown()
        self.assertIn("いまの資材マスタと違う資材", html)
        self.assertIn("PETバンド", html)

    def test_資材マスタの一覧はいまの内容を出す(self):
        from modules.packing_pena_label.app.routes.api import ApiRoutes
        api = ApiRoutes(self.wf, self.cfg, time.time())
        self.repo.load()
        set_mass(self.db, "PETバンド", "0.013")
        rows = {r["name"]: r["mass"] for r in api._materials({})["rows"]}
        self.assertEqual(rows["PETバンド"], "0.013")


# ---------------------------------------------------------------- 9
class SizeMasterJsonTest(_Base):
    def test_JSONを外で直したら再起動しなくても読み直す_上書きも掛け直す(self):
        sql(self.db, 'INSERT INTO "ラベル台紙一覧" VALUES (5, "", "OVR555")')
        self.wf.refresh_label_sheets(force=True)
        self.assertEqual(self.master.get(5).kataban, "OVR555")
        with open(self.json, encoding="utf-8") as f:
            data = json.load(f)
        for row in data["sizes"]:
            if 1 in row["obIdx"]:
                row["kataban"] = "JSON-NEW"
        with open(self.json, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False)
        self.wf.label_view(1)                   # ラベル台紙を開いた
        self.assertEqual(self.master.get(1).kataban, "JSON-NEW")
        self.assertEqual(self.master.get(5).kataban, "OVR555", "上書きが外れた")


if __name__ == "__main__":
    unittest.main()
