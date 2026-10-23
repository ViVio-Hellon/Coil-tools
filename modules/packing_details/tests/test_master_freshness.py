"""マスタ・台帳を直したら前の中身で刷らない・動かない(統合 1.2.49・現場の指摘)

    マスタを直しても前の中身のまま動く・表示する事がないかのチェックをしてください
    例：資材重量 変更後 前の重量のまま計算

梱包明細は資材重量を使わない(紙の重量は画面で打ったコイル重量だけ)。点検で見つかった
「前の中身のまま刷る・動く」を、ここで確かめる。

1. 開いたままの紙面のタブが、作り直し・取り消し・ロットの切り替えのあとも古い中身で刷れた
2. 途中経過から戻した丈数・前工程実績数が台帳と違っても黙っていた(案内文も実際と違った)
3. 動いている間に取り込み元が変わっても「新しい台帳がある」と言わなかった
4. 自動取り込みが「変わっていない」と取り違えた(WAL・古い更新時刻での置き換え・読んでいる間の書き換え)
5. 一部失敗した取り込みのあと「開いているロットは古いまま」が出なかった
8. マスタの右上の文字を空にしたとき、控えの前の値で黙って刷った
10. 「中身を見る」の行数の控えが -wal の書き込みを見ていなかった
11. 元の表が空のとき「0件」とだけ出て、手元が前の中身のままだと分からなかった
"""
from __future__ import annotations

import json
import os
import re
import shutil
import sqlite3
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

from common import source_copy
from modules.packing_details.meisai import (data_sync, history_repo, master_browse, qa_mark,
                                           session, shared_settings)

from . import _db
from .test_web import HEADERS, TOKEN, _client, _post, _seed


def _tmpdir(case: unittest.TestCase) -> Path:
    path = Path(tempfile.mkdtemp(prefix="meisai_fresh_"))
    case.addCleanup(shutil.rmtree, path, True)
    return path


def _lot_source(path: Path, lots, created_at="2026-09-14T09:00:26") -> Path:
    return _db.make_source(path, _db.LOT_COLUMNS, [_db.lot_row(n) for n in lots],
                           created_at=created_at)


def _lots(conn) -> list[str]:
    return sorted(r[0] for r in conn.execute("SELECT ロット番号 FROM 仕掛ロット"))


# ======================================================================
# 4. 取り込み元が変わったかの見分け
# ======================================================================
class NeedsImportTest(unittest.TestCase):
    def setUp(self):
        self.conn = _db.case_db(self)
        self.dir = _tmpdir(self)
        source_copy.forget()
        self.addCleanup(source_copy.forget)

    def test_WALで_wal_にだけ書かれても気づく(self):
        """別の接続が書いて checkpoint していない(本体の大きさも更新時刻も変わらない)。"""
        path = _lot_source(self.dir / "SIKALOT.sqlite3", ["AAA0001"])
        writer = sqlite3.connect(str(path))
        self.addCleanup(writer.close)
        writer.execute("PRAGMA journal_mode=WAL")
        writer.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        data_sync.import_table(self.conn, "仕掛ロット", path)
        self.assertFalse(data_sync.needs_import(self.conn, path))
        main = path.stat()
        writer.execute('INSERT INTO "仕掛" ("ﾛｯﾄ番号") VALUES (?)', ("BBB0002",))
        writer.commit()
        self.assertEqual((path.stat().st_size, path.stat().st_mtime_ns),
                         (main.st_size, main.st_mtime_ns), "本体は変わっていない(前提)")
        self.assertTrue(data_sync.needs_import(self.conn, path))

    def test_更新時刻の古いファイルで置き換えても気づく(self):
        backup = _lot_source(self.dir / "bk.sqlite3", ["OLD0001", "OLD0002"])
        old = backup.stat()
        path = _lot_source(self.dir / "SIKALOT.sqlite3", ["NEW0001"])
        os.utime(path, ns=(old.st_atime_ns, old.st_mtime_ns + 10_000_000_000))
        data_sync.import_table(self.conn, "仕掛ロット", path)
        shutil.copy2(backup, path)                 # 更新時刻を保ったまま戻した
        self.assertLess(path.stat().st_mtime_ns, old.st_mtime_ns + 10_000_000_000)
        self.assertTrue(data_sync.needs_import(self.conn, path))

    def test_読み終えてから記録を書くまでに書き換わったら次に読み直す(self):
        path = _lot_source(self.dir / "SIKALOT.sqlite3", ["V1AAAA1"])
        original = data_sync.read_source

        def read_then_upstream_writes(table, src):
            out = original(table, src)
            _lot_source(Path(str(path) + ".new"), ["V2BBBB2"])
            os.replace(str(path) + ".new", path)   # 読み終えた直後に上流が書いた
            return out

        with mock.patch.object(data_sync, "read_source", read_then_upstream_writes):
            data_sync.import_table(self.conn, "仕掛ロット", path)
        self.assertEqual(_lots(self.conn), ["V1AAAA1"])
        self.assertTrue(data_sync.needs_import(self.conn, path),
                        "中身は前のものなのに「取り込んだ」と覚えていた")
        data_sync.import_all(self.conn, lot_dir=self.dir, konpo_dir=self.dir / "無い")
        self.assertEqual(_lots(self.conn), ["V2BBBB2"])

    def test_作成日時は読んだ中身と同じときのもの(self):
        path = _lot_source(self.dir / "SIKALOT.sqlite3", ["V1AAAA1"], created_at="2026-09-14T09:00:00")
        original = data_sync.read_source

        def read_then_upstream_writes(table, src):
            out = original(table, src)
            _lot_source(Path(str(path) + ".new"), ["V2BBBB2"], created_at="2026-09-15T09:00:00")
            os.replace(str(path) + ".new", path)
            return out

        with mock.patch.object(data_sync, "read_source", read_then_upstream_writes):
            data_sync.import_table(self.conn, "仕掛ロット", path)
        row = self.conn.execute("SELECT 元作成日時 FROM 取り込み記録 WHERE テーブル = '仕掛ロット'"
                                ).fetchone()
        self.assertEqual(row[0], "2026-09-14T09:00:00", "取り込んだ中身より新しい日時を名乗った")

    def test_前の版の記録_目印が空_は読み直す(self):
        path = _lot_source(self.dir / "SIKALOT.sqlite3", ["AAA0001"])
        data_sync.import_table(self.conn, "仕掛ロット", path)
        with self.conn:
            self.conn.execute("UPDATE 取り込み記録 SET 目印 = '', 更新時刻 = ?",
                              (time.time() + 3600,))
        self.assertTrue(data_sync.needs_import(self.conn, path))

    def test_変わっていなければ読み直さない_手動はいつも読む(self):
        path = _lot_source(self.dir / "SIKALOT.sqlite3", ["AAA0001"])
        data_sync.import_table(self.conn, "仕掛ロット", path)
        self.assertFalse(data_sync.needs_import(self.conn, path))
        with mock.patch.object(data_sync, "import_table", wraps=data_sync.import_table) as spy:
            data_sync.import_all(self.conn, lot_dir=self.dir, konpo_dir=self.dir / "無い")
            self.assertEqual(spy.call_count, 0)
            data_sync.import_all(self.conn, lot_dir=self.dir, konpo_dir=self.dir / "無い",
                                 force=True)
            self.assertEqual(spy.call_count, 1)


# ======================================================================
# 11. 元の表が空
# ======================================================================
class EmptySourceTest(unittest.TestCase):
    def setUp(self):
        self.conn = _db.case_db(self)
        self.dir = _tmpdir(self)

    def test_元が空なら前の中身のままと書き添える_動きは変えない(self):
        path = _lot_source(self.dir / "SIKALOT.sqlite3", ["AAA0001", "BBB0002"])
        data_sync.import_table(self.conn, "仕掛ロット", path)
        empty = _lot_source(self.dir / "空.sqlite3", [])
        os.replace(empty, path)
        result = data_sync.import_table(self.conn, "仕掛ロット", path)
        self.assertEqual(_lots(self.conn), ["AAA0001", "BBB0002"], "動きは変えない")
        self.assertEqual(result.imported, {"仕掛ロット": 0})
        self.assertTrue(any("前の中身のまま" in n and "2件" in n for n in result.notes),
                        result.notes)
        # この元は見た(「新しい台帳があります」を出し続けない)
        self.assertFalse(data_sync.needs_import(self.conn, path))


# ======================================================================
# 3. 新しい台帳がある
# ======================================================================
class NewLedgerTest(unittest.TestCase):
    def setUp(self):
        self.client, self.conn = _client(self)
        self.dir = _tmpdir(self)
        self.path = _lot_source(self.dir / "SIKALOT.sqlite3", ["AAA0001"])
        for name, value in (("find_sources", lambda *a, **k: {"仕掛ロット": self.path}),
                            ("find_second_lots", lambda: {})):
            patcher = mock.patch.object(data_sync, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)

    def check(self):
        res = self.client.get("/api/import/check", headers=HEADERS)
        self.assertEqual(res.status_code, 200, res.get_data(as_text=True))
        return res.get_json()

    def test_取り込んだあと元が変われば知らせる(self):
        self.assertEqual(self.check()["newer"], [], "一度も取り込んでいない表は言わない")
        data_sync.import_table(self.conn, "仕掛ロット", self.path)
        body = self.check()
        self.assertTrue(body["checked"])
        self.assertEqual(body["newer"], [])
        _lot_source(Path(str(self.path) + ".new"), ["AAA0001", "BBB0002"])
        os.replace(str(self.path) + ".new", self.path)
        body = self.check()
        self.assertEqual(body["newer"], ["仕掛ロット"])
        self.assertIn("新しい台帳があります", body["message"])
        # 取り込めば消える
        data_sync.import_all(self.conn, force=True)
        self.assertEqual(self.check()["newer"], [])

    def test_中身は読まない(self):
        data_sync.import_table(self.conn, "仕掛ロット", self.path)
        with mock.patch.object(data_sync.source_db, "read_table",
                               side_effect=AssertionError("読んだ")):
            self.check()

    def test_画面に帯と取り込むボタンがある(self):
        html = self.client.get("/meisai").get_data(as_text=True)
        self.assertIn('id="newLedgerBanner"', html)
        self.assertIn('id="btnNewLedgerImport"', html)
        js = Path(__file__).resolve().parents[1] / "app" / "static" / "js" / "meisai.js"
        text = js.read_text(encoding="utf-8")
        self.assertIn("/api/import/check", text)
        self.assertIn('$("btnNewLedgerImport").addEventListener("click", runImport)', text)


# ======================================================================
# 2・5. ロットの入れ直しと取り込みの案内
# ======================================================================
class ReopenTest(unittest.TestCase):
    def setUp(self):
        self.client, self.conn = _client(self)
        _seed(self.conn, tate="2", yoko="3")              # 丈2・前工程6
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        self.assertEqual(_post(self.client, "/api/strands", {}).status_code, 200)

    def ledger(self, tate, yoko):
        with self.conn:
            self.conn.execute("UPDATE 仕掛当工程 SET 当工程設計_縦割数 = ?, 当工程設計_横割数 = ?",
                              (tate, yoko))

    def test_途中経過の値で戻すとき台帳と違えば言う(self):
        self.ledger(3, 4)                                  # 取り込みで台帳が直った
        body = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"}).get_json()
        state = body["state"]
        self.assertEqual((state["jou_su"], state["zen_kotei"]), (2, 6), "動きは VBA のまま")
        self.assertIn("台帳は丈3・前工程12に変わっています", body["message"])
        self.assertIn("全クリアで台帳の値になります", body["message"])
        self.assertIn("台帳は丈3・前工程12", state["lot"]["ledger_note"])
        # 全クリアで台帳の値になり、断り書きも消える
        state = _post(self.client, "/api/clear-all", {"confirm": True}).get_json()["state"]
        self.assertEqual((state["jou_su"], state["zen_kotei"]), (3, 12))
        self.assertEqual(state["lot"]["ledger_note"], "")

    def test_台帳と同じなら何も言わない(self):
        body = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"}).get_json()
        self.assertNotIn("台帳は", body["message"])
        self.assertEqual(body["state"]["lot"]["ledger_note"], "")

    def import_with(self, result):
        with mock.patch.object(data_sync, "import_all", return_value=result):
            res = _post(self.client, "/api/import", {"force": True})
        self.assertEqual(res.status_code, 200)
        return res.get_json()

    def test_取り込みの案内は実際の動きのとおり(self):
        note = self.import_with(data_sync.ImportResult(imported={"仕掛当工程": 1}))["note"]
        self.assertIn("取り込み前の内容のまま", note)
        self.assertIn("途中経過の値に戻ります", note)
        self.assertNotIn("手で直した前工程実績数は入れ直し", note)
        # 途中経過が無ければ、手で直した数が台帳の値に戻る
        history_repo.delete_snapshot(self.conn, "L5160Z0")
        note = self.import_with(data_sync.ImportResult(imported={"仕掛当工程": 1}))["note"]
        self.assertIn("手で直した前工程実績数は台帳の値に戻ります", note)

    def test_一部失敗しても入れ替わったなら言う(self):
        body = self.import_with(data_sync.ImportResult(
            imported={"仕掛ロット": 5}, errors=["仕掛引当: 読めません"]))
        self.assertFalse(body["ok"])
        self.assertIn("取り込み前の内容のまま", body["note"])

    def test_何も入れ替わらなければ言わない(self):
        body = self.import_with(data_sync.ImportResult(errors=["取り込み元が見つかりません"]))
        self.assertEqual(body["note"], "")


# ======================================================================
# 1. 開いたままの紙面
# ======================================================================
class StaleSlipTest(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)
        _db.master_file(self)
        self.client, self.conn = _client(self)
        _seed(self.conn)
        _seed(self.conn, lot="L9999Z9")
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        _post(self.client, "/api/weights", {"weights": ["250", "1248"]})
        _post(self.client, "/api/strands", {})
        self.stack_and_output(["1-10", "2-8"])

    def stack_and_output(self, keys):
        _post(self.client, "/api/stack-max", {"value": str(len(keys))})
        for key in keys:
            _post(self.client, "/api/stack", {"key": key})
        res = _post(self.client, "/api/output", {"confirm": True})
        self.assertEqual(res.status_code, 200, res.get_data(as_text=True))

    def page(self, url):
        res = self.client.get(url)
        self.assertEqual(res.status_code, 200, res.get_data(as_text=True))
        return res.get_data(as_text=True)

    @staticmethod
    def slips(html):
        return json.loads(re.search(r"var SLIPS = (\[.*?\]);", html).group(1))

    def ask(self, slips):
        q = ",".join(slips)
        res = self.client.get(f"/report/slip-check?slips={q}&t={TOKEN}")
        self.assertEqual(res.status_code, 200)
        return res.get_json()

    def test_開いた紙面は身元を持ち_まだ使える(self):
        html = self.page(f"/report/L5160Z0/1?t={TOKEN}")
        slips = self.slips(html)
        self.assertEqual(len(slips), 1)
        self.assertTrue(self.ask(slips)["valid"])
        # 前に戻ったときの問い合わせ(右上の文字)にも一緒に入る
        body = self.client.get(f"/report/qa-mark?slips={slips[0]}&t={TOKEN}").get_json()
        self.assertTrue(body["slips"]["valid"])
        # 刷る直前(印刷する・Ctrl+P)にも聞く。古ければ紙面の代わりに断りを刷る
        self.assertIn('addEventListener("beforeprint", function () { slipOk(); })', html)
        self.assertIn("window.cptPrint ? window.cptPrint() : window.print()", html)
        self.assertIn("body.slip-stale .sheet{ display:none !important; }", html)

    def test_出力済み修正で作り直したら古い紙面は刷れない(self):
        slips = self.slips(self.page(f"/report/L5160Z0/1?t={TOKEN}"))
        self.assertEqual(_post(self.client, "/api/edit-output", {"no": 1}).status_code, 200)
        self.stack_and_output(["2-8", "1-10"])
        body = self.ask(slips)
        self.assertFalse(body["valid"])
        self.assertIn("作り直されました", body["title"])
        self.assertIn("L5160Z0-No1", body["detail"])
        # 開き直した紙面は使える
        self.assertTrue(self.ask(self.slips(self.page(f"/report/L5160Z0/1?t={TOKEN}")))["valid"])

    def test_全クリアで取り消したら古い紙面は刷れない(self):
        slips = self.slips(self.page(f"/report/L5160Z0?nos=1&t={TOKEN}"))
        _post(self.client, "/api/clear-all", {"confirm": True})
        body = self.ask(slips)
        self.assertFalse(body["valid"])
        self.assertIn("取り消されました", body["title"])

    def test_ロットを切り替えたら古い紙面は刷れない(self):
        slips = self.slips(self.page(f"/report/L5160Z0/1?t={TOKEN}"))
        _post(self.client, "/api/printed", {})
        _post(self.client, "/api/lot", {"lot_no": "L9999Z9", "confirm": True})
        body = self.ask(slips)
        self.assertFalse(body["valid"])
        self.assertIn("開き直してください", body["title"])

    def test_履歴から作り直した紙面は止めない(self):
        """過去の記録として刷ってよい。身元を持たない(聞かない)。"""
        from modules.packing_details.meisai import slip_history
        from datetime import date
        slip_history.send_pending(self.conn)
        today = date.today().isoformat()
        rows = _post(self.client, "/api/history/search",
                     {"from": today, "to": today, "lot_no": "L5160Z0"}).get_json()["rows"]
        _post(self.client, "/api/clear-all", {"confirm": True})
        html = self.page(f"/report/history?ids={rows[0]['送信ID']}&t={TOKEN}")
        self.assertEqual(self.slips(html), [])
        self.assertIn('onclick="window.print()">印刷する', html)


# ======================================================================
# 8. 右上の文字をマスタで空にした
# ======================================================================
class QaMarkNoRowTest(unittest.TestCase):
    def snap(self, master):
        return shared_settings.Snapshot(
            {"qa_mark": "OLD.QA", "qa_origin": shared_settings._compose(
                {"master": master, "json": {"qa_mark": "OLD.QA"}})["qa_origin"]},
            "shared", Path("."), master=master)

    def test_マスタの行が空で控えの前の値を使うときは紙面の画面に出す(self):
        snap = self.snap(shared_settings.Master(Path("m.sqlite3"), table=True, value=None))
        self.assertEqual(qa_mark.resolve(snap)[0], "OLD.QA", "刷る値は前のまま(紙は止めない)")
        text = qa_mark.notice(snap)
        self.assertIn("控えのJSONに残っている前の値", text)
        self.assertIn("ID=1", text)

    def test_マスタに値があれば何も出さない(self):
        snap = self.snap(shared_settings.Master(Path("m.sqlite3"), table=True, value="NEW.QA"))
        self.assertEqual(qa_mark.notice(snap), "")


# ======================================================================
# 10. 「中身を見る」の行数の控え
# ======================================================================
class BrowseCountsTest(unittest.TestCase):
    def setUp(self):
        _db.settings_file(self)
        self.path = _db.master_file(self)
        source_copy.forget()
        self.addCleanup(source_copy.forget)
        master_browse._counts_cache.clear()
        self.addCleanup(master_browse._counts_cache.clear)

    def test_walにだけ書かれた行も数え直す(self):
        writer = sqlite3.connect(str(self.path))
        self.addCleanup(writer.close)
        writer.execute("PRAGMA journal_mode=WAL")
        writer.execute("PRAGMA wal_checkpoint(TRUNCATE)")
        before = {t.table: t.rows for t in master_browse.browse("master").tables}
        main = self.path.stat()
        writer.execute('INSERT INTO "PalletMaster" DEFAULT VALUES')
        writer.commit()
        self.assertEqual((self.path.stat().st_size, self.path.stat().st_mtime_ns),
                         (main.st_size, main.st_mtime_ns), "本体は変わっていない(前提)")
        after = {t.table: t.rows for t in master_browse.browse("master").tables}
        self.assertEqual(after["PalletMaster"], before["PalletMaster"] + 1)

    def test_説明文は今の作り(self):
        doc = master_browse.browse.__doc__
        self.assertNotIn("**共有から直に読む。**", doc)
        self.assertIn("手元へ写してから", doc)


if __name__ == "__main__":
    unittest.main()
