"""管理者以外停止モード(統合 1.2.67。`common/remote_stop.py`・`common/stop_keep.py`)

現場の決め(業務ツール python-web-tools VER4.8.129〜130 と同じ):
    お願いを出せる      管理者パスワード かつ この端末のアクセス権限に Administrator
    止める範囲          お願いを出した PC 以外の、起動中の PC
    保存していない入力・ 帯で知らせて利用者に任せる。何も残っていなければ自動で止まる。5分で強制
    処理中               (保存していない入力は控えに残し、次の起動で知らせる・資材計算の作業は戻す)
    更新中               お願いを出した PC 以外は起動させない
    ランチャー           同時には使わない(お願いが出ているあいだ stop.bat は断る)
"""
from __future__ import annotations

import io
import json
import os
import tempfile
import unittest
from contextlib import redirect_stdout
from pathlib import Path
from unittest import mock

from common import remote_stop, stop_keep


class _Base(unittest.TestCase):
    PC = "PC-ADMIN"
    ADMIN_ROWS = [{"ログインID": "", "PC名": "PC-ADMIN", "権限": "L-1、Administrator", "有効": -1}]

    def setUp(self) -> None:
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        self.share = Path(tmp.name) / "共有"
        self.share.mkdir()
        self.local = Path(tmp.name) / "ローカル"
        for patcher in (
                mock.patch.object(remote_stop, "master_dir", lambda: self.share),
                mock.patch.object(remote_stop, "READ_RETRY_SEC", ()),
                mock.patch.dict(os.environ, {"COMPUTERNAME": self.PC, "USERNAME": "kanri",
                                             "COIL_PACKING_TOOLS_LOCAL_DIR": str(self.local)})):
            patcher.start()
            self.addCleanup(patcher.stop)
        remote_stop.reset()
        self.addCleanup(remote_stop.reset)

    def allow(self, password_ok: bool = True, rows=None):
        rows = self.ADMIN_ROWS if rows is None else rows
        return (mock.patch.object(remote_stop, "_password_ok", return_value=password_ok),
                mock.patch.object(remote_stop, "read_access_rows", return_value=rows))

    def issue(self, **kw):
        a, b = self.allow(**kw)
        with a, b:
            return remote_stop.issue("pw", "更新のため")

    def clear(self):
        a, b = self.allow()
        with a, b:
            return remote_stop.clear("pw")

    def as_pc(self, name: str):
        return mock.patch.dict(os.environ, {"COMPUTERNAME": name})


class AdministratorTests(unittest.TestCase):
    def test_権限の欄にAdministratorがあるか_並べても幅や大小が違ってもよい(self) -> None:
        self.assertTrue(remote_stop.has_administrator(["L-1、Administrator"]))
        self.assertTrue(remote_stop.has_administrator(["ＡＤＭＩＮＩＳＴＲＡＴＯＲ"]))
        self.assertTrue(remote_stop.has_administrator(["mode:field administrator"]))
        self.assertFalse(remote_stop.has_administrator(["作業長", "Administrators", None]))

    def test_行が効くのは_条件が合い有効なときだけ(self) -> None:
        match = remote_stop.rule_matches
        self.assertTrue(match({"ログインID": "", "PC名": "pc-admin", "有効": 1}, "PC-ADMIN", "x"))
        self.assertTrue(match({"ログインID": "KANRI", "PC名": "", "有効": "True"}, "PC-9", "kanri"))
        self.assertFalse(match({"ログインID": "kanri", "PC名": "PC-1", "有効": 1}, "PC-2", "kanri"))
        self.assertFalse(match({"ログインID": "", "PC名": "PC-ADMIN", "有効": 0}, "PC-ADMIN", "x"))
        self.assertFalse(match({"ログインID": "", "PC名": "PC-ADMIN", "有効": "0"}, "PC-ADMIN", "x"))
        self.assertFalse(match({"ログインID": "", "PC名": "", "有効": 1}, "PC-ADMIN", "x"),
                         "条件の無い行は誰にも効かせない")

    def test_アクセス権限の表が無ければ空_読めなければ読めないと言う(self) -> None:
        from modules.packing_material_calculation.coil_tool import data_sync, source_db
        with mock.patch.object(data_sync, "find_master_db", return_value=Path("x.sqlite3")), \
                mock.patch.object(remote_stop, "READ_RETRY_SEC", ()):
            with mock.patch.object(source_db, "read_query",
                                   side_effect=source_db.SourceError("x を読めませんでした: no such table: アクセス権限")):
                self.assertEqual(remote_stop.read_access_rows(), [])
            with mock.patch.object(source_db, "read_query",
                                   side_effect=source_db.SourceError("database is locked")):
                with self.assertRaises(remote_stop.RulesUnreadable):
                    remote_stop.read_access_rows()
        with mock.patch.object(data_sync, "find_master_db", return_value=None):
            with self.assertRaises(remote_stop.RulesUnreadable):
                remote_stop.read_access_rows()

    def test_本物のマスタからアクセス権限を読む(self) -> None:
        import sqlite3
        from modules.packing_material_calculation.coil_tool import data_sync
        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "梱包資材マスタ.sqlite3"
            conn = sqlite3.connect(path)
            conn.execute('CREATE TABLE "アクセス権限" ("管理番号" INTEGER, "ログインID" TEXT, "PC名" TEXT,'
                         ' "権限" TEXT, "有効" INTEGER, "備考" TEXT)')
            conn.execute('INSERT INTO "アクセス権限" VALUES (1, "", "PC-ADMIN", "Administrator", -1, "")')
            conn.commit()
            conn.close()
            with mock.patch.object(data_sync, "find_master_db", return_value=path):
                self.assertTrue(remote_stop.is_administrator("pc-admin", "who"))
                self.assertFalse(remote_stop.is_administrator("PC-LINE1", "who"))


class IssueTests(_Base):
    def test_パスワードとAdministratorの両方が要る(self) -> None:
        self.assertEqual(self.issue(password_ok=False).reason, remote_stop.REFUSE_PASSWORD)
        self.assertEqual(self.issue(rows=[]).reason, remote_stop.REFUSE_NOT_ADMIN)
        with self.as_pc("PC-LINE1"):
            self.assertEqual(self.issue().reason, remote_stop.REFUSE_NOT_ADMIN, "ほかの PC の行は効かない")
        self.assertFalse(remote_stop.request_path().exists())

    def test_アクセス権限を読めなければ出さない_Administratorでないとは言わない(self) -> None:
        with mock.patch.object(remote_stop, "_password_ok", return_value=True), \
                mock.patch.object(remote_stop, "read_access_rows",
                                  side_effect=remote_stop.RulesUnreadable("locked")):
            result = remote_stop.issue("pw")
        self.assertEqual(result.reason, remote_stop.REFUSE_RULES)
        self.assertIn("読めない", result.message)
        self.assertFalse(remote_stop.request_path().exists())

    def test_お願いを出す_出ているあいだは重ねて出さない_解除する(self) -> None:
        self.assertTrue(self.issue().ok)
        req = remote_stop.read_request()
        self.assertEqual((req["by_pc"], req["reason"], req["grace_sec"]), (self.PC, "更新のため", 300))
        self.assertIn("コイル梱包ツール", str(remote_stop.request_path()), "業務ツールのお願いと分ける")
        self.assertEqual(self.issue().reason, remote_stop.REFUSE_ALREADY)
        self.assertTrue(self.clear().ok)
        self.assertEqual(self.clear().reason, remote_stop.REFUSE_NOT_ACTIVE)
        self.assertIsNone(remote_stop.read_request())

    def test_読めないお願いは上書きしない(self) -> None:
        remote_stop.folder().mkdir(parents=True)
        remote_stop.request_path().write_text("{壊れている", encoding="utf-8")
        self.assertEqual(self.issue().reason, remote_stop.REFUSE_UNREADABLE)
        self.assertEqual(self.clear().reason, remote_stop.REFUSE_UNREADABLE)
        self.assertEqual(remote_stop.request_path().read_text(encoding="utf-8"), "{壊れている")

    def test_共有のフォルダに届かなければ出せない(self) -> None:
        with mock.patch.object(remote_stop, "master_dir", lambda: self.share / "無い"):
            self.assertEqual(self.issue().reason, remote_stop.REFUSE_NO_FOLDER)

    def test_業務ツールのお願いでは止まらない(self) -> None:
        other = self.share / remote_stop.FOLDER / remote_stop.REQUEST_NAME
        other.parent.mkdir(parents=True)
        other.write_text(json.dumps({"active": True, "id": "x", "by_pc": "PC-9"}), encoding="utf-8")
        with self.as_pc("PC-LINE1"):
            self.assertIsNone(remote_stop.startup_block())


class StartupTests(_Base):
    def test_更新中はお願いを出したPC以外で起動しない(self) -> None:
        self.issue()
        with self.as_pc("PC-LINE1"):
            why = remote_stop.startup_block()
        self.assertIn("更新が終わるまで起動できません", why)
        self.assertIn("別の名前に変えると解除できます", why)
        with self.as_pc("pc-admin"):
            self.assertIsNone(remote_stop.startup_block(), "お願いを出した PC は起動できる")

    def test_読めなければ起動は止めない(self) -> None:
        remote_stop.folder().mkdir(parents=True)
        remote_stop.request_path().write_text("{壊れている", encoding="utf-8")
        with self.as_pc("PC-LINE1"):
            self.assertIsNone(remote_stop.startup_block())

    def test_解除したら起動できる(self) -> None:
        self.issue()
        self.clear()
        with self.as_pc("PC-LINE1"):
            self.assertIsNone(remote_stop.startup_block())

    def test_起動の入口で_理由を出してしばらくして終わる(self) -> None:
        import start_app
        self.issue()
        srv = mock.Mock()
        with self.as_pc("PC-LINE1"), \
                mock.patch.object(remote_stop, "BLOCK_SHOW_SEC", 0.01), \
                mock.patch.object(start_app, "_hold_until_stopped") as hold:
            self.assertTrue(start_app._blocked_by_remote_stop(srv, mock.Mock()))
        srv.boot.mark_error.assert_called_once()
        self.assertIn("管理者が更新のため", srv.boot.mark_error.call_args.args[0])
        hold.assert_called_once()

    def test_お願いが無ければ起動の入口は素通り(self) -> None:
        import start_app
        srv = mock.Mock()
        self.assertFalse(start_app._blocked_by_remote_stop(srv, mock.Mock()))
        srv.boot.mark_error.assert_not_called()

    def test_起動の2つの入口とも_本体を組み立てる前に見る(self) -> None:
        import inspect

        import start_app
        for func in (start_app.start, start_app.start_bridge):
            text = inspect.getsource(func)
            self.assertLess(text.index("_blocked_by_remote_stop"), text.index("srv.build()"), func.__name__)


class StopBatTests(_Base):
    def test_お願いが出ているあいだstop_batは断る(self) -> None:
        import process_manager
        self.issue()
        out = io.StringIO()
        with redirect_stdout(out), mock.patch.object(process_manager, "stop") as stop:
            code = process_manager.main(["--force"])
        self.assertEqual(code, process_manager.EXIT_FAILED)
        self.assertTrue(out.getvalue().startswith("止められませんでした: 管理者以外停止モード中"),
                        out.getvalue())
        stop.assert_not_called()

    def test_お願いが無い_読めないならいつもどおり(self) -> None:
        self.assertIsNone(remote_stop.stop_bat_refusal())
        remote_stop.folder().mkdir(parents=True)
        remote_stop.request_path().write_text("{壊れている", encoding="utf-8")
        self.assertIsNone(remote_stop.stop_bat_refusal(), "読めなければ止める道を閉じない")


class ScreenReportTests(_Base):
    def test_開いている画面の答えを待ち_答えれば名前を返す(self) -> None:
        remote_stop.page_seen("shell-a", now=100.0)
        self.assertEqual(remote_stop.screens_left(now=101.0), [remote_stop.WAITING_SCREEN])
        remote_stop.report("shell-a", ["梱包明細: 重量"], [{"screen": "梱包明細", "label": "重量1", "value": "12.3"}],
                           now=102.0)
        self.assertEqual(remote_stop.screens_left(now=103.0), ["梱包明細: 重量"])
        remote_stop.report("shell-a", [], [], now=104.0)
        self.assertEqual(remote_stop.screens_left(now=105.0), [])

    def test_答えが古くなれば待つ_閉じた画面は待たない(self) -> None:
        remote_stop.report("shell-a", [], [], now=100.0)
        remote_stop.page_seen("shell-a", now=100.0 + remote_stop.REPORT_FRESH_SEC + 5)
        self.assertEqual(remote_stop.screens_left(now=100.0 + remote_stop.REPORT_FRESH_SEC + 6),
                         [remote_stop.WAITING_SCREEN])
        remote_stop.page_gone("shell-a")
        self.assertEqual(remote_stop.screens_left(now=200.0), [])

    def test_長く問い合わせてこない画面は居ないとみなす(self) -> None:
        remote_stop.page_seen("shell-a", now=100.0)
        self.assertEqual(remote_stop.screens_left(now=100.0 + remote_stop.PAGE_ALIVE_SEC + 1), [])


class WatcherTests(_Base):
    def watcher(self, *, unsaved=None, busy=None, saved=None):
        self.now = 1000.0
        self.stopped = []
        self.unsaved = list(unsaved or [])
        self.busy = list(busy or [])
        return remote_stop.Watcher(lambda: self.stopped.append(True), lambda: self.unsaved,
                                   lambda: self.busy, before_force=(lambda: saved or []),
                                   clock=lambda: self.now)

    def answer(self, w) -> dict:
        return json.loads(w.presence_path().read_text(encoding="utf-8"))

    def test_何も残っていなければ帯を見せてから止まる(self) -> None:
        self.issue()
        with self.as_pc("PC-LINE1"):
            w = self.watcher()
            w.tick()
            self.assertEqual(w.state, remote_stop.STATE_NOTICE)
            self.assertIsNotNone(w.screen_state())
            self.assertEqual(self.stopped, [])
            self.now += remote_stop.NOTICE_SEC
            w.tick()
        self.assertEqual(self.stopped, [True])
        self.assertEqual(self.answer(w)["state"], remote_stop.STATE_STOPPED)
        self.assertTrue(w.screen_state()["stopping"])
        w.leave()
        self.assertTrue(w.presence_path().exists(), "止まった答えは管理者が見るので残す")

    def test_保存していない入力があれば待ち_保存したら止まる(self) -> None:
        self.issue()
        with self.as_pc("PC-LINE1"):
            w = self.watcher(unsaved=["梱包明細: 重量"])
            self.now += remote_stop.NOTICE_SEC + 10
            w.tick()
            self.assertEqual((w.state, w.detail), (remote_stop.STATE_WAITING, "梱包明細: 重量"))
            self.assertEqual(w.screen_state()["left"], ["梱包明細: 重量"])
            self.assertEqual(self.stopped, [])
            self.unsaved.clear()
            w.tick()
        self.assertEqual(self.stopped, [True])

    def test_処理中も待つ(self) -> None:
        self.issue()
        with self.as_pc("PC-LINE1"):
            w = self.watcher(busy=["資材計算: 取り込み"])
            self.now += remote_stop.NOTICE_SEC + 10
            w.tick()
        self.assertEqual((w.state, self.stopped), (remote_stop.STATE_WAITING, []))

    def test_5分たっても止まらなければ控えてから強制で止める(self) -> None:
        self.issue()
        with self.as_pc("PC-LINE1"):
            w = self.watcher(unsaved=["梱包明細: 重量"], busy=["取り込み"], saved=["梱包明細: 重量"])
            w.tick()
            self.now += remote_stop.GRACE_SEC - 1
            w.tick()
            self.assertEqual(self.stopped, [])
            self.now += 1
            w.tick()
        self.assertEqual(self.stopped, [True])
        answer = self.answer(w)
        self.assertEqual(answer["state"], remote_stop.STATE_FORCED)
        self.assertIn("5分たっても", answer["detail"])
        self.assertIn("控えに残しました", answer["detail"])

    def test_控えに失敗しても期限には止める(self) -> None:
        self.issue()
        with self.as_pc("PC-LINE1"):
            w = self.watcher(unsaved=["梱包明細: 重量"])
            w._before_force = mock.Mock(side_effect=OSError("disk"))
            w.tick()
            self.now += remote_stop.GRACE_SEC
            w.tick()
        self.assertEqual(self.stopped, [True])

    def test_お願いを出したPCは止めない(self) -> None:
        self.issue()
        w = self.watcher()
        w.tick()
        self.now += remote_stop.GRACE_SEC + 1
        w.tick()
        self.assertEqual((self.stopped, w.state), ([], remote_stop.STATE_RUNNING))
        self.assertIsNone(w.screen_state())

    def test_取り下げられたら止めずに続ける(self) -> None:
        self.issue()
        with self.as_pc("PC-LINE1"):
            w = self.watcher(unsaved=["梱包明細: 重量"])
            w.tick()
        self.assertTrue(self.clear().ok)        # 管理者の PC で解除
        with self.as_pc("PC-LINE1"):
            w.tick()
        self.assertEqual((self.stopped, w.state), ([], remote_stop.STATE_RUNNING))
        self.assertTrue(any("取り下げ" in n for n in remote_stop.notes()))

    def test_お願いを読めなくなっても取り下げとは見ない(self) -> None:
        self.issue()
        with self.as_pc("PC-LINE1"):
            w = self.watcher(unsaved=["梱包明細: 重量"])
            w.tick()
            remote_stop.request_path().write_text("{書いている途中", encoding="utf-8")
            w.tick()
            self.assertEqual(w.state, remote_stop.STATE_WAITING)
            self.now += remote_stop.GRACE_SEC
            w.tick()
        self.assertEqual(self.stopped, [True])

    def test_お願いの最中に利用者が自分で終えたら_止まりましたを残す(self) -> None:
        self.issue()
        with self.as_pc("PC-LINE1"):
            w = self.watcher(unsaved=["梱包明細: 重量"])
            w.tick()
            w.leave()
        answer = self.answer(w)
        self.assertEqual(answer["state"], remote_stop.STATE_STOPPED)
        self.assertIn("利用者が終了しました", answer["detail"])

    def test_ふだんの終わりなら知らせを消す(self) -> None:
        w = self.watcher()
        w.tick()
        self.assertTrue(w.presence_path().exists())
        w.leave()
        self.assertFalse(w.presence_path().exists())

    def test_共有に書けないあいだは記録をためない(self) -> None:
        w = self.watcher()
        with mock.patch.object(remote_stop, "write_json", return_value=False) as write, \
                self.assertLogs(remote_stop.log.name, "WARNING") as logs:
            for _ in range(3):
                self.now += remote_stop.BEAT_SEC
                w.tick()
        self.assertEqual(write.call_count, 3)
        self.assertEqual(len([r for r in logs.records if "書けません" in r.getMessage()]), 1)

    def test_止める処理は錠の外で呼ぶ(self) -> None:
        """止めると終わりの後始末(leave)が同じ錠を取りに来る。錠を持ったまま止めると待ち合う。"""
        self.issue()
        with self.as_pc("PC-LINE1"):
            w = self.watcher()
            w._stop = lambda: self.stopped.append(w._lock.locked())
            w.tick()
            self.now += remote_stop.NOTICE_SEC
            w.tick()
        self.assertEqual(self.stopped, [False])

    def test_管理者の一覧に各PCの様子が出る(self) -> None:
        self.issue()
        with self.as_pc("PC-LINE1"):
            w = self.watcher(unsaved=["梱包明細: 重量"])
            self.now = __import__("time").time()
            w.tick()
            self.now += remote_stop.NOTICE_SEC + 1
            w.tick()
        view = remote_stop.overview().to_dict()
        self.assertTrue(view["available"])
        self.assertIn("PC-ADMIN", view["request_text"])
        rows = [p for p in view["pcs"] if p["pc"] == "PC-LINE1"]
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows[0]["label"], remote_stop.STATE_LABELS[remote_stop.STATE_WAITING])
        self.assertEqual(rows[0]["detail"], "梱包明細: 重量")


class StopKeepTests(_Base):
    def setUp(self) -> None:
        super().setUp()
        from modules.packing_material_calculation.app import session
        self.session = session
        session.reset()
        self.addCleanup(session.reset)

    def test_何も無ければ控えを書かない(self) -> None:
        self.assertEqual(stop_keep.save_before_force(), [])
        self.assertEqual(list(stop_keep.keep_dir().glob("*.json")) if stop_keep.keep_dir().exists() else [], [])

    def test_画面の入力を控えて_次の起動で一度だけ知らせる(self) -> None:
        remote_stop.report("shell-a", ["梱包明細: 重量(空の丈があるので、まだ送っていない)"],
                           [{"screen": "梱包明細", "label": "重量1", "value": "12.3"},
                            {"screen": "梱包明細", "label": "重量2", "value": ""}])
        saved = stop_keep.save_before_force()
        self.assertEqual(saved, ["梱包明細: 重量(空の丈があるので、まだ送っていない)"])
        notes = stop_keep.restore()
        self.assertEqual(len(notes), 1)
        self.assertIn("重量1「12.3」", notes[0])
        self.assertIn("打ち直して", notes[0])
        self.assertEqual(stop_keep.restore(), [], "二度は知らせない")
        self.assertEqual(len(list((stop_keep.keep_dir() / stop_keep.DONE_NAME).glob("*.json"))), 1,
                         "知らせた控えは消さずに移す")

    def test_待っている印は控えに入れない(self) -> None:
        remote_stop.page_seen("shell-a")
        self.assertEqual(remote_stop.screens_left(), [remote_stop.WAITING_SCREEN])
        self.assertEqual(stop_keep.save_before_force(), [])

    def test_資材計算の作業を控えて_次の起動で戻す(self) -> None:
        state = self.session.get()
        state.LOT = "A12B3C4"
        state.積数 = "3"
        state.order.受注番号 = "J-1"
        state.候補 = ["J-1", "J-2"]
        state.flags = {"no_single"}
        saved = stop_keep.save_before_force()
        self.assertEqual(saved, ["資材計算の作業(LOT A12B3C4・受注 J-1)"])
        self.session.reset()
        notes = stop_keep.restore()
        self.assertIn("戻しました", notes[0])
        back = self.session.get()
        self.assertEqual((back.LOT, back.積数, back.order.受注番号, back.候補, back.flags),
                         ("A12B3C4", "3", "J-1", ["J-1", "J-2"], {"no_single"}))

    def test_いま別の作業があれば戻さない(self) -> None:
        self.session.get().LOT = "A12B3C4"
        stop_keep.save_before_force()
        self.session.reset()
        self.session.get().LOT = "Z99Z9Z9"
        notes = stop_keep.restore()
        self.assertIn("戻しませんでした", notes[0])
        self.assertEqual(self.session.get().LOT, "Z99Z9Z9")

    def test_知らない欄は捨てて読む(self) -> None:
        state = stop_keep.build_material_state({"LOT": "A12B3C4", "未来の欄": "x", "order": {"受注番号": "J", "未来": 1}})
        self.assertEqual((state.LOT, state.order.受注番号), ("A12B3C4", "J"))

    def test_壊れた控えは知らせず残す(self) -> None:
        stop_keep.keep_dir().mkdir(parents=True)
        bad = stop_keep.keep_dir() / f"{stop_keep.PREFIX}20260101-000000.json"
        bad.write_text("{壊れ", encoding="utf-8")
        self.assertEqual(stop_keep.restore(), [])
        self.assertTrue(bad.exists())


class RouteTests(_Base):
    def setUp(self) -> None:
        super().setUp()
        import app as app_module
        self.app = app_module.create_app(token="tok", modules=[])
        self.client = self.app.test_client()
        self.h = {"X-Tool-Token": "tok"}

    def test_合言葉が無ければ答えない(self) -> None:
        self.assertEqual(self.client.get("/api/remote-stop/state").status_code, 403)
        self.assertEqual(self.client.post("/api/remote-stop/issue", json={}).status_code, 403)

    def test_出す_一覧_解除(self) -> None:
        a, b = self.allow()
        with a, b:
            res = self.client.post("/api/remote-stop/issue", json={"password": "pw", "reason": "VER更新"},
                                   headers=self.h)
            self.assertEqual(res.status_code, 200, res.get_json())
            view = self.client.get("/api/remote-stop/overview", headers=self.h).get_json()
            self.assertIn("VER更新", view["request_text"])
            res = self.client.post("/api/remote-stop/clear", json={"password": "pw"}, headers=self.h)
            self.assertTrue(res.get_json()["ok"])
        a, b = self.allow(password_ok=False)
        with a, b:
            res = self.client.post("/api/remote-stop/issue", json={"password": "x"}, headers=self.h)
        self.assertEqual((res.status_code, res.get_json()["reason"]), (400, remote_stop.REFUSE_PASSWORD))

    def test_画面の答えと知らせ(self) -> None:
        state = self.client.get("/api/remote-stop/state?page=shell-a", headers=self.h).get_json()
        self.assertEqual((state["request"], state["notes"]), (None, []))
        self.client.post("/api/remote-stop/report", headers=self.h,
                         json={"page": "shell-a", "labels": ["梱包明細: 重量"], "fields": "壊れた形"})
        self.assertEqual(remote_stop.screens_left(), ["梱包明細: 重量"])
        remote_stop.user_note("前回の控えを知らせます")
        self.assertEqual(self.client.get("/api/remote-stop/state", headers=self.h).get_json()["notes"],
                         ["前回の控えを知らせます"])
        self.client.post("/api/remote-stop/notes-seen", json={"notes": ["前回の控えを知らせます"]}, headers=self.h)
        self.assertEqual(remote_stop.notes(), [])
        self.client.post("/api/alive", json={"client": "shell-a", "leaving": True})
        self.assertEqual(remote_stop.screens_left(), [], "閉じた画面は待たない")


if __name__ == "__main__":
    unittest.main()
