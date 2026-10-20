"""構造点検(統合 1.2.26)で見つかった不具合の再現(統合部分)

直す前に、**直ったら通る試験**として置く(`expectedFailure`)。直したら印を外す。
印が付いたまま通ってしまったら(=気づかずに直った・別の所が変わった)、試験が落ちて知らせる。
"""
from __future__ import annotations

import os
import threading
import types
import unittest
from unittest import mock

import launch_guard
import process_manager
from common import app_config, close_ask
from tests.test_app import TOKEN, _make_app

H = {"X-Tool-Token": TOKEN}


def _info():
    return types.SimpleNamespace(pid=424242, port=50999, token="tok", url="http://127.0.0.1:50999/",
                                 started_text="", app_root="")


class ReplaceStaleTest(unittest.TestCase):
    """版の入れ替え: 古い版の**待ち受けが閉じただけ**では、終わったとみなさない(1.2.19 と同じ決まり)。"""

    def setUp(self):
        self.removed = []
        for name, value in {
            "ask_shutdown": lambda port, token, **k: {"ok": True},
            "probe_health": lambda port, timeout=0: None,          # 待ち受けはもう答えない
            "process_alive_state": lambda pid: True,               # でもプロセスは居る
            "is_process_alive": lambda pid: True,
            "remove_lock": lambda mode: self.removed.append(mode),
            "REPLACE_WAIT_SEC": 0.3,
        }.items():
            patcher = mock.patch.object(launch_guard, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)

    def test_プロセスが残っているうちは_ロックを消して立て直さない(self):
        result = launch_guard._replace_stale("main", _info(), "1.2.0", "統合 1.2.0 → 1.2.26")
        self.assertEqual(self.removed, [], "古いプロセスが居るのにロックを消した")
        self.assertFalse(result.should_start, "古いプロセスが居るのに立て直そうとした")

    def test_しばらくして終われば立て直す(self):
        polls = iter([True, False])           # 1回目はまだ居る、2回目で居なくなる
        with mock.patch.object(launch_guard, "process_alive_state", lambda pid: next(polls, False)):
            result = launch_guard._replace_stale("main", _info(), "1.2.0", "x")
        self.assertTrue(result.should_start)
        self.assertEqual(self.removed, ["main"])

    def test_有無を確かめられず待ち受けが閉じていれば立て直す(self):
        with mock.patch.object(launch_guard, "process_alive_state", lambda pid: None):
            result = launch_guard._replace_stale("main", _info(), "1.2.0", "x")
        self.assertTrue(result.should_start)

    def test_答えたpidのほうを待つ(self):
        seen = []
        with mock.patch.object(launch_guard, "process_alive_state",
                               lambda pid: seen.append(pid) or False):
            launch_guard._replace_stale("main", _info(), "1.2.0", "x", {"pid": 777})
        self.assertEqual(seen[0], 777)

    def test_途中の処理があれば止めずに合流(self):
        with mock.patch.object(launch_guard, "ask_shutdown",
                               lambda *a, **k: {"ok": False, "busy": True, "reason": "busy"}):
            result = launch_guard._replace_stale("main", _info(), "1.2.0", "x")
        self.assertFalse(result.should_start)
        self.assertEqual(self.removed, [])


class CheckExistingNoAnswerTest(unittest.TestCase):
    """起動時: ロックのプロセスは居るが答えない。**別物と決めつけて2つ目を立てない。**"""

    def setUp(self):
        self.removed = []
        info = types.SimpleNamespace(**vars(_info()))
        for name, value in {
            "read_lock": lambda mode: info,
            "is_process_alive": lambda pid: True,
            "probe_health": lambda port, timeout=0: None,
            "remove_lock": lambda mode, **k: self.removed.append(mode),
            "REPLACE_WAIT_SEC": 0.3,
        }.items():
            patcher = mock.patch.object(launch_guard, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)

    def test_待ち受けが開いていれば合流する(self):
        with mock.patch.object(launch_guard, "is_port_free", lambda port, host="": False):
            result = launch_guard.check_existing("main")
        self.assertFalse(result.should_start)
        self.assertEqual(self.removed, [])

    def test_待ち受けが閉じていれば終わるのを待ってから立てる(self):
        polls = iter([True, False])
        with mock.patch.object(launch_guard, "is_port_free", lambda port, host="": True), \
                mock.patch.object(launch_guard, "process_alive_state", lambda pid: next(polls, False)):
            result = launch_guard.check_existing("main")
        self.assertTrue(result.should_start)
        self.assertEqual(self.removed, ["main"])


class RemoveOwnLockTest(unittest.TestCase):
    """終わるプロセスは自分のロックだけ消す(入れ替えで新しい版が書いたロックを消さない)。"""

    def test_別のプロセスのロックは残す(self):
        unlinked = []
        fake = types.SimpleNamespace(unlink=lambda: unlinked.append(1))
        with mock.patch.object(launch_guard, "read_lock", lambda mode: types.SimpleNamespace(pid=111)), \
                mock.patch.object(launch_guard, "lock_path", lambda mode: fake):
            launch_guard.remove_lock("main", pid=222)
            self.assertEqual(unlinked, [])
            launch_guard.remove_lock("main", pid=111)
            self.assertEqual(unlinked, [1])


class StopWithoutForceTest(unittest.TestCase):
    """stop.bat(--force なし): 止める頼みに答えないとき、**強制終了はしない**。

    Windows の `os.kill(pid, SIGTERM)` はその場で落とす(TerminateProcess)。取り込みの
    途中かどうかも確かめずに落とすことになる。"""

    def setUp(self):
        self.kills = []
        info = _info()
        health = {"app_id": app_config.app_id(), "mode": "main", "pid": info.pid, "ready": True}
        for target, name, value in (
            (launch_guard, "read_lock", lambda mode: info),
            (launch_guard, "remove_lock", lambda mode: None),
            (launch_guard, "probe_health", lambda port, timeout=0: dict(health)),
            (launch_guard, "process_alive_state", lambda pid: True),
            (launch_guard, "is_process_alive", lambda pid: True),
            (process_manager, "_request_shutdown", lambda *a, **k: {"ok": False}),   # 時間切れなど
            (process_manager, "GRACEFUL_WAIT_SEC", 0.1),
            (process_manager, "FORCE_WAIT_SEC", 0.1),
            (os, "kill", lambda pid, sig: self.kills.append(sig)),
        ):
            patcher = mock.patch.object(target, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)

    def test_頼みに答えないだけなら_落とさずに断る(self):
        result = process_manager.stop("main", force=False)
        self.assertEqual(self.kills, [], "--force なしで落とした")
        self.assertFalse(result.stopped)

    def test_落とさないときは_forceを案内する(self):
        result = process_manager.stop("main", force=False)
        self.assertIn("--force", result.message)

    def test_force_なら落とす(self):
        process_manager.stop("main", force=True)
        self.assertTrue(self.kills)


class FeatureShutdownAsksScreensTest(unittest.TestCase):
    """機能の停止口(/material/api/shutdown など)も、統合の停止口と同じく
    **開いている画面に保存していない入力を確かめてから**止める。"""

    def setUp(self):
        import app as app_module

        from modules import packing_material_calculation as material

        self.app = _make_app(self)
        self.client = self.app.test_client()
        self.stopped = threading.Event()
        app_module.set_shutdown_hook(self.stopped.set)
        self.addCleanup(app_module.set_shutdown_hook, None)
        # 機能の停止口は、機能ごとに差し込まれた止め方を使う(server.py が起動時に差し込む)
        material.set_shutdown_hook(self.stopped.set)
        self.addCleanup(material.set_shutdown_hook, None)
        patcher = mock.patch.object(close_ask.CloseAsk, "WAIT_SEC", 0.3)
        patcher.start()
        self.addCleanup(patcher.stop)

    def test_画面が開いていれば_機能の停止口でもすぐには止めない(self):
        self.client.post("/api/alive", json={"client": "shell-1", "page": "shell-1", "state": "visible"})
        res = self.client.post("/material/api/shutdown", json={}, headers=H)
        self.assertNotEqual(res.status_code, 200, res.get_json())
        self.assertFalse(res.get_json()["ok"], "資材計算の「終了」は ok だけを見る")
        self.assertFalse(self.stopped.wait(0.8), "画面に確かめずに止めた")

    def test_forceなら画面に確かめずに止める(self):
        self.client.post("/api/alive", json={"client": "shell-1", "page": "shell-1", "state": "visible"})
        res = self.client.post("/material/api/shutdown", json={"force": True}, headers=H)
        self.assertEqual(res.status_code, 200, res.get_json())
        self.assertTrue(self.stopped.wait(3))


class OnWakeAllModulesTest(unittest.TestCase):
    """スリープから戻ったら、3機能すべての後始末を呼ぶ(並びの先頭だけでなく)。"""

    def test_持っている機能は全部呼び_1つ失敗してもほかは呼ぶ(self):
        import app as app_module
        called = []

        def boom(gap):
            raise RuntimeError("失敗")

        mods = (types.SimpleNamespace(LABEL="A", on_wake=boom),
                types.SimpleNamespace(LABEL="B"),
                types.SimpleNamespace(LABEL="C", on_wake=lambda gap: called.append(gap)))
        with mock.patch.object(app_module, "all_modules", lambda: mods):
            app_module.on_wake(42.0)
        self.assertEqual(called, [42.0])


if __name__ == "__main__":
    unittest.main()
