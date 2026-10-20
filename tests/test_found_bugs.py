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
            "request_shutdown": lambda port, token: True,
            "probe_health": lambda port, timeout=0: None,          # 待ち受けはもう答えない
            "process_alive_state": lambda pid: True,               # でもプロセスは居る
            "is_process_alive": lambda pid: True,
            "remove_lock": lambda mode: self.removed.append(mode),
            "REPLACE_WAIT_SEC": 0.3,
        }.items():
            patcher = mock.patch.object(launch_guard, name, value)
            patcher.start()
            self.addCleanup(patcher.stop)

    @unittest.expectedFailure
    def test_プロセスが残っているうちは_ロックを消して立て直さない(self):
        result = launch_guard._replace_stale("main", _info(), "1.2.0", "統合 1.2.0 → 1.2.26")
        self.assertEqual(self.removed, [], "古いプロセスが居るのにロックを消した")
        self.assertFalse(result.ok, "古いプロセスが居るのに立て直そうとした")


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

    @unittest.expectedFailure
    def test_頼みに答えないだけなら_落とさずに断る(self):
        result = process_manager.stop("main", force=False)
        self.assertEqual(self.kills, [], "--force なしで落とした")
        self.assertFalse(result.stopped)

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

    @unittest.expectedFailure
    def test_画面が開いていれば_機能の停止口でもすぐには止めない(self):
        self.client.post("/api/alive", json={"client": "shell-1", "page": "shell-1", "state": "visible"})
        res = self.client.post("/material/api/shutdown", json={}, headers=H)
        self.assertNotEqual(res.status_code, 200, res.get_json())
        self.assertFalse(self.stopped.wait(0.8), "画面に確かめずに止めた")


if __name__ == "__main__":
    unittest.main()
