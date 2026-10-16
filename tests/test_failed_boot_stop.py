"""本体を組み立てられなかった Python を、stop.bat で確かに止める(統合 1.2.19)

業務ツール統合ツール all-tools で起きた不具合と同じ形を、このツールで潰す:

    1. 起動の途中で本体を組み立てられない(フォルダの一部が欠けた・設定が壊れた・部品が無い)
       → Python は待機画面で理由を出したまま残る
    2. stop.bat で止めようとする
       → 待機画面が止める合図に答えない(以前は 503)。PID で落とす前の照合
         (コマンドラインの読み取り)ができない PC では止まらず、フォルダを消せない
    3. stop.bat が止まったかを確かめずに「終わった」と出す

ここで確かめること:

- 待機画面が止める合図(トークン付き)に答えて、自分で終わる
- 誰も見ていなければ、組み立てられなかった Python は自分で終わる
- stop.bat は**プロセスが居なくなったのを確かめてから**「終了しました」と言う
- コマンドラインを読めない PC でも、アプリ自身の答え(app_id・mode・pid)で本人と分かる
- 実際に組み立てに失敗した Python を、コマンドラインを読めない前提で止められる
"""
from __future__ import annotations

import json
import os
import socket
import subprocess
import sys
import textwrap
import threading
import time
import types
from pathlib import Path

import pytest

import boot_server
import launch_guard
import process_manager
import start_app
from common import app_config

ROOT = Path(__file__).resolve().parent.parent


# ==================================================================
# 待機画面が止める合図に答える
# ==================================================================
def _call(app, method="POST", path="/api/shutdown", token=None):
    environ = {"REQUEST_METHOD": method, "PATH_INFO": path}
    if token is not None:
        environ["HTTP_X_TOOL_TOKEN"] = token
    got = {}

    def start_response(status, headers):
        got["status"] = int(status.split()[0])

    body = b"".join(app(environ, start_response))
    return got["status"], json.loads(body.decode("utf-8"))


def _boot(calls=None):
    app = boot_server.BootApp("main", 0, "tok-123")
    if calls is not None:
        app.on_shutdown = lambda: calls.append(time.monotonic())
    return app


def test_待機画面はPOSTでないと止めない():
    calls = []
    status, _ = _call(_boot(calls), method="GET", token="tok-123")
    assert status == 405
    time.sleep(0.5)
    assert calls == []


def test_待機画面はトークンが違えば止めない():
    calls = []
    for token in (None, "", "tok-12", "tok-1234"):
        status, _ = _call(_boot(calls), token=token)
        assert status == 403
    time.sleep(0.5)
    assert calls == []


def test_待機画面は止め方を知らなければまだと答える():
    status, body = _call(_boot(), token="tok-123")
    assert status == 503
    assert body["error"]["code"] == "starting"


def test_待機画面はトークンが合えば返事をしてから止める():
    calls = []
    status, body = _call(_boot(calls), token="tok-123")
    assert status == 200 and body["stopped"] is True
    assert calls == []                      # 返事が先(合図を送った側に返事が届く)
    deadline = time.monotonic() + 3
    while not calls and time.monotonic() < deadline:
        time.sleep(0.05)
    assert len(calls) == 1


def test_様子を聞かれるたびに見られた時刻が進む():
    app = _boot()
    app.last_seen = time.monotonic() - 1000
    status, health = _call(app, method="GET", path="/api/health")
    assert status == 200 and health["ready"] is False
    assert time.monotonic() - app.last_seen < 5


def test_本体は待機画面の止める合図を自分の止め方へつなぐ():
    import server
    srv = server.AppServer("main", 0)
    assert srv.boot.on_shutdown == srv.stop


# ==================================================================
# 誰も見ていなければ自分で終わる
# ==================================================================
class _FakeSrv:
    def __init__(self, quiet_sec: float) -> None:
        self.boot = types.SimpleNamespace(last_seen=time.monotonic() - quiet_sec)
        self.stopped = threading.Event()

    @property
    def stop_requested(self) -> bool:
        return self.stopped.is_set()

    def stop(self) -> None:
        self.stopped.set()


def test_組み立てられず誰も見ていなければ終わる():
    srv = _FakeSrv(quiet_sec=10)
    t = start_app._watch_failed_boot(srv, idle_sec=1, check_sec=0.05)
    t.join(timeout=3)
    assert srv.stopped.is_set()


def test_待機画面が様子を聞き続けていれば終わらない():
    srv = _FakeSrv(quiet_sec=0)
    stop_beat = threading.Event()

    def beat():
        while not stop_beat.is_set():
            srv.boot.last_seen = time.monotonic()
            time.sleep(0.05)

    threading.Thread(target=beat, daemon=True).start()
    start_app._watch_failed_boot(srv, idle_sec=0.5, check_sec=0.05)
    time.sleep(1.2)
    assert not srv.stopped.is_set()
    stop_beat.set()                         # 画面を閉じた
    deadline = time.monotonic() + 3
    while not srv.stopped.is_set() and time.monotonic() < deadline:
        time.sleep(0.05)
    assert srv.stopped.is_set()


def test_見張りは誰も見ていない時間を数分待つ():
    # 短すぎると、理由を読んでいる最中・ブラウザの開き直しで落ちる
    assert start_app.FAILED_BOOT_IDLE_SEC >= 60
    # 待機画面は理由を出したあとも様子を聞く(見張りより十分短い間隔で)
    from common import boot_screen
    page = boot_screen.render(display_name="x", version_label="x", token="t",
                              app_id="x", poll_ms=1000, home_url="/")
    assert "watchFailed" in page and "/api/shutdown" in page and "quitBtn" in page


# ==================================================================
# stop.bat: 止まったのを確かめてから言う
# ==================================================================
@pytest.fixture
def lock(monkeypatch):
    info = types.SimpleNamespace(pid=424242, port=50999, token="tok", started_text="",
                                 app_root=str(ROOT))
    removed = []
    monkeypatch.setattr(launch_guard, "read_lock", lambda mode: info)
    monkeypatch.setattr(launch_guard, "remove_lock", lambda mode: removed.append(mode))
    monkeypatch.setattr(process_manager, "GRACEFUL_WAIT_SEC", 0.2)
    monkeypatch.setattr(process_manager, "FORCE_WAIT_SEC", 0.2)
    info.removed = removed
    return info


def _alive(monkeypatch, fn):
    """プロセスの有無の答えを差し替える(None は「確かめられない」)。"""
    monkeypatch.setattr(launch_guard, "process_alive_state", fn)
    monkeypatch.setattr(launch_guard, "is_process_alive", lambda pid: fn(pid) is not False)


def _answers(monkeypatch, info, **extra):
    health = {"app_id": app_config.app_id(), "mode": "main", "pid": info.pid,
              "ready": False, "startup_error": "アプリを組み立てられませんでした: x"}
    health.update(extra)
    monkeypatch.setattr(launch_guard, "probe_health", lambda port, timeout=0: dict(health))
    return health


def test_待ち受けが閉じてもプロセスが残っていれば終わったと言わない(monkeypatch, lock):
    _answers(monkeypatch, lock)
    monkeypatch.setattr(process_manager, "_request_shutdown", lambda *a, **k: {"ok": True})
    monkeypatch.setattr(process_manager, "_wait_gone", lambda port, t: True)
    _alive(monkeypatch, lambda pid: True)  # 残り続ける
    monkeypatch.setattr(process_manager, "_looks_like_our_process", lambda info: False)
    kills = []
    monkeypatch.setattr(os, "kill", lambda pid, sig: kills.append(sig))

    result = process_manager.stop("main", force=True)

    assert not result.stopped
    assert "まだ動いています" in result.message
    assert lock.removed == []               # 残っているあいだはロックも残す
    assert kills                            # 本人と分かったので落としにはいった


def test_起動できなかった画面から止まれば理由つきで済と言う(monkeypatch, lock):
    _answers(monkeypatch, lock)
    monkeypatch.setattr(process_manager, "_request_shutdown", lambda *a, **k: {"ok": True})
    monkeypatch.setattr(process_manager, "_wait_gone", lambda port, t: True)
    _alive(monkeypatch, lambda pid: False)

    result = process_manager.stop("main")

    assert result.stopped and result.method == "graceful"
    assert "起動できなかった画面" in result.message
    assert lock.removed == ["main"]


def test_コマンドラインを読めなくてもアプリ自身の答えで本人と分かる(monkeypatch, lock):
    _answers(monkeypatch, lock)
    # 待機画面が合図に答えない(古い版・止めきれない)。PID で止めるしかない
    monkeypatch.setattr(process_manager, "_request_shutdown", lambda *a, **k: {"ok": False})
    monkeypatch.setattr(process_manager, "_looks_like_our_process", lambda info: False)
    alive = {"v": True}
    _alive(monkeypatch, lambda pid: alive["v"])
    monkeypatch.setattr(os, "kill", lambda pid, sig: alive.update(v=False))

    result = process_manager.stop("main")

    assert result.stopped, result.message
    assert lock.removed == ["main"]


def test_pidが違う答えでは本人と見なさない(monkeypatch, lock):
    _answers(monkeypatch, lock, pid=lock.pid + 1)  # ポートに居るのは別のプロセス
    monkeypatch.setattr(process_manager, "_request_shutdown", lambda *a, **k: {"ok": False})
    monkeypatch.setattr(process_manager, "_looks_like_our_process", lambda info: False)
    _alive(monkeypatch, lambda pid: True)
    kills = []
    monkeypatch.setattr(os, "kill", lambda pid, sig: kills.append(sig))

    result = process_manager.stop("main", force=True)

    assert not result.stopped
    assert "このアプリではない" in result.message
    assert kills == []


def test_強制終了は居なくなったのを確かめてから済と言う(monkeypatch, lock):
    _answers(monkeypatch, lock)
    monkeypatch.setattr(process_manager, "_request_shutdown", lambda *a, **k: {"ok": False})
    monkeypatch.setattr(process_manager, "_looks_like_our_process", lambda info: True)
    state = {"sigs": []}

    def kill(pid, sig):
        state["sigs"].append(sig)

    # SIGTERM でも SIGKILL でも居なくならない(タスクマネージャーでしか切れない状態)
    _alive(monkeypatch, lambda pid: True)
    monkeypatch.setattr(os, "kill", kill)

    result = process_manager.stop("main", force=True)

    assert not result.stopped
    assert "まだ動いています" in result.message
    assert lock.removed == []
    assert not str(result).startswith("[済]")  # 「済」とは出さない


def test_プロセスの有無を確かめられないPCでは確かめられないと添える(monkeypatch, lock):
    # tasklist を実行できない PC。「まだ動いています」と嘘をつかず、「済」とも言い切らない
    _answers(monkeypatch, lock)
    monkeypatch.setattr(process_manager, "_request_shutdown", lambda *a, **k: {"ok": True})
    monkeypatch.setattr(process_manager, "_wait_gone", lambda port, t: True)
    _alive(monkeypatch, lambda pid: None)

    result = process_manager.stop("main")

    assert result.stopped and result.method == "graceful"
    assert result.message.endswith(process_manager.UNSURE_NOTE)
    assert lock.removed == ["main"]


def test_確かめられず待ち受けも残っていれば止まったと言わない(monkeypatch, lock):
    _answers(monkeypatch, lock)
    monkeypatch.setattr(process_manager, "_request_shutdown", lambda *a, **k: {"ok": False})
    monkeypatch.setattr(process_manager, "_wait_gone", lambda port, t: False)
    monkeypatch.setattr(process_manager, "_looks_like_our_process", lambda info: True)
    _alive(monkeypatch, lambda pid: None)
    monkeypatch.setattr(os, "kill", lambda pid, sig: None)

    result = process_manager.stop("main", force=True)

    assert not result.stopped
    assert "まだ動いています" in result.message
    assert lock.removed == []


# ==================================================================
# 通し: 実際に組み立てに失敗した Python を止める
# ==================================================================
_LAUNCH = textwrap.dedent("""
    import sys
    sys.path.insert(0, sys.argv[1])
    import server

    def build(self):
        raise RuntimeError("部品が足りません(試験)")

    server.AppServer.build = build
    import start_app
    start_app.FAILED_BOOT_IDLE_SEC = float(sys.argv[2])
    start_app.FAILED_BOOT_CHECK_SEC = 0.2
    raise SystemExit(start_app.main(["--no-browser"]))
""")


def _free_port() -> int:
    with socket.socket() as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]


@pytest.fixture
def failed_boot(tmp_path, monkeypatch):
    """組み立てに失敗して待機画面で残った Python。試験の側も同じロックを読む。"""
    conf = tmp_path / "app.json"
    real = json.loads((ROOT / "config" / "app.json").read_text(encoding="utf-8"))
    real["server"] = dict(real.get("server") or {}, host="127.0.0.1", port_retry=3,
                          roles={"main": {"port": _free_port()}})
    conf.write_text(json.dumps(real, ensure_ascii=False), encoding="utf-8")
    env = {"COIL_PACKING_TOOLS_APP_CONFIG": str(conf),
           "COIL_PACKING_TOOLS_LOCAL_DIR": str(tmp_path / "local"),
           "COIL_PACKING_TOOLS_LOG_DIR": str(tmp_path / "logs")}
    for k, v in env.items():
        monkeypatch.setenv(k, v)
    app_config.load(force=True)
    procs = []

    def launch(idle_sec: float = 600.0):
        script = tmp_path / "launch.py"
        script.write_text(_LAUNCH, encoding="utf-8")
        proc = subprocess.Popen([sys.executable, str(script), str(ROOT), str(idle_sec)],
                                cwd=str(tmp_path), env=dict(os.environ),
                                stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        procs.append(proc)
        # 実際は起動した側(start.bat・Start.vbs)が先に終わる。試験では親がこの試験なので、
        # 終わった Python を拾っておく(拾わないと「まだ居る」に見える)
        threading.Thread(target=proc.wait, daemon=True).start()
        deadline = time.monotonic() + 60
        while time.monotonic() < deadline:
            info = launch_guard.read_lock("main")
            health = launch_guard.probe_health(info.port) if info else None
            if health and health.get("startup_error"):
                return proc
            if proc.poll() is not None:
                raise AssertionError("起動の試験用 Python が先に終わりました")
            time.sleep(0.2)
        raise AssertionError("組み立ての失敗が待機画面に出ませんでした")

    yield launch
    for proc in procs:
        if proc.poll() is None:
            proc.kill()
            proc.wait(timeout=10)
    app_config.load(force=True)


def test_通し_コマンドラインを読めないPCでもstopで止まる(failed_boot, monkeypatch):
    proc = failed_boot()
    # wmic も PowerShell も使えない PC(all-tools で止められなかった条件)
    monkeypatch.setattr(process_manager, "_looks_like_our_process", lambda info: False)

    result = process_manager.stop("main")

    assert result.stopped, result.message
    assert result.method == "graceful"
    assert "起動できなかった画面" in result.message
    assert proc.wait(timeout=15) is not None   # Python が居なくなった
    assert launch_guard.read_lock("main") is None


def test_通し_誰も見ていなければ組み立てられなかったPythonは自分で終わる(failed_boot):
    proc = failed_boot(idle_sec=2.0)
    assert proc.wait(timeout=30) == 1
    assert launch_guard.read_lock("main") is None   # 次の起動を惑わせない
