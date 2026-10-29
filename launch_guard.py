"""多重起動の防止とインスタンス管理 (基盤仕様書 2.4)

同じアプリが二重に起動すると、ポート競合・二重処理・設定ファイル競合が
起きる。そこで起動ごとに**ロックファイル**(PID・ポート・起動時刻・mode)を
書き、次の起動はそれを見て判断する。

    生きている同じアプリが居る → 新しく起動せず、ブラウザだけ開く
    死んだロックが残っている   → 消して続行する(記録は残す)

「生きているか」の判定は2段構えにする:

    1. PIDのプロセスが存在するか
    2. そのポートの `/api/health` が**同じ app_id と mode** を返すか

1だけでは足りない。PIDは使い回されるので、無関係なプロセスが同じ番号を
持っていることがある。2だけでも足りない。別のアプリがHTTPを返している
場合があるので、`app_id` の照合が要る(基盤仕様書 2.3)。

このモジュールは Flask に依存しない。起動の判断だけを持つので、
サーバが立たない状況でも動く。
"""
from __future__ import annotations

import json
import os
import socket
import subprocess
import sys
import time
import urllib.error
import urllib.request
from contextlib import contextmanager
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Optional

from common import app_config, modes
from common.logging_utils import get_logger

log = get_logger("coil_packing_tools", "launch_guard")

# `/api/health` を叩くときの待ち時間(秒)。ローカルなので短くてよい。
# 長いと、死んだロックの掃除に毎回この分だけ待たされる
HEALTH_TIMEOUT_SEC = 1.5


@dataclass
class LockInfo:
    """`runtime/<mode>.lock` の中身。"""

    app_id: str
    mode: str
    pid: int
    port: int
    url: str
    started_at: float
    python: str = ""
    app_root: str = ""
    # 起動トークン。`stop.bat` が「正常終了を要求する」ために要る
    # (基盤仕様書 2.8 は、PIDで落とす前にまずアプリ自身へ頼むことを
    # 求めている)。ロックは利用者ごとのローカル領域にあり、
    # 読めるのは同じ利用者だけ。そこまで入れる相手はプロセスを直接
    # 落とせるので、ここに置いても守りの強さは変わらない。
    # 逆に、悪意あるWebページはローカルのファイルを読めないため、
    # トークンによる防御(別オリジンからの操作を弾く)はそのまま効く。
    token: str = ""

    @property
    def started_text(self) -> str:
        return time.strftime("%Y/%m/%d %H:%M:%S", time.localtime(self.started_at))


# ------------------------------------------------------------------
# 起動の順番待ち (基盤仕様書 2.4)
# ------------------------------------------------------------------
# **ロックファイルを見るだけでは、同時に押されたときに防げません。**
#
# 判定から書き込みまでは「調べる → ポートを決める → 待ち受ける →
# ロックを書く」の4段で、ここに1秒弱かかります。その間に2つ目が
# 起動すると、**どちらもロックが無いと判断して両方が立ち上がります**
# (実測: 2つとも同じポートを選び、ロックを2回上書きして、
#  ロックの pid が実際に待ち受けているプロセスと食い違った)。
#
# そこで、この4段を**1つずつしか通れない道**にします。2つ目は少し待ち、
# 通れるようになった時点でもう一度調べるので、1つ目のロックを見つけて
# 合流します。
#
# 締め方は環境で違うので、使えるものを順に試します。
#   Windows … `msvcrt.locking`
#   その他   … `fcntl.flock`
# どちらも使えなければ、締めずに通します ── **起動できないより、
# まれに2つ立ち上がるほうがまし**です(理由は下の `startup_gate`)。
try:                                              # pragma: no cover - 環境依存
    import fcntl
except ImportError:                               # pragma: no cover - Windows
    fcntl = None
try:                                              # pragma: no cover - 環境依存
    import msvcrt
except ImportError:                               # pragma: no cover - POSIX
    msvcrt = None

# 順番待ちの上限(秒)。1つ目が待ち受けを始めてロックを書くまでの
# 実測は1秒弱なので、その数倍を見ておく
GATE_WAIT_SEC = 15.0
GATE_POLL_SEC = 0.1


def gate_path(mode: str) -> Path:
    """順番待ちの印。**ロック本体とは別のファイル。**

    同じファイルを使うと、締めるために開いたことでロックの中身を
    読み書きする処理とぶつかります。
    """
    if mode not in modes.KEYS:
        raise ValueError(f"未知のmode: {mode!r}")
    return app_config.local_dir("runtime") / f"{mode}.start"


def _try_lock(handle) -> bool:
    """締められたら True。すでに誰かが通っていれば False。"""
    if fcntl is not None:
        try:
            fcntl.flock(handle.fileno(), fcntl.LOCK_EX | fcntl.LOCK_NB)
            return True
        except OSError:
            return False
    if msvcrt is not None:                        # pragma: no cover - Windows
        try:
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_NBLCK, 1)
            return True
        except OSError:
            return False
    return True                                   # 締められない環境は素通し


def _unlock(handle) -> None:
    if fcntl is not None:
        try:
            fcntl.flock(handle.fileno(), fcntl.LOCK_UN)
        except OSError:                           # pragma: no cover - 解放し損ね
            pass
    elif msvcrt is not None:                      # pragma: no cover - Windows
        try:
            handle.seek(0)
            msvcrt.locking(handle.fileno(), msvcrt.LK_UNLCK, 1)
        except OSError:
            pass


#: デスクトップ版(exe)が動いているあいだ握っている錠(`src-tauri/src/main.rs` の
#: `take_instance_lock`)。ローカル領域の runtime の下
DESKTOP_LOCK_NAME = "desktop.lock"
#: デスクトップ版の **Python** が動いているあいだ握っている錠(`hold_desktop_lock`。統合 1.2.1)
DESKTOP_PY_LOCK_NAME = "desktop-python.lock"

#: `hold_desktop_lock` が握っている錠(プロセスが終われば OS が外す)
_desktop_hold = None


def desktop_running() -> bool:
    """デスクトップ版(コイル梱包ツール.exe)が動いているか。

    exe は動いているあいだ `runtime/desktop.lock` を、exe が起こした Python は
    `runtime/desktop-python.lock` を OS のロックで握っている(落ちれば OS が外すので、
    残った印に惑わされない)。どちらかが締められなければ動いている。
    **ブラウザ版とデスクトップ版は同時に動かさない**(同じ手元のDB・作業状態を使う)。

    【2つ見る理由】Microsoft Store の Python は `%LOCALAPPDATA%` の下のファイルを自分だけの
    場所に置き換えることがあり、exe(Python の外)と Python とで**同じ名前でも別のファイルを
    見る**ことがある。Python どうし(ブラウザ版とデスクトップ版の Python)は同じ場所を見るので、
    Python の錠も握っておけば、置き換えがどう働いても取り違えない。
    """
    folder = app_config.local_dir("runtime")
    return any(_held_by_other(folder / name) for name in (DESKTOP_LOCK_NAME, DESKTOP_PY_LOCK_NAME))


def _held_by_other(path: Path) -> bool:
    if not path.exists():
        return False
    try:
        with open(path, "a+b") as handle:
            if _try_lock(handle):
                _unlock(handle)
                return False
            return True
    except OSError:
        return False


def hold_desktop_lock() -> bool:
    """デスクトップ版の Python が、動いているあいだ `desktop-python.lock` を握る。

    握れたら True。プロセスが終われば OS が外す(残った印で起動できなくなることはない)。
    """
    global _desktop_hold
    if _desktop_hold is not None:
        return True
    path = app_config.local_dir("runtime") / DESKTOP_PY_LOCK_NAME
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        handle = open(path, "a+b")
    except OSError as exc:
        log.warning("デスクトップ版の印を置けませんでした(そのまま続けます): %s", exc)
        return False
    if not _try_lock(handle):
        handle.close()
        log.warning("デスクトップ版の印をほかが握っています: %s", path)
        return False
    _desktop_hold = handle
    return True


def release_desktop_lock() -> None:
    """試験用(本物はプロセスが終われば外れる)。"""
    global _desktop_hold
    if _desktop_hold is not None:
        _unlock(_desktop_hold)
        _desktop_hold.close()
        _desktop_hold = None


@contextmanager
def startup_gate(mode: str, *, wait_sec: float = GATE_WAIT_SEC):
    """起動の4段を1つずつしか通さない。

    2つ目はここで待ち、通れるようになってから `check_existing` を
    やり直します ── **待っているあいだに1つ目がロックを書き終える**ので、
    そこで合流できます。

    【待ちきれなくても通します】
    締められない環境(ファイルロックが効かない共有フォルダ等)や、
    1つ目が固まっている場合に、ここで起動そのものを止めると
    **アプリが二度と開けなくなります**。まれに2つ立ち上がるほうが、
    確実に開けないより害が小さいので、理由を残して通します。
    """
    path = gate_path(mode)
    path.parent.mkdir(parents=True, exist_ok=True)
    handle = None
    held = False
    try:
        handle = open(path, "a+b")
        deadline = time.monotonic() + wait_sec
        waited = False
        while True:
            if _try_lock(handle):
                held = True
                break
            if time.monotonic() >= deadline:
                log.warning("起動の順番待ちが %.0f秒を過ぎました。"
                            "そのまま続けます", wait_sec)
                break
            if not waited:
                log.info("ほかの起動が進んでいます。順番を待ちます")
                waited = True
            time.sleep(GATE_POLL_SEC)
        yield held
    except OSError as exc:                        # noqa: BLE001 - 締められなくても続ける
        log.warning("起動の順番待ちを用意できませんでした(そのまま続けます): %s", exc)
        yield False
    finally:
        if handle is not None:
            if held:
                _unlock(handle)
            try:
                handle.close()
            except OSError:                       # pragma: no cover
                pass


def lock_path(mode: str) -> Path:
    """mode ごとのロック。統合版のモードは `main` の1つだけ(移植元は現場/資材の2つだった)。"""
    if mode not in modes.KEYS:
        raise ValueError(f"未知のmode: {mode!r}")
    return app_config.local_dir("runtime") / f"{mode}.lock"


# ------------------------------------------------------------------
# ロックファイル
# ------------------------------------------------------------------
def write_lock(info: LockInfo) -> Path:
    path = lock_path(info.mode)
    path.parent.mkdir(parents=True, exist_ok=True)
    # 別名で書いてから置き換える(統合 1.2.61)。書いている途中を読まれると「壊れている = 無い」と
    # 取り違えられ、終わるプロセスが新しいほうのロックを消すことがあった
    tmp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    tmp.write_text(json.dumps(asdict(info), ensure_ascii=False, indent=2), encoding="utf-8")
    os.replace(tmp, path)
    # トークンを含むので、本人だけが読める権限にする。
    # Windows には chmod が効かないが、`%LOCALAPPDATA%` 自体が
    # 利用者ごとのフォルダなので実質同じ扱いになる
    try:
        os.chmod(path, 0o600)
    except OSError as exc:                        # noqa: BLE001 - 権限設定は best effort
        log.debug("ロックの権限を変更できませんでした: %s", exc)
    log.info("ロックを書きました: %s (pid=%s port=%s)", path, info.pid, info.port)
    return path


def read_lock(mode: str) -> Optional[LockInfo]:
    """壊れていれば `None`。読めないことを理由に起動を止めない。"""
    return read_lock_state(mode)[1]


#: `read_lock_state` の答え
LOCK_ABSENT = "absent"
LOCK_OK = "ok"
LOCK_UNREADABLE = "unreadable"


def read_lock_state(mode: str) -> tuple[str, Optional[LockInfo]]:
    """(状態, ロック)。状態は 無い(`absent`)・読めた(`ok`)・**あるのに読めない**(`unreadable`)。

    統合 1.2.61。「読めない」を「無い」と取り違えない(別のツールで、読めなかったのを「無い」と取り違えて
    送り直し、重複を大量に作った事案があった)。消す・止まったと言う前はこちらで見る。
    """
    path = lock_path(mode)
    try:
        raw = json.loads(path.read_text(encoding="utf-8"))
        return LOCK_OK, LockInfo(**{k: raw[k] for k in LockInfo.__dataclass_fields__ if k in raw})
    except FileNotFoundError:
        return LOCK_ABSENT, None
    except Exception as exc:                      # noqa: BLE001 - 壊れたロックは起動を止めない
        log.warning("ロックを読めませんでした (%s): %s", path, exc)
        return LOCK_UNREADABLE, None


def remove_lock(mode: str, *, pid: Optional[int] = None) -> None:
    """ロックを消す。`pid` を渡すと、**そのプロセスのロックのときだけ**消す。

    終わるプロセスが自分のロックを片付けるときは `pid` を渡す(統合 1.2.26)。版の入れ替えでは
    古いほうが終わりきる前に新しいほうがロックを書くことがあり、古いほうが終わりぎわに
    新しいほうのロックを消していた(次の起動が2つ目を立て、stop.bat は「起動していません」)。
    """
    if pid is not None:
        current = read_lock(mode)
        if current is None and read_lock_state(mode)[0] == LOCK_UNREADABLE:
            # 自分のものか確かめられない。消さない(統合 1.2.61)
            log.info("ロックを読めないので、自分のものか確かめられず残します")
            return
        if current is not None and int(current.pid) != int(pid):
            log.info("ロックは別のプロセス(pid %s)のものなので残します", current.pid)
            return
    try:
        lock_path(mode).unlink()
        log.info("ロックを消しました: %s", mode)
    except FileNotFoundError:
        pass
    except OSError as exc:                        # noqa: BLE001
        log.warning("ロックを消せませんでした: %s", exc)


# ------------------------------------------------------------------
# プロセスの生死
# ------------------------------------------------------------------
def is_process_alive(pid: int) -> bool:
    """そのPIDのプロセスが存在するか。中身までは見ない。分からなければ「居る」。"""
    return process_alive_state(pid) is not False


def process_alive_state(pid: int) -> Optional[bool]:
    """そのPIDのプロセスが居るか。**確かめられなければ None**(tasklist を実行できないなど)。

    止めたあとに「終わりました」と言えるかの判定に使う(統合 1.2.19)。分からないのに
    「居る」と答えると、止まっていても「まだ動いています」と出してしまう。
    """
    if pid <= 0:
        return False
    if os.name == "nt":
        return _is_alive_windows(pid)
    try:
        os.kill(pid, 0)          # シグナル0は存在確認だけ
    except ProcessLookupError:
        return False
    except PermissionError:
        return True              # 別ユーザーのプロセス = 生きている
    return not _is_zombie(pid)


def _is_zombie(pid: int) -> bool:
    """Linux で、終わったが親にまだ引き取られていないプロセス(ゾンビ)か。

    ゾンビにもシグナル0は届くので「居る」と見えてしまい、終わったのに「止まらない」と判定していた
    (統合 1.2.51。試験の台本が起動したプロセスを、止めたあとで引き取るときに出た。Windows には無い)。
    """
    try:
        with open(f"/proc/{pid}/stat", "rb") as f:
            stat = f.read()
    except OSError:
        return False             # /proc が無い(mac など)・読めない → これまでどおり
    end = stat.rfind(b")")       # 名前に空白や括弧が入ってもよいよう、最後の ")" の後ろを見る
    return end >= 0 and stat[end + 2:end + 3] == b"Z"


def console_text(raw: Optional[bytes]) -> str:
    """Windows のコマンド(tasklist・wmic・PowerShell)の出力を文字にする。**落ちない**。

    出力は端末のコードページ(日本語の Windows では cp932)で来る。`text=True` に任せると、
    Python の UTF-8 モード(`-X utf8`・環境変数 `PYTHONUTF8=1`)では UTF-8 として読み、
    日本語の行(tasklist の「情報: 指定された条件に一致するタスクは…」)で読み手のスレッドが
    落ちて、結果が None になる。統合 1.2.0 のデスクトップ版は、これで**起動できなかった**
    (現場の PC。英語の Windows の自動確認では日本語が出ないので見つからなかった)。
    そこでバイトのまま受け取り、ここで読み方を決める(読めない字は置き換える)。
    """
    if not raw:
        return ""
    if raw[:2] == b"\xff\xfe" or (len(raw) > 3 and raw[1:2] == b"\x00" and raw[3:4] == b"\x00"):
        return raw.decode("utf-16-le", "replace").lstrip("\ufeff")    # wmic をファイルへ向けたときの形
    if os.name == "nt":
        for encoding in ("oem", "mbcs"):                                # 端末のコードページ
            try:
                return raw.decode(encoding, "replace")
            except LookupError:
                continue
    return raw.decode("utf-8", "replace")


def _run_console(args: list, timeout: float = 5) -> bytes:
    """Windows のコマンドを黒い窓を出さずに流し、出力を**バイトのまま**返す(読み方は呼び手)。"""
    out = subprocess.run(
        args, capture_output=True, timeout=timeout, stdin=subprocess.DEVNULL,
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    return out.stdout or b""


def _is_alive_windows(pid: int) -> Optional[bool]:
    """Windows では `tasklist` で確認する。

    `OpenProcess` を ctypes で叩く手もあるが、権限やハンドルの後始末を
    誤ると別の不具合を招く。起動時に1回だけの判定なので、
    外部コマンドの数十msは問題にならない。

    **出力はバイトのまま見る**(`"1234"` は ASCII なので、どのコードページでも同じバイト)。
    文字にしないので、日本語の「情報: …」が出ても、UTF-8 モードでも落ちない。
    """
    try:
        raw = _run_console(["tasklist", "/FI", f"PID eq {pid}", "/NH", "/FO", "CSV"])
    except (OSError, subprocess.SubprocessError) as exc:
        log.warning("tasklist を実行できませんでした: %s", exc)
        return None              # 分からない(is_process_alive は「生きている」に倒す)
    return f'"{pid}"'.encode("ascii") in raw


def process_command_line(pid: int) -> str:
    """そのPIDが何を実行しているか。**停止前の確認に使う**。

    基盤仕様書 2.8 は「Pythonをプロセス名だけで一括終了しない」ことを
    求めている。無関係なPythonアプリを巻き添えにしないため、
    落とす前にコマンドラインとアプリの場所を照合する。
    取れなければ空文字(呼び出し側は安全側に倒して落とさない)。
    """
    if pid <= 0:
        return ""
    if os.name == "nt":
        return _command_line_windows(pid)
    try:
        raw = Path(f"/proc/{pid}/cmdline").read_bytes()
        return raw.replace(b"\0", b" ").decode("utf-8", "replace").strip()
    except OSError:
        return ""


def _command_line_windows(pid: int) -> str:
    """Windows でそのPIDのコマンドライン。取れなければ空文字。

    wmic は Windows 11 の新しい版では入っていないことがある。無ければ PowerShell で引く。
    出力はバイトで受け、`console_text` で読む(日本語の Windows・UTF-8 モードでも落ちない)。
    """
    tries = (
        ["wmic", "process", "where", f"ProcessId={pid}", "get", "CommandLine", "/format:list"],
        ["powershell", "-NoProfile", "-NonInteractive", "-Command",
         f"$p = Get-CimInstance Win32_Process -Filter 'ProcessId={pid}'; "
         f"if ($p) {{ 'CommandLine=' + $p.CommandLine }}"],
    )
    for args in tries:
        try:
            text = console_text(_run_console(args, timeout=10))
        except (OSError, subprocess.SubprocessError) as exc:
            log.warning("%s を実行できませんでした: %s", args[0], exc)
            continue
        for line in text.splitlines():
            if line.startswith("CommandLine="):
                return line.split("=", 1)[1].strip()
    return ""


# ------------------------------------------------------------------
# 起動確認API
# ------------------------------------------------------------------
# 自分自身(127.0.0.1)への通信に**プロキシを通さない**ための送信口。
#
# `urllib.request.urlopen()` の既定は、環境変数 `HTTP_PROXY` や
# Windowsのインターネット設定からプロキシを拾う。社内PCではたいてい
# プロキシが入っており、除外一覧に `127.0.0.1` が無いと、**自分自身への
# 通信までプロキシへ送られて失敗する**。
#
# 実際に踏んだ: サーバは待ち受けを始めているのに `/api/health` が
# 返らず、15秒待って「サーバを起動できませんでした」になった。
# ループバックにプロキシを挟む理由は無いので、ここで明示的に外す。
_LOCAL_OPENER = urllib.request.build_opener(urllib.request.ProxyHandler({}))


def local_request(url: str, *, timeout: float, data: Optional[bytes] = None,
                  method: Optional[str] = None,
                  headers: Optional[dict] = None):
    """127.0.0.1 への要求。プロキシを経由しない。

    このアプリが外に出す通信はこれだけなので、送信口を1つに集約しておく。
    """
    request = urllib.request.Request(url, data=data, method=method,
                                     headers=headers or {})
    return _LOCAL_OPENER.open(request, timeout=timeout)


def probe_health(port: int, *, timeout: float = HEALTH_TIMEOUT_SEC) -> Optional[dict]:
    """`GET /api/health` を叩く。応答しなければ `None`。"""
    url = f"http://127.0.0.1:{port}/api/health"
    try:
        with local_request(url, timeout=timeout) as res:
            return json.loads(res.read().decode("utf-8"))
    except (urllib.error.URLError, OSError, ValueError, json.JSONDecodeError):
        return None


def is_port_accepting(port: int, *, timeout: float = 0.5) -> bool:
    """TCPで繋がるか。HTTPまでは見ない。

    「待ち受けていない」のか「待ち受けてはいるが応答が返らない」のかを
    分けるために使う。後者はプロキシやセキュリティ製品が挟まっている
    ことが多く、直し方がまったく違う。
    """
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.settimeout(timeout)
        return sock.connect_ex(("127.0.0.1", port)) == 0


def proxy_settings() -> dict[str, str]:
    """いま効いているプロキシ設定。原因調査のためだけに使う。"""
    try:
        return {k: v for k, v in urllib.request.getproxies().items()
                if k in ("http", "https")}
    except Exception:                    # noqa: BLE001 - 調査用なので握る
        return {}


def is_our_app(health: Optional[dict], mode: str) -> bool:
    """その応答が**自分と同じアプリの同じmode**か。

    `app_id` を見ないと、たまたま同じポートを使っている別のアプリを
    自分だと誤認する(基盤仕様書 2.3)。mode まで見るのは移植元の名残(移植元は現場と資材が
    別ポートで動いた)。統合版のモードは `main` だけなので、実際は app_id で見分けている。
    """
    if not health:
        return False
    return (health.get("app_id") == app_config.app_id()
            and health.get("mode") == mode)


# ------------------------------------------------------------------
# 既存インスタンスの判定
# ------------------------------------------------------------------
@dataclass
class GuardResult:
    """起動してよいか。"""

    should_start: bool
    url: str = ""
    existing: Optional[LockInfo] = None
    reason: str = ""
    # 動いているのが**古い版**だった場合、その版。合流してはいけない
    stale_version: str = ""


# 古い版が動いていたとき、終わるのを待つ上限(秒)
REPLACE_WAIT_SEC = 10.0


def check_existing(mode: str) -> GuardResult:
    """すでに同じアプリが動いていないか調べる。

    動いていれば `should_start=False` と、開くべきURLを返す。

    **ただし版が違えば合流しません**(統合ツールの版と3機能の版の組で比べる。
    `_stale_versions`)。入れ替えたのに古いプロセスが
    残っていると、ここが「すでに起動しています」と答えて古いほうの
    ブラウザを開き、**新しい版がいつまでも動きません**(画面の版バッジも
    古いままになる)。版が違うときは古いほうを終わらせて、こちらで
    立て直します。
    """
    info = read_lock(mode)
    if info is None:
        return GuardResult(True, reason="ロックなし")

    if not is_process_alive(info.pid):
        log.info("死んだロックを掃除します (pid=%s は不在)", info.pid)
        remove_lock(mode)
        return GuardResult(True, reason=f"ロックは残っていたがpid {info.pid} は不在")

    health = probe_health(info.port)
    if health is None:
        # 答えない。**別物と決めつけない**(統合 1.2.26)。以前はここでロックを消して2つ目を
        # 立て、同じ手元の DB を2つで開くことがあった(1.5秒で答えられないほど重い処理の途中、
        # または終わりぎわで待ち受けを閉じたあと片付けている最中)
        if not is_port_free(info.port):
            log.warning("pid=%s は生きていて %s を待ち受けているが答えない。合流します",
                        info.pid, info.port)
            return GuardResult(False, url=info.url, existing=info,
                               reason="同じアプリが応答しません(処理の途中かもしれません)")
        if wait_pid_gone(info.pid, REPLACE_WAIT_SEC) is not False:
            log.info("終わりぎわのプロセス(pid=%s)が終わるのを待ちました", info.pid)
        remove_lock(mode)
        return GuardResult(True, reason="ロックのプロセスは終わった(または別物だった)")
    if not is_our_app(health, mode):
        # プロセスは居るがアプリではない(PIDの使い回し、または
        # 起動途中で落ちた)。ロックを消して新しく起動する
        log.info("pid=%s は生きているが %s:%s は自分ではない",
                 info.pid, mode, info.port)
        remove_lock(mode)
        return GuardResult(True, reason="ロックのプロセスは別物だった")

    running = str((health or {}).get("version", ""))
    stale = _stale_versions(health or {})
    if stale:
        return _replace_stale(mode, info, running or "(不明)", " / ".join(stale), health)

    log.info("すでに起動しています: %s (pid=%s port=%s 版=%s)",
             mode, info.pid, info.port,
             str((health or {}).get("version_set", "")) or running or "(不明)")
    return GuardResult(False, url=info.url, existing=info,
                       reason="同じアプリが起動中")


def _stale_versions(health: dict) -> list:
    """動いているプロセスと、このファイルの版が違えば、その違い。同じなら空。

    **統合ツールの版と3機能の版の組で比べる**(common/versions.py)。機能の版
    だけを上げて配ったときも入れ替える ── 統合ツールの版だけで比べると、
    古いプロセスに合流して、新しい機能の中身がいつまでも動かない。
    組を答えない古いプロセスは、統合ツールの版だけで比べる。
    """
    from common import versions

    running = str(health.get("version", ""))
    running_set = str(health.get("version_set", ""))
    if running_set:
        try:
            mine = versions.all_versions()
        except Exception as exc:                    # noqa: BLE001 - 判定で起動を止めない
            log.warning("機能の版を読めませんでした(統合ツールの版だけで比べます): %s", exc)
            mine = {versions.APP_KEY: app_config.version()}
            theirs = {versions.APP_KEY: running}
        else:
            theirs = versions.parse_set(running_set)
        return versions.differences(theirs, mine)
    if running and running != app_config.version():
        return [f"{app_config.display_name()} {running} → {app_config.version()}"]
    return []


def _replace_stale(mode: str, info: "LockInfo",
                   running: str, changes: str, health: Optional[dict] = None) -> GuardResult:
    """動いているのが違う版。**終わらせてから立て直す。**

    `running` は動いているプロセスの統合ツールの版、`changes` は違いの説明
    (例 `梱包明細 0.13.1 → 0.13.2`)。

    処理の途中(取り込みなど)なら終わらせません ── 中途半端なデータを
    残すほうが害が大きいので、そのときは合流して、次の起動に任せます。
    """
    log.warning("動いているのは別の版です(%s)。古いほうを終わらせます", changes)
    pid = owner_pid(info, health)
    # force は付けない: 取り込みの途中・画面で確かめ中なら断られ、そのときは止めずに合流する
    # (中途半端なデータを残すほうが害が大きい。次の起動に任せる)
    if not ask_shutdown(info.port, info.token, timeout=5).get("ok"):
        log.warning("古い版を終わらせられませんでした。合流します")
        return GuardResult(False, url=info.url, existing=info,
                           stale_version=running,
                           reason=f"別の版が動いていますが止められません({changes})")

    # **待ち受けが閉じただけでは「終わった」としない**(統合 1.2.26)。閉じたあとも古い版は
    # 機能の片付け(手元の DB を閉じる)を続ける。終わる前に立て直すと、同じ DB を2つで開く
    gone = confirm_gone(pid, info.port, REPLACE_WAIT_SEC)
    if gone is False:
        log.warning("古い版(pid %s)が %.0f秒 で終わりませんでした。合流します", pid, REPLACE_WAIT_SEC)
        return GuardResult(False, url=info.url, existing=info,
                           stale_version=running,
                           reason=f"別の版が動いています({changes})")

    remove_lock(mode)
    log.info("古い版を終わらせました(%s)。新しい版で立て直します", changes)
    return GuardResult(True, stale_version=running,
                       reason=f"別の版を終わらせました({changes})")


# ------------------------------------------------------------------
# 統合前の単体版(同じ手元のデータを使う)
# ------------------------------------------------------------------
@dataclass
class LegacyInstance:
    """動いている単体版。"""

    label: str
    port: int
    pid: int = 0

    def describe(self) -> str:
        return f"{self.label}(ポート {self.port}" + (f"・pid {self.pid}" if self.pid else "") + ")"


#: 統合前の単体版: (呼び名, アプリID, 既定のポート)。
#: 3つとも **統合版と同じ手元のデータ**(%LOCALAPPDATA% の PackingDetails・
#: PackingPenaLabel・CoilMaterialTool)を読み書きする
LEGACY_APPS = (
    ("梱包明細(単体版)", "nlm.packing-details", 8733),
    ("ペナラベル(単体版)", "PackingPenaLabel", 8731),
    ("資材計算(単体版)", "nlm.coil-material-tool", 8733),
)


def _legacy_lock_files() -> dict:
    """単体版が起動中に残す記録(pid とポート)の置き場所。統合版と同じ手元の領域。"""
    out: dict = {}
    try:
        from modules.packing_details.meisai import app_config as details_cfg
        out["nlm.packing-details"] = details_cfg.local_dir("runtime") / "meisai.lock"
    except Exception:                                   # noqa: BLE001 - 判定で起動を止めない
        pass
    try:
        from modules.packing_material_calculation.coil_tool import app_config as material_cfg
        out["nlm.coil-material-tool"] = material_cfg.local_dir("runtime") / "main.lock"
    except Exception:                                   # noqa: BLE001
        pass
    try:
        from modules.packing_pena_label.app.config import load_config
        out["PackingPenaLabel"] = Path(load_config().runtime_dir) / "instance.json"
    except Exception:                                   # noqa: BLE001
        pass
    return out


def _answers_as(port: int, app_id: str) -> bool:
    health = probe_health(port, timeout=LEGACY_PROBE_SEC)
    if not health:
        return False
    return (health.get("app_id") or health.get("appId")) == app_id


#: 単体版の起動確認を待つ上限(秒)。**統合版の起動を遅らせない**よう短く
LEGACY_PROBE_SEC = 0.6


def find_legacy_instances() -> list:
    """**統合前の単体版**が動いていないか。動いていれば `LegacyInstance` の並び。

    【なぜ要るのか】
    統合版は単体版と**同じ手元のデータ**(作業状態・手元DB・設定)を使う。
    多重起動の印(ロック)は統合版のものしか見ないので、移行の途中で単体版の
    Start.vbs を押すと、両方が動いて同じデータを書き換える(移植漏れの点検で
    見つかった)。ペナラベルでは、画面の検査番号と刷られる検査番号が
    食い違いうる。統合版はここで見つけたら起動しない(単体版は自動では止めない ──
    保存前の作業を持っているかもしれないため)。

    見分け方: 単体版が残す記録(pid とポート)を読み、そのポートの起動確認が
    単体版のアプリIDを返せば動いている。記録が無いときは既定のポートも見る。
    """
    found: list = []
    locks = _legacy_lock_files()
    for label, app_id, default_port in LEGACY_APPS:
        try:
            port, pid = 0, 0
            path = locks.get(app_id)
            if path is not None:
                try:
                    data = json.loads(Path(path).read_text(encoding="utf-8"))
                    port, pid = int(data.get("port") or 0), int(data.get("pid") or 0)
                except (OSError, ValueError, TypeError):
                    port, pid = 0, 0
            if port and pid and is_process_alive(pid) and _answers_as(port, app_id):
                found.append(LegacyInstance(label, port, pid))
                continue
            if _answers_as(default_port, app_id):
                found.append(LegacyInstance(label, default_port))
        except Exception as exc:                        # noqa: BLE001 - 見分けで起動を止めない
            # 単体版が動いているかの見分けは**念のため**。見分けられないことで起動を
            # 止めない(統合 1.2.0 で、見分けの途中の思わぬエラーで起動できなかった)
            log.warning("単体版(%s)が動いているか見分けられませんでした: %s", label, exc)
    return found


# ------------------------------------------------------------------
# 止める頼みと、止まったかの確かめ(停止 `process_manager` と版の入れ替え `_replace_stale` で共通)
# ------------------------------------------------------------------
# 統合 1.2.26: 以前は停止と版の入れ替えで別々に書いてあり、1.2.19 の「待ち受けが閉じた
# だけでは終わったとみなさない」修正が停止の側にしか入っていなかった(構造点検で見つかった)。
def ask_shutdown(port: int, token: str, *, force: bool = False, timeout: float = 10.0) -> dict:
    """`POST /api/shutdown` で正常終了を頼む(基盤仕様書 2.8 の1段目)。

    戻り値: ``{"ok": 受け付けられた}``。断られた(409)ときは ``busy`` / ``reason``(busy・
    asking・refused)/ ``message`` / ``running`` も返す。届かない・その他の断りは ``{"ok": False}``。

    トークンはロックファイルから読む。ロックは利用者ごとのローカル領域にあるので、
    読める相手はそもそもプロセスを直接落とせる。プロキシを経由しない送信口
    (`local_request`)を使う ── 社内PCのプロキシ設定に 127.0.0.1 の除外が無いと、
    自分自身への停止要求までプロキシへ送られて届かない。
    """
    url = f"http://127.0.0.1:{port}/api/shutdown"
    body = json.dumps({"force": force}).encode("utf-8")
    headers = {"Content-Type": "application/json"}
    if token:
        headers["X-Tool-Token"] = token
    try:
        with local_request(url, timeout=timeout, data=body, method="POST",
                           headers=headers) as res:
            payload = json.loads(res.read().decode("utf-8"))
            return {"ok": bool(payload.get("stopped"))}
    except urllib.error.HTTPError as exc:
        if exc.code == 409:                        # 実行中の処理・画面で確かめ中・「閉じない」
            try:
                payload = json.loads(exc.read().decode("utf-8"))
            except Exception:                      # noqa: BLE001
                payload = {}
            reason = str(payload.get("reason") or "busy")
            return {"ok": False, "busy": reason == "busy", "reason": reason,
                    "message": payload.get("message", ""),
                    "running": payload.get("running", ["(不明)"])}
        log.info("停止要求は %s で断られました", exc.code)
        return {"ok": False}
    except (urllib.error.URLError, OSError, ValueError) as exc:
        log.info("停止要求を送れませんでした (%s)", exc)
        return {"ok": False}


def wait_port_closed(port: int, timeout: float) -> bool:
    """そのポートが応答しなくなるまで待つ。"""
    deadline = time.monotonic() + timeout
    while time.monotonic() < deadline:
        if probe_health(port, timeout=0.3) is None:
            return True
        time.sleep(0.2)
    return False


def wait_pid_gone(pid: int, timeout: float) -> Optional[bool]:
    """居なくなるまで待つ。True: 居なくなった / False: まだ居る / None: 確かめられない。"""
    deadline = time.monotonic() + timeout
    while True:
        state = process_alive_state(pid)
        if state is False:
            return True
        if state is None:
            return None
        if time.monotonic() >= deadline:
            return False
        time.sleep(0.2)


def confirm_gone(pid: int, port: int, timeout: float) -> Optional[bool]:
    """止めたあと、本当に居なくなったか(統合 1.2.19 の決まり)。

    True: 居なくなった / False: まだ居る / None: プロセスの有無は確かめられないが、
    待ち受けは閉じている(tasklist を実行できない PC)。待ち受けも残っていれば False。
    **待ち受けが閉じただけでは「終わった」としない** ── 閉じたあとも機能の片付け
    (手元の DB を閉じるなど)が数秒続く。
    """
    gone = wait_pid_gone(pid, timeout)
    if gone is None:
        return None if wait_port_closed(port, timeout) else False
    return gone


def owner_pid(info: "LockInfo", health: Optional[dict]) -> int:
    """そのポートで実際に答えたプロセスの pid(答えに無ければロックの pid)。"""
    try:
        return int((health or {}).get("pid") or 0) or int(info.pid)
    except (TypeError, ValueError):
        return int(info.pid)


# ------------------------------------------------------------------
# ポート
# ------------------------------------------------------------------
def is_port_free(port: int, host: str = "") -> bool:
    """そのポートで待ち受けを開始できるか。

    `SO_REUSEADDR` は**付けない**。付けると TIME_WAIT のポートまで
    「空いている」と判定してしまい、実際の bind で失敗する。
    """
    host = host or app_config.host()
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        try:
            sock.bind((host, port))
        except OSError:
            return False
    return True


def pick_port(mode: str) -> Optional[int]:
    """使えるポートを1つ選ぶ。全部塞がっていれば `None`。

    候補が mode どうしで重ならないことは `app_config.port_range_conflicts()` が確かめる
    (移植元の2モードの名残。統合版のモードは `main` だけ)。
    """
    for port in app_config.port_candidates(mode):
        if is_port_free(port):
            return port
        log.info("ポート %s は使用中です", port)
    return None


# ------------------------------------------------------------------
# 起動時の記録
# ------------------------------------------------------------------
def build_lock_info(mode: str, port: int, token: str = "") -> LockInfo:
    return LockInfo(
        app_id=app_config.app_id(),
        mode=mode,
        pid=os.getpid(),
        port=port,
        url=f"http://{app_config.host()}:{port}/",
        started_at=time.time(),
        python=sys.executable,
        app_root=str(app_config.APP_ROOT),
        token=token,
    )
