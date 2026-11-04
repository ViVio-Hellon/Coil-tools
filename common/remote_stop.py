"""管理者以外停止モード ── 管理者の PC から、ほかの PC のコイル梱包ツールを止める(統合 1.2.67)

現場の声:「ランチャーを介さず、アクセス権限・管理者の条件がそろっている PC から、他の PC を停止したい」。
業務ツール(python-web-tools VER4.8.129〜130 の `remote_stop.py`)と同じ決まりで作る:

    お願いを出せる      管理者パスワード **かつ** この端末のアクセス権限に Administrator
    止める範囲          お願いを出した PC 以外の、起動中の PC
    保存していない入力・ 画面に「管理者が更新のため停止を求めています」と出して利用者に任せる。
    処理中               何も残っていなければまもなく止まる。**5分たっても止まらなければ強制**で止める
                        (保存していない入力は控えに残し、次の起動で知らせる。`common/stop_keep.py`)
    更新中               お願いが出ているあいだ、お願いを出した PC 以外では**起動させない**
    ランチャー           「全 PC で止める」は据え置き。ただし**同時には使わない**
                        (お願いが出ているあいだ stop.bat は断る)

【どこに置くか】
共有の梱包資材マスタのフォルダ(資材計算の `config.master_db_dir()`。配布設定で全 PC が同じ場所を見る)の
`管理者以外停止\\コイル梱包ツール\\` に、小さな JSON を置く。共有の SQLite には書かない(錠の取り合いをしない)。
業務ツールは同じフォルダの `管理者以外停止\\` 直下を使うので、**中を分けて**互いのお願いで止まらないようにする。

    お願い.json                       管理者の PC だけが書く。{"active": true/false, …}
    稼働中\\<PC>__<ID>__<pid>.json     動いている各プロセスが**自分のぶんだけ**書く(1分ごとの知らせと、
                                     お願いへの答え)。管理者の画面はこれを並べて「どの PC が止まったか」を見る

**丸ごと書いてから入れ替える**(書きかけを読ませない)。
**読めないときは「お願いは無い」にしない**(統合 1.2.61 の決まり): 見張りは前の判断のまま、
起動は止めずに記録だけ残す、読めないお願いは上書きしない。

【止め方】
画面の「終了」・stop.bat と同じ道(`server.AppServer.stop`)。ブラウザ版も、デスクトップ版も同じ。
"""
from __future__ import annotations

import getpass
import json
import os
import platform
import re
import threading
import time
import unicodedata
import uuid
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Callable, Iterable, Optional

from common.logging_utils import get_logger

log = get_logger("coil_packing_tools", "remote_stop")

FOLDER = "管理者以外停止"
TOOL_FOLDER = "コイル梱包ツール"
REQUEST_NAME = "お願い.json"
PRESENCE_DIR = "稼働中"
ACCESS_TABLE = "アクセス権限"
ADMINISTRATOR = "Administrator"

# 保存していない入力・処理中の PC を待つ長さ(秒)。過ぎたら強制で止める(現場の決め)
GRACE_SEC = 300
# 何も残っていない PC でも、帯を見せてから止めるまでの間(秒)。画面の問い合わせ(5秒ごと)より十分長く
NOTICE_SEC = 30
# 見張りの間隔(秒)と、「動いています」の知らせの間隔(秒)
POLL_SEC = 5.0
BEAT_SEC = 60.0
# 知らせがこれより古ければ、管理者の画面で「応答なし」
STALE_SEC = 180.0
# 起動を止めたとき、理由の画面を出しておく長さ(秒)。過ぎたら終わる
BLOCK_SHOW_SEC = 60.0
# たまたま読めない(書いている最中・掴まれている)ときの読み直し(秒)
READ_RETRY_SEC = (0.2, 0.5, 1.0)
# 統合画面がこの間に問い合わせてきていれば「開いている」。開いている画面の答えがこれより古ければ待つ
PAGE_ALIVE_SEC = 30.0
REPORT_FRESH_SEC = 20.0
WAITING_SCREEN = "統合画面の確認(保存していない入力があるか)を待っています"

# 各プロセスの答え
STATE_RUNNING = "running"        # 動いている(お願いは無い)
STATE_NOTICE = "notice"          # お願いを受けた。帯を出している(何も残っていなければまもなく止まる)
STATE_WAITING = "waiting"        # 保存していない入力・処理中。利用者を待っている
STATE_STOPPED = "stopped"        # 止まった(利用者が終えた・何も残っていなかった)
STATE_FORCED = "forced"          # 5分たったので強制で止めた
STATE_LABELS = {
    STATE_RUNNING: "動いています",
    STATE_NOTICE: "お願いを受けました(まもなく止まります)",
    STATE_WAITING: "利用者を待っています",
    STATE_STOPPED: "止まりました",
    STATE_FORCED: "強制で止めました",
}

# 断りの種類(文言から推し量らない)
REFUSE_PASSWORD = "wrong_password"
REFUSE_NOT_ADMIN = "not_administrator"
REFUSE_RULES = "rules_unreadable"
REFUSE_NO_FOLDER = "no_folder"
REFUSE_UNREADABLE = "unreadable"
REFUSE_ALREADY = "already_active"
REFUSE_NOT_ACTIVE = "not_active"
REFUSE_WRITE = "write_failed"


class Unreadable(Exception):
    """お願いのファイルが読めない(届かない・壊れている)。「無い」とは別。"""


class RulesUnreadable(Exception):
    """アクセス権限を読めない。「Administrator ではない」とは別(理由を分けて断る)。"""


# ------------------------------------------------------------------
# 置き場所・この端末
# ------------------------------------------------------------------
def master_dir() -> Path:
    """共有の梱包資材マスタのフォルダ(資材計算の設定。配布設定・このPCの設定で変わる)。"""
    from modules.packing_material_calculation.coil_tool import config
    return config.master_db_dir()


def folder() -> Path:
    return master_dir() / FOLDER / TOOL_FOLDER


def request_path() -> Path:
    return folder() / REQUEST_NAME


def presence_dir() -> Path:
    return folder() / PRESENCE_DIR


def _reachable() -> bool:
    try:
        return master_dir().is_dir()
    except OSError:
        return False


def identity() -> tuple[str, str]:
    """(PC名, ログインID)。Windows の COMPUTERNAME / USERNAME(開発機は一般的な手段)。"""
    pc = (os.environ.get("COMPUTERNAME") or platform.node() or "").strip()
    login = (os.environ.get("USERNAME") or "").strip()
    if not login:
        try:
            login = getpass.getuser().strip()
        except Exception:                     # noqa: BLE001 - 環境しだいで落ちる
            login = ""
    return pc, login


def same(a: str, b: str) -> bool:
    """Windows の PC 名・ログイン ID は大文字小文字を区別しない。空どうしは同じとしない。"""
    a, b = (a or "").strip(), (b or "").strip()
    return bool(a) and a.casefold() == b.casefold()


def _safe(text: str) -> str:
    return re.sub(r'[\\/:*?"<>|\s]+', "_", text or "") or "不明"


# ------------------------------------------------------------------
# 読み書き
# ------------------------------------------------------------------
def _read_once(path: Path) -> Optional[dict]:
    try:
        text = path.read_text(encoding="utf-8-sig")
    except FileNotFoundError:
        return None
    data = json.loads(text)
    if not isinstance(data, dict):
        raise ValueError("中身が JSON のまとまりではありません")
    return data


def read_json(path: Path, *, retry: bool = True) -> Optional[dict]:
    """読む。**無いと確かめられたときだけ** None。読めなければ `Unreadable`(たまたまなら読み直す)。"""
    waits = READ_RETRY_SEC if retry else ()
    last: Optional[BaseException] = None
    for wait in (0.0, *waits):
        if wait:
            time.sleep(wait)
        try:
            return _read_once(path)
        except (OSError, ValueError) as exc:
            last = exc
    raise Unreadable(f"{path}: {last}")


def write_json(path: Path, data: dict, *, warn: bool = True) -> bool:
    """丸ごと書いてから入れ替える(書きかけを読ませない)。書けたら True。"""
    tmp = path.with_name(f"{path.name}.{os.getpid()}.tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        os.replace(tmp, path)
        return True
    except OSError as exc:
        try:
            tmp.unlink(missing_ok=True)
        except OSError:
            pass
        (log.warning if warn else log.debug)("管理者以外停止のファイルを書けませんでした(%s): %s", path, exc)
        return False


def read_request(*, retry: bool = True) -> Optional[dict]:
    """いまのお願い。**出ていなければ None**(無い・解除済み)。読めなければ `Unreadable`。"""
    data = read_json(request_path(), retry=retry)
    if data is None or not data.get("active"):
        return None
    return data


def describe_request(req: dict) -> str:
    """「誰が・いつ・何のために」の一言。"""
    try:
        when = time.strftime("%m/%d %H:%M", time.localtime(float(req.get("created_at") or 0)))
    except (TypeError, ValueError, OverflowError, OSError):
        when = "?"
    who = f"{req.get('by_pc') or '?'}({req.get('by_login') or '?'})"
    reason = req.get("reason") or "更新のため"
    return f"{who} が {when} に出しました: {reason}"


# ------------------------------------------------------------------
# お願いを出せるか: 管理者パスワード かつ Administrator
# ------------------------------------------------------------------
_TOKEN_SEP = re.compile(r"[\s,、;；|｜/／]+")


def _norm(text: Any) -> str:
    return unicodedata.normalize("NFKC", "" if text is None else str(text)).strip()


def has_administrator(values: Iterable[Any]) -> bool:
    """権限の欄に ``Administrator`` があるか(「Administrator、L-1」のように並べてもよい。幅・大小は問わない)。"""
    return any(token.casefold() == ADMINISTRATOR.casefold()
               for value in values for token in _TOKEN_SEP.split(_norm(value)) if token)


def _enabled(value: Any) -> bool:
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return value != 0
    return _norm(value).casefold() not in ("", "0", "false", "no", "いいえ", "無効", "off")


def rule_matches(row: dict, pc: str, login: str) -> bool:
    """アクセス権限の1行がこの端末に効くか。**ログインID も PC名 も空の行は誰にも効かせない**
    (全員に効いてしまい、権限を設ける意味が消える。業務ツールと同じ)。"""
    if not _enabled(row.get("有効")):
        return False
    rid, rpc = _norm(row.get("ログインID")), _norm(row.get("PC名"))
    if not rid and not rpc:
        return False
    if rid and not same(rid, login):
        return False
    if rpc and not same(rpc, pc):
        return False
    return True


def read_access_rows() -> list[dict]:
    """共有の梱包資材マスタのアクセス権限(読み取り専用)。**表が無ければ空**。読めなければ `RulesUnreadable`。"""
    from modules.packing_material_calculation.coil_tool import data_sync, source_db
    path = data_sync.find_master_db()
    if path is None:
        raise RulesUnreadable(f"梱包資材マスタのファイルが見つかりません: {master_dir()}")
    last: Optional[BaseException] = None
    for wait in (0.0, *READ_RETRY_SEC):
        if wait:
            time.sleep(wait)
        try:
            return source_db.read_query(
                path, f'SELECT "ログインID", "PC名", "権限", "有効" FROM "{ACCESS_TABLE}"')
        except Exception as exc:              # noqa: BLE001 - 読めない理由はさまざま
            if "no such table" in str(exc).lower():
                return []
            last = exc
    raise RulesUnreadable(str(last))


def is_administrator(pc: str = "", login: str = "",
                     rows: Optional[list[dict]] = None) -> bool:
    """この端末に効く行の **権限** に Administrator があるか。読めなければ `RulesUnreadable`。"""
    if not pc and not login:
        pc, login = identity()
    rows = read_access_rows() if rows is None else rows
    return has_administrator(r.get("権限") for r in rows if rule_matches(r, pc, login))


def _password_ok(password: str) -> bool:
    """梱包明細の管理者パスワード(全ラインで共有。ログの設定と同じ)。照合できなければ通さない。"""
    try:
        from modules.packing_details.meisai import admin_password
        return bool(admin_password.verify(str(password or "")))
    except Exception:                         # noqa: BLE001
        log.exception("管理者パスワードを照合できませんでした")
        return False


@dataclass
class Result:
    ok: bool
    message: str
    reason: str = ""

    def to_dict(self) -> dict:
        return {"ok": self.ok, "message": self.message, "reason": self.reason}


def can_issue(password: str) -> Result:
    """お願いを出せるか(管理者パスワード **かつ** この端末の権限に Administrator)。"""
    if not _password_ok(password):
        return Result(False, "管理者パスワードが違います。", REFUSE_PASSWORD)
    pc, login = identity()
    try:
        admin = is_administrator(pc, login)
    except RulesUnreadable as exc:
        return Result(False, f"アクセス権限を読めないので確かめられません(共有の梱包資材マスタ): {exc}",
                      REFUSE_RULES)
    if not admin:
        return Result(False, f"この端末({pc} / {login})のアクセス権限に Administrator がありません。"
                      "管理者以外停止モードは Administrator の端末からだけ使えます。", REFUSE_NOT_ADMIN)
    if not _reachable():
        return Result(False, f"共有の梱包資材マスタのフォルダに届きません: {master_dir()}"
                      "(お願いはそのフォルダに置きます)", REFUSE_NO_FOLDER)
    return Result(True, "")


def issue(password: str, reason: str = "") -> Result:
    """お願いを出す。**出ているお願いがあれば出さない**(読めなければ出さない)。"""
    allowed = can_issue(password)
    if not allowed.ok:
        return allowed
    try:
        current = read_request()
    except Unreadable as exc:
        return Result(False, f"いまのお願いを読めないので出しません(上書きしないため): {exc}", REFUSE_UNREADABLE)
    if current is not None:
        return Result(False, "もうお願いが出ています。" + describe_request(current), REFUSE_ALREADY)
    pc, login = identity()
    data = {
        "active": True, "id": uuid.uuid4().hex, "by_pc": pc, "by_login": login,
        "reason": (reason or "").strip()[:200] or "更新のため",
        "created_at": time.time(), "grace_sec": GRACE_SEC, "version": _version(),
    }
    if not write_json(request_path(), data):
        return Result(False, f"お願いを書けませんでした: {request_path()}", REFUSE_WRITE)
    log.warning("管理者以外停止のお願いを出しました: %s", describe_request(data))
    _forget_old_answers()
    return Result(True, "お願いを出しました。ほかの PC は保存していない入力・処理中が無ければまもなく止まり、"
                  f"残っていても {GRACE_SEC // 60}分で強制で止まります。")


def clear(password: str) -> Result:
    """解除する(更新が終わった)。ほかの PC がまた起動できるようになる。"""
    allowed = can_issue(password)
    if not allowed.ok:
        return allowed
    try:
        current = read_request()
    except Unreadable as exc:
        return Result(False, f"いまのお願いを読めないので解除しません: {exc}", REFUSE_UNREADABLE)
    if current is None:
        return Result(False, "お願いは出ていません。", REFUSE_NOT_ACTIVE)
    pc, login = identity()
    done = {**current, "active": False, "cleared_at": time.time(),
            "cleared_by_pc": pc, "cleared_by_login": login}
    if not write_json(request_path(), done):
        return Result(False, f"解除を書けませんでした: {request_path()}", REFUSE_WRITE)
    log.warning("管理者以外停止を解除しました(%s / %s)。ほかの PC も起動できます", pc, login)
    _forget_old_answers()
    return Result(True, "解除しました。ほかの PC も起動できます。")


def _forget_old_answers() -> None:
    """止まった・古い答えのファイルを片付ける(**読めたものだけ**。読めないものは消さない)。"""
    try:
        files = list(presence_dir().glob("*.json"))
    except OSError:
        return
    now = time.time()
    for path in files:
        try:
            data = read_json(path, retry=False)
        except Unreadable:
            continue
        if data is None:
            continue
        old = now - float(data.get("beat_at") or 0) > STALE_SEC * 10
        if data.get("state") in (STATE_STOPPED, STATE_FORCED) or old:
            try:
                path.unlink()
            except OSError:
                pass


def _version() -> str:
    try:
        from common import app_config
        return app_config.version()
    except Exception:                         # noqa: BLE001
        return ""


@dataclass
class Overview:
    """管理者の画面に出すもの。"""

    available: bool
    why: str = ""
    request: Optional[dict] = None
    request_text: str = ""
    unreadable: str = ""
    pcs: list = field(default_factory=list)
    me_pc: str = ""
    me_login: str = ""
    folder: str = ""

    def to_dict(self) -> dict:
        return {"available": self.available, "why": self.why, "request": self.request,
                "request_text": self.request_text, "unreadable": self.unreadable,
                "pcs": self.pcs, "me_pc": self.me_pc, "me_login": self.me_login,
                "folder": self.folder, "grace_sec": GRACE_SEC}


def overview() -> Overview:
    """いまのお願いと、各 PC の様子。"""
    pc, login = identity()
    if not _reachable():
        return Overview(False, f"共有の梱包資材マスタのフォルダに届きません: {master_dir()}",
                        me_pc=pc, me_login=login)
    view = Overview(True, me_pc=pc, me_login=login, folder=str(folder()))
    try:
        view.request = read_request(retry=False)
    except Unreadable as exc:
        view.unreadable = f"お願いのファイルを読めません: {exc}"
    if view.request:
        view.request_text = describe_request(view.request)
    rid = (view.request or {}).get("id")
    by_pc = (view.request or {}).get("by_pc", "")
    now = time.time()
    try:
        files = sorted(presence_dir().glob("*.json"))
    except OSError:
        files = []
    for path in files:
        try:
            data = read_json(path, retry=False)
        except Unreadable:
            view.pcs.append({"pc": path.stem, "login": "", "state": "", "label": "読めません",
                             "detail": "", "beat": "", "stale": True})
            continue
        if data is None:
            continue
        state = data.get("state") or STATE_RUNNING
        answered = bool(rid) and data.get("request_id") == rid
        finished = state in (STATE_STOPPED, STATE_FORCED)
        stale = (not finished) and now - float(data.get("beat_at") or 0) > STALE_SEC
        if finished and not answered:
            continue                              # 前のお願いで止まったもの
        label = STATE_LABELS.get(state, state)
        issuer = bool(rid) and same(data.get("pc", ""), by_pc)
        if stale:
            label = "応答なし(知らせが途絶えています)"
        elif rid and not answered and not issuer:
            label = "まだお願いを受けていません"
        if issuer:
            label += "(お願いを出した PC。止めません)"
        view.pcs.append({
            "pc": data.get("pc", ""), "login": data.get("login", ""), "kind": data.get("kind", ""),
            "version": data.get("version", ""), "state": state, "label": label,
            "detail": data.get("detail", ""), "stale": stale,
            "beat": time.strftime("%H:%M:%S", time.localtime(float(data.get("beat_at") or 0))),
        })
    return view


# ------------------------------------------------------------------
# 起動: 更新中は起動させない / stop.bat と同時に使わない
# ------------------------------------------------------------------
def startup_block() -> Optional[str]:
    """起動してよいか。止めるなら理由の文、よければ None。

    お願いを出した PC は止めない(更新したあと、その PC で起動して確かめ、解除する)。
    **読めないときは止めない**(共有に届かないだけで全員が起動できなくなるのを避ける)。記録には残す。
    """
    try:
        if not _reachable():
            return None
        req = read_request()
    except Unreadable as exc:
        log.warning("管理者以外停止のお願いを確かめられないまま起動します: %s", exc)
        return None
    except Exception as exc:                  # noqa: BLE001 - 起動の判定で落ちない
        log.warning("管理者以外停止のお願いを確かめられませんでした: %s", exc)
        return None
    if req is None:
        return None
    pc, _login = identity()
    if same(pc, req.get("by_pc", "")):
        log.info("管理者以外停止のお願いが出ていますが、お願いを出した PC なので起動します: %s",
                 describe_request(req))
        return None
    log.warning("管理者以外停止モード中のため起動しません: %s", describe_request(req))
    return ("管理者が更新のためコイル梱包ツールを止めています(管理者以外停止モード)。"
            "更新が終わるまで起動できません。"
            f"\n{describe_request(req)}"
            f"\nこの画面は {int(BLOCK_SHOW_SEC)}秒後に閉じます。しばらくしてから起動し直してください。"
            f"\n(管理者へ: お願いを出した PC で解除できないときは、{request_path()} を"
            "別の名前に変えると解除できます)")


def stop_bat_refusal() -> Optional[str]:
    """stop.bat(ランチャーの「全 PC で止める」)を断るなら1行目の理由。**同時には使わない**(現場の決め)。

    読めなければ断らない(止める道を閉じない)。
    """
    try:
        if not _reachable():
            return None
        req = read_request()
    except Exception:                         # noqa: BLE001 - Unreadable を含む
        return None
    if req is None:
        return None
    return ("管理者以外停止モード中のため、stop.bat(ランチャーの「全 PC で止める」)は"
            f"使いません。解除してから使ってください({describe_request(req)})")


# ------------------------------------------------------------------
# 画面の答え(保存していない入力)。統合画面が5秒ごとに問い合わせ、お願いが出ていれば答える
# ------------------------------------------------------------------
_pages_lock = threading.Lock()
_pages: dict[str, float] = {}                         # 画面の名乗り → 最後に問い合わせた時刻
_reports: dict[str, tuple[float, list, list]] = {}    # 画面の名乗り → (時刻, 名前, 打っていた値)


def page_seen(page: str, now: Optional[float] = None) -> None:
    if not page:
        return
    with _pages_lock:
        _pages[page[:64]] = time.monotonic() if now is None else now


def page_gone(page: str) -> None:
    """画面を閉じた(閉じる前の確認は画面が済ませている)。もう待たない。"""
    if not page:
        return
    with _pages_lock:
        _pages.pop(page[:64], None)
        _reports.pop(page[:64], None)


def report(page: str, labels: Iterable[Any], fields: Iterable[Any] = (),
           now: Optional[float] = None) -> None:
    """画面の答え: 保存していない入力の名前と、打っていた値(強制で止めるときの控え)。"""
    if not page:
        return
    names = [str(x)[:200] for x in list(labels or [])[:50]]
    values = []
    for item in list(fields or [])[:200]:
        if isinstance(item, dict):
            values.append({k: str(item.get(k, ""))[:300] for k in ("screen", "label", "value")})
    at = time.monotonic() if now is None else now
    with _pages_lock:
        _pages[page[:64]] = at
        _reports[page[:64]] = (at, names, values)


def screens_left(now: Optional[float] = None) -> list[str]:
    """開いている画面の、保存していない入力。**答えが新しくない画面があれば「待っている」を返す**。"""
    now = time.monotonic() if now is None else now
    out: list[str] = []
    waiting = False
    with _pages_lock:
        for page, at in list(_pages.items()):
            if now - at > PAGE_ALIVE_SEC:
                continue
            rep = _reports.get(page)
            if rep is None or now - rep[0] > REPORT_FRESH_SEC:
                waiting = True
                continue
            for name in rep[1]:
                if name not in out:
                    out.append(name)
    if waiting and not out:
        out.append(WAITING_SCREEN)
    return out


def screen_values() -> tuple[list, list]:
    """画面が最後に答えた (名前, 打っていた値)。強制で止める前の控え用(古くても使う)。"""
    names: list = []
    values: list = []
    with _pages_lock:
        for _at, labels, fields in _reports.values():
            names += [n for n in labels if n not in names]
            values += fields
    return names, values


def _reset_pages() -> None:
    with _pages_lock:
        _pages.clear()
        _reports.clear()


# ------------------------------------------------------------------
# 受け側: 動いているプロセスの見張り
# ------------------------------------------------------------------
class Watcher:
    """お願いを見張り、帯を出し、止める。プロセスに1つ。"""

    def __init__(self, stop: Callable[[], None], unsaved: Callable[[], list],
                 busy: Callable[[], list], *, kind: str = "browser", mode: str = "",
                 before_force: Optional[Callable[[], list]] = None,
                 clock: Callable[[], float] = time.time) -> None:
        self._stop = stop
        self._unsaved = unsaved
        self._busy = busy
        self._before_force = before_force
        self._clock = clock
        self.kind = kind
        self.mode = mode
        self.pc, self.login = identity()
        self.started_at = clock()
        self.state = STATE_RUNNING
        self.detail = ""
        self.request: Optional[dict] = None
        self.received_at = 0.0
        self.deadline = 0.0
        self.notice_until = 0.0
        self.stopping = False
        self._last_beat = 0.0
        self._beat_failing = False
        self._stop_called = False
        self._lock = threading.Lock()
        self._thread: Optional[threading.Thread] = None
        self._halt = threading.Event()

    # --- 自分の知らせ ---------------------------------------------------
    def presence_path(self) -> Path:
        return presence_dir() / f"{_safe(self.pc)}__{_safe(self.login)}__{os.getpid()}.json"

    def _beat(self, *, force: bool = False) -> None:
        now = self._clock()
        if not force and now - self._last_beat < BEAT_SEC:
            return
        if not _reachable():
            return
        self._last_beat = now
        ok = write_json(self.presence_path(), {
            "pc": self.pc, "login": self.login, "pid": os.getpid(), "kind": self.kind, "mode": self.mode,
            "version": _version(), "started_at": self.started_at, "beat_at": now,
            "state": self.state, "detail": self.detail,
            "request_id": (self.request or {}).get("id", ""),
        }, warn=not self._beat_failing)
        # 書けないあいだは1分ごとに記録をためない(最初の1回と、直ったときだけ)
        if not ok and not self._beat_failing:
            log.warning("管理者以外停止の知らせを共有に書けません(書けるようになるまで記録しません): %s",
                        self.presence_path())
        elif ok and self._beat_failing:
            log.info("管理者以外停止の知らせを共有に書けるようになりました")
        self._beat_failing = not ok

    def leave(self) -> None:
        """終わるとき。止まった答えは残す(管理者が見る)、ふだんの終わりなら消す。

        **お願いを受けている最中に利用者が自分で終えた**(帯の「終了」・窓を閉じた)ときも「止まりました」を
        残す(消すと管理者の一覧から消えて、止まったのか分からない)。利用者は「終了」で保存していない入力を
        確かめてから終えているので、ここでは控えを書かない。
        """
        self._halt.set()
        with self._lock:
            if self.state in (STATE_STOPPED, STATE_FORCED):
                return
            if self.request is not None:
                self.state = STATE_STOPPED
                self.detail = "利用者が終了しました"
                self._beat(force=True)
                return
        try:
            self.presence_path().unlink(missing_ok=True)
        except OSError:
            pass

    # --- 見張り ---------------------------------------------------------
    def start(self) -> None:
        if self._thread is not None:
            return
        self._thread = threading.Thread(target=self._loop, name="remote-stop", daemon=True)
        self._thread.start()

    def _loop(self) -> None:
        while not self._halt.is_set():
            try:
                self.tick()
            except Exception:                 # noqa: BLE001 - 見張りは止めない
                log.debug("管理者以外停止の見張りで失敗しました(次で試します)", exc_info=True)
            if self.stopping:
                return
            self._halt.wait(POLL_SEC)

    def tick(self) -> None:
        """1回ぶん見る。**止めるのは錠の外で**(止めると終わりの後始末 `leave` が同じ錠を取りに来る)。"""
        with self._lock:
            self._tick()
            stop_now = self.stopping and not self._stop_called
            if stop_now:
                self._stop_called = True
        if stop_now:
            try:
                self._stop()
            except Exception:                 # noqa: BLE001
                log.exception("管理者以外停止で止められませんでした")

    def _tick(self) -> None:
        if self.stopping:
            return
        now = self._clock()
        try:
            req = read_request() if _reachable() else self.request
        except Unreadable as exc:
            # **読めないことを「お願いは無い」にしない。** いまの判断のまま待つ(期限は進める)
            log.debug("管理者以外停止のお願いを読めません(前の判断のまま): %s", exc)
            req = self.request
        if req is not None and same(self.pc, req.get("by_pc", "")):
            req = None                            # お願いを出した PC は止めない
        if req is None:
            if self.request is not None:
                log.info("管理者以外停止のお願いが取り下げられました。止めずに続けます")
                user_note("管理者が停止のお願いを取り下げました。そのまま使えます。")
            self.request = None
            self.state, self.detail = STATE_RUNNING, ""
            self._beat()
            return
        if self.request is None or self.request.get("id") != req.get("id"):
            self.request = req
            self.received_at = now
            try:
                grace = float(req.get("grace_sec") or GRACE_SEC)
            except (TypeError, ValueError):
                grace = GRACE_SEC
            self.deadline = now + grace
            self.notice_until = now + NOTICE_SEC
            log.warning("管理者以外停止のお願いを受けました: %s", describe_request(req))
        left = self._left()
        # 何も残っていなければ止める。帯を見せる間(NOTICE_SEC)は待つが、利用者を待っていた(帯はもう
        # 見えている)あとに保存・終了して何も残らなくなったなら、すぐ止める
        if not left and (now >= self.notice_until or self.state == STATE_WAITING):
            self._finish(STATE_STOPPED, "保存していない入力・実行中の処理が無いので止めました")
            return
        if now >= self.deadline:
            saved = self._save_before_force()
            detail = f"{int(round((self.deadline - self.received_at) / 60))}分たっても止まらないので強制で止めました"
            if left:
                detail += f"(残っていたもの: {'、'.join(left)})"
            if saved:
                detail += f"。控えに残しました: {'、'.join(saved)}"
            self._finish(STATE_FORCED, detail)
            return
        self.state = STATE_WAITING if left else STATE_NOTICE
        self.detail = "、".join(left)
        self._beat(force=True)

    def _left(self) -> list[str]:
        names = [str(x) for x in (self._unsaved() or [])]
        names += [str(job) for job in (self._busy() or [])]
        return names

    def _save_before_force(self) -> list[str]:
        if self._before_force is None:
            return []
        try:
            return [str(x) for x in (self._before_force() or [])]
        except Exception:                     # noqa: BLE001 - 控えに失敗しても止める(期限は守る)
            log.exception("保存していない入力の控えを残せませんでした")
            return []

    def _finish(self, state: str, detail: str) -> None:
        self.state, self.detail = state, detail
        self.stopping = True
        self._beat(force=True)
        log.warning("管理者以外停止: %s", detail)

    # --- 画面へ --------------------------------------------------------
    def screen_state(self) -> Optional[dict]:
        """帯に出すもの。お願いが無ければ None。"""
        req = self.request
        if req is None:
            return None
        return {
            "id": req.get("id", ""),
            "text": describe_request(req),
            "by": f"{req.get('by_pc') or '?'}({req.get('by_login') or '?'})",
            "reason": req.get("reason") or "更新のため",
            "deadline_ms": int(self.deadline * 1000),
            "now_ms": int(self._clock() * 1000),
            "left": self.detail.split("、") if self.detail and self.state == STATE_WAITING else [],
            "state": self.state,
            "stopping": self.stopping,
        }


_watcher: Optional[Watcher] = None
_notes: list[str] = []
_notes_lock = threading.Lock()


def install(stop: Callable[[], None], unsaved: Callable[[], list], busy: Callable[[], list], *,
            kind: str = "browser", mode: str = "",
            before_force: Optional[Callable[[], list]] = None) -> Watcher:
    global _watcher
    if _watcher is not None:
        return _watcher
    _watcher = Watcher(stop, unsaved, busy, kind=kind, mode=mode, before_force=before_force)
    _watcher.start()
    return _watcher


def get() -> Optional[Watcher]:
    return _watcher


def leave() -> None:
    if _watcher is not None:
        _watcher.leave()


def screen_state() -> Optional[dict]:
    return _watcher.screen_state() if _watcher is not None else None


def user_note(message: str) -> None:
    """画面に出す一言。**利用者が「閉じる」を押すまで**出し続ける(画面を開き直しても出る)。"""
    with _notes_lock:
        if message not in _notes:
            _notes.append(message)


def notes() -> list[str]:
    with _notes_lock:
        return list(_notes)


def notes_seen(shown: Iterable[str]) -> None:
    """画面で「閉じる」を押した一言を消す(見せていないものは残す)。"""
    seen = set(shown or [])
    with _notes_lock:
        _notes[:] = [n for n in _notes if n not in seen]


def reset() -> None:
    """試験用。"""
    global _watcher
    if _watcher is not None:
        _watcher._halt.set()
    _watcher = None
    with _notes_lock:
        _notes.clear()
    _reset_pages()
