r"""エラーの記録 ── なぜなぜ分析のために「そのときの事実」を1件1ファイルで残す

現場の指摘: エラー等の後追いが現状できない。ログを残し、なぜなぜで分析できるように
しておいてほしい(統合 1.0.12)。

ログ(1日1ファイル)には何でも流れてくるので、あとから「13日の 10時ごろのあのエラー」を
探すのは骨が折れる。そこで、**エラーが起きたらその場で1件のファイル**を作る:

    <ログの出力パス>\coilpacking\incidents\coilpacking_なぜなぜ_<PC名>_E20261001-104512-3F9A.md
    (統合 1.2.53 から。日報複合ツールの決まり「<ツールID>_<種類>_<PC名>_…」に合わせた。
     前は 1.2.50 まで E….md、1.2.51〜1.2.52 は コイル梱包ツール_<PC名>_E….md。前の物も読めて、
     残す日数を過ぎたら消える。もう書かない)

中身はなぜなぜ分析の順に並べる:

    1. 何が起きたか(現象)  画面に出した文言・いつ・どのPC・どの機能のどの操作
    2. 直接の原因          例外とその場所(プログラムが止まった行)・送られた値
    3. エラーまでの流れ      そのときのログの直前の行(同じ要求の行には印)
    4. そのときの状態        版・PC・Python・ログの置き場所
    5. なぜなぜ分析(記入用) なぜ1〜5・対策・再発防止の欄

**エラー番号は画面にも出す**(「処理中にエラーが起きました(エラー番号 E…)」)。
現場の人がその番号を伝えれば、このファイルにすぐ辿り着ける。

【決まり】
- ここで例外を出さない(記録に失敗しても、元の処理の断り方は変えない)
- パスワード・合言葉・トークンは書かない(`***` に伏せる)
- 同じ画面のエラーが続けて来たら、1件にまとめて回数を足す(`DEDUP_SEC`)
"""
from __future__ import annotations

import json
import logging
import os
import platform
import re
import secrets
import sys
import threading
import time
import traceback
from datetime import date, datetime, timedelta
from pathlib import Path
from typing import Any, Optional

from . import logging_utils

log = logging.getLogger("coil_packing_tools.incidents")

FOLDER = "incidents"
#: ファイル名の種類(`<ツールID>_なぜなぜ_<PC名>_<番号>.md`)
KIND = "なぜなぜ"
ID_RE = re.compile(r"^E\d{8}-\d{6}-[0-9A-F]{4}$")

#: 同じ内容の画面のエラーを1件にまとめる時間(秒)
DEDUP_SEC = 600
#: 「エラーまでの流れ」に写す行数
TRAIL_LINES = 200
#: 送られた値を写す上限(文字)
VALUE_LIMIT = 4000

#: 伏せる鍵(小文字で含んでいれば伏せる)
SECRET_WORDS = ("password", "passwd", "token", "secret", "パスワード", "合言葉")

KIND_LABEL = {
    "server": "サーバの処理が止まった",
    "screen": "画面(ブラウザ)のエラー",
    "write": "書き込めなかった",
}

_lock = threading.Lock()
_recent: dict[str, tuple[str, float]] = {}       # まとめる鍵 → (番号, 最後の時刻)


def folder() -> Path:
    return logging_utils.log_dir() / FOLDER


_ID = r"(E\d{8}-\d{6}-[0-9A-F]{4})"


def file_name(eid: str) -> str:
    """1件のファイル名: `<ツールID>_なぜなぜ_<PC名>_<番号>.md`(取り出しても、どのツール・どの PC の記録か分かる)。"""
    return f"{logging_utils.file_prefix(KIND)}_{eid}.md"


def _own_id(name: str) -> str:
    """この端末の、いまの名前の記録なら番号。"""
    m = re.match(rf"^{re.escape(logging_utils.file_prefix(KIND))}_{_ID}\.md$", name)
    return m.group(1) if m else ""


def _legacy_id(name: str) -> str:
    """前の名前(`E….md`・`コイル梱包ツール_<この PC>_E….md`)の記録なら番号。"""
    m = (re.match(rf"^{_ID}\.md$", name)
         or re.match(rf"^{re.escape(logging_utils.OLD_TOOL_NAME)}_{re.escape(logging_utils.pc_name())}_{_ID}\.md$",
                     name))
    return m.group(1) if m else ""


def _places() -> list[tuple[Path, Any]]:
    """読む・片付ける場所と、そこで自分の物と見なす名前。いまの場所が先、前の場所があと(もう書かない)。"""
    current = list(dict.fromkeys([folder(), logging_utils.default_log_dir() / FOLDER]))
    return ([(p, _own_id) for p in current]
            + [(p / FOLDER, _legacy_id) for p in logging_utils.legacy_dirs()])


def path_for(eid: str) -> Path:
    """番号の記録のファイル。いまの名前 → 前の場所・前の名前 → 同じツールのほかの PC の物、の順に探す。

    無ければいまの名前(新しく書く場所)。
    """
    new = folder() / file_name(eid)
    if new.exists():
        return new
    for place, match in _places():
        try:
            names = os.listdir(place)
        except OSError:
            continue
        for name in names:
            if match(name) == eid:
                return place / name
    prefix = f"{logging_utils.tool_id()}_{KIND}_"
    for hit in sorted(folder().glob(f"{prefix}*_{eid}.md")):
        return hit
    return new


def new_id(now: Optional[datetime] = None) -> str:
    now = now or datetime.now()
    return f"E{now:%Y%m%d-%H%M%S}-{secrets.token_hex(2).upper()}"


# ------------------------------------------------------------------
# 伏せる
# ------------------------------------------------------------------
def is_secret(key: str) -> bool:
    k = str(key).lower()
    return any(w in k for w in SECRET_WORDS)


def mask(value: Any) -> Any:
    """パスワード・合言葉・トークンを `***` に伏せる(入れ子も)。"""
    if isinstance(value, dict):
        return {k: ("***" if is_secret(k) and v not in ("", None) else mask(v))
                for k, v in value.items()}
    if isinstance(value, list):
        return [mask(v) for v in value]
    return value


def mask_url(text: str) -> str:
    """URL の中のトークン(`t=` など)を伏せる。"""
    return re.sub(r"([?&](?:t|token)=)[^&#\s]*", r"\1***", str(text or ""))


def _short(value: Any, limit: int = VALUE_LIMIT) -> str:
    try:
        text = json.dumps(mask(value), ensure_ascii=False, indent=1, default=str)
    except Exception:                               # noqa: BLE001 - 書けなければ文字で
        text = str(value)
    if len(text) > limit:
        text = text[:limit] + f"\n…(以下 {len(text) - limit} 文字は省きました)"
    return text


# ------------------------------------------------------------------
# そのときの状態
# ------------------------------------------------------------------
def environment() -> dict:
    """版・PC・Python・ログの置き場所。**例外を出さない。**"""
    env: dict = {
        "PC": logging_utils.pc_name(),
        "Windows の利用者": os.environ.get("USERNAME") or os.environ.get("USER") or "",
        "Python": sys.version.split()[0],
        "OS": platform.platform(),
        "ログの置き場所": str(logging_utils.log_dir()),
    }
    try:
        from flask import current_app, has_app_context
        if has_app_context():
            conf = current_app.config
            env["版"] = " / ".join(f"{k} {v}" for k, v in (conf.get("VERSIONS") or {}).items())
            env["起動してから"] = f"{round(time.time() - conf.get('STARTED_AT', time.time()))} 秒"
            env["ポート"] = conf.get("PORT", "")
    except Exception:                               # noqa: BLE001
        pass
    return env


def request_facts() -> dict:
    """要求の中なら、その要求の事実(どの画面から・何を・どの値で)。"""
    try:
        from flask import g, has_request_context, request
    except Exception:                               # noqa: BLE001
        return {}
    if not has_request_context():
        return {}
    facts: dict = {
        "要求": f"{request.method} {mask_url(request.full_path.rstrip('?'))}",
        "要求の印": getattr(g, "request_id", ""),
        "機能": _module_label(request.path),
        "画面の名乗り": (request.headers.get("X-Tool-Screen") or request.args.get("screen")
                       or ""),
        "開いていた画面": mask_url(request.headers.get("Referer", "")),
        "ブラウザ": request.headers.get("User-Agent", "")[:200],
    }
    body = request.get_json(silent=True) if request.is_json else None
    if body is None and request.form:
        body = request.form.to_dict()
    if body not in (None, {}, ""):
        facts["送られた値"] = body
    return {k: v for k, v in facts.items() if v not in ("", None)}


def _module_label(path: str) -> str:
    for prefix, label in (("/details", "梱包明細"), ("/pena", "ペナラベル"),
                          ("/material", "資材計算")):
        if path == prefix or path.startswith(prefix + "/"):
            return label
    return "統合画面"


# ------------------------------------------------------------------
# 書く
# ------------------------------------------------------------------
def record(kind: str, title: str, *, shown: str = "", exc: Optional[BaseException] = None,
           facts: Optional[dict] = None, dedup_key: str = "", eid: str = "") -> str:
    """1件書いて、エラー番号を返す。**例外を出さない。**

    `kind`  … server(サーバの処理が止まった)/ screen(画面のエラー)/ write(書けなかった)
    `title` … 1行の要約(一覧に出す)
    `shown` … 画面に出した(出す)文言
    `exc`   … 例外(あれば、場所とトレースバックを写す)
    `facts` … 要求の事実など。省けば、要求の中なら `request_facts()` を使う
    `dedup_key` … 同じものを1件にまとめる鍵(画面のエラー)
    `eid` … 先に決めた番号(画面に出す文言に番号を入れてから記録するとき。`new_id()`)
    """
    now = time.time()
    if dedup_key:
        with _lock:
            hit = _recent.get(dedup_key)
            if hit and now - hit[1] < DEDUP_SEC:
                _recent[dedup_key] = (hit[0], now)
                _append_repeat(hit[0])
                return hit[0]
    eid = eid if ID_RE.match(eid or "") else new_id()
    if dedup_key:
        with _lock:
            _recent[dedup_key] = (eid, now)
            for k in [k for k, (_i, t) in _recent.items() if now - t > DEDUP_SEC]:
                _recent.pop(k, None)
    facts = facts if facts is not None else request_facts()
    path = folder() / file_name(eid)
    try:
        text = _render(eid, kind, title, shown, exc, facts)
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
        log.error("エラーの記録 %s: %s ── %s", eid, title, path)
    except Exception as err:                        # noqa: BLE001 - 記録できなくても続ける
        log.error("エラーの記録 %s を書けませんでした(%s): %s", eid, err, title)
    return eid


def _append_repeat(eid: str) -> None:
    path = folder() / file_name(eid)                     # いまの場所だけ(前の場所にはもう書かない)
    if not path.exists():
        return
    try:
        with path.open("a", encoding="utf-8") as f:
            f.write(f"- {time.strftime('%Y/%m/%d %H:%M:%S')} 同じエラーがもう一度起きました\n")
    except OSError:
        pass


def _render(eid: str, kind: str, title: str, shown: str, exc: Optional[BaseException],
            facts: dict) -> str:
    now = datetime.now()
    rid = str(facts.get("要求の印", ""))
    lines: list[str] = []
    w = lines.append
    w(f"# エラーの記録 {eid}")
    w("")
    w(f"- 種類: {KIND_LABEL.get(kind, kind)}")
    w(f"- 要約: {title}")
    w("")
    w("## 1. 何が起きたか(現象)")
    w("")
    w(f"- 日時: {now:%Y/%m/%d %H:%M:%S}")
    w(f"- PC: {logging_utils.pc_name()}")
    if facts.get("機能"):
        w(f"- 機能: {facts['機能']}")
    if shown:
        w(f"- 画面に出した文言: {shown}")
    for key in ("要求", "開いていた画面", "画面の名乗り", "画面", "操作"):
        if facts.get(key):
            w(f"- {key}: {facts[key]}")
    w("")
    w("## 2. 直接の原因(プログラムが止まった場所)")
    w("")
    if exc is not None:
        frames = traceback.extract_tb(exc.__traceback__)
        where = frames[-1] if frames else None
        w(f"- 例外: {type(exc).__name__}: {exc}")
        if where is not None:
            w(f"- 場所: {_rel(where.filename)}:{where.lineno}({where.name})")
            if where.line:
                w(f"- その行: `{where.line.strip()}`")
        w("")
        w("```")
        w("".join(traceback.format_exception(type(exc), exc, exc.__traceback__)).rstrip())
        w("```")
    elif facts.get("例外"):
        w(f"- 例外: {facts['例外']}")
        if facts.get("場所"):
            w(f"- 場所: {facts['場所']}")
        if facts.get("スタック"):
            w("")
            w("```")
            w(str(facts["スタック"])[:VALUE_LIMIT])
            w("```")
    else:
        w("- (例外はありません。下の流れを見てください)")
    if "送られた値" in facts:
        w("")
        w("送られた値(パスワード・合言葉・トークンは伏せています):")
        w("")
        w("```json")
        w(_short(facts["送られた値"]))
        w("```")
    w("")
    w(f"## 3. エラーまでの流れ(ログの直前 {TRAIL_LINES} 行まで)")
    w("")
    if rid:
        w(f"`>>` の行はこの操作(要求 {rid})の中の行です。")
        w("")
    w("```")
    for line in logging_utils.recent_lines(TRAIL_LINES):
        w((">> " if rid and f"[要求 {rid}]" in line else "   ") + mask_url(line))
    w("```")
    w("")
    w("## 4. そのときの状態")
    w("")
    for key, value in environment().items():
        w(f"- {key}: {value}")
    extra = {k: v for k, v in facts.items()
             if k not in ("要求", "要求の印", "機能", "画面の名乗り", "開いていた画面", "画面",
                          "操作", "送られた値", "例外", "場所", "スタック")}
    for key, value in extra.items():
        w(f"- {key}: {mask_url(str(value))[:500]}")
    w("")
    w("## 5. なぜなぜ分析(記入用)")
    w("")
    w("1〜4 の事実から、「なぜそうなったか」を5回くり返して、仕組みの原因まで下ります。")
    w("人のせいで止めず、「なぜその間違いができてしまったか」まで書きます。")
    w("")
    for n in range(1, 6):
        w(f"- なぜ{n}: ")
    w("- 真の原因: ")
    w("- 対策(すぐやること): ")
    w("- 再発防止(仕組みで防ぐこと): ")
    w("- 記入者・日付: ")
    w("")
    w("## 経過")
    w("")
    w(f"- {now:%Y/%m/%d %H:%M:%S} 記録しました")
    return "\n".join(lines) + "\n"


def _rel(filename: str) -> str:
    """ソースの場所をアプリのフォルダからの相対で(PC ごとの置き場所の違いを消す)。"""
    try:
        from . import app_config
        return str(Path(filename).resolve().relative_to(Path(app_config.APP_ROOT).resolve()))
    except Exception:                               # noqa: BLE001
        return filename


# ------------------------------------------------------------------
# 読む(画面の一覧と中身)
# ------------------------------------------------------------------
def list_recent(limit: int = 100) -> list[dict]:
    """新しい順。1件 = 番号・日時・種類・要約・機能。"""
    found: dict[str, Path] = {}
    for place, match in _places():                       # いまの場所が先(同じ番号は先に見つけた方)
        try:
            names = os.listdir(place)
        except OSError:
            continue
        for name in names:
            eid = match(name)
            if eid and eid not in found:
                found[eid] = place / name
    out = []
    for eid in sorted(found, reverse=True)[:limit]:      # 番号(日時)の新しい順
        path = found[eid]
        info = {"id": eid, "when": _when(eid), "kind": "", "title": "", "module": "",
                "repeats": 0}
        try:
            with path.open(encoding="utf-8") as f:
                for line in f:
                    if line.startswith("- 種類: ") and not info["kind"]:
                        info["kind"] = line[6:].strip()
                    elif line.startswith("- 要約: ") and not info["title"]:
                        info["title"] = line[6:].strip()
                    elif line.startswith("- 機能: ") and not info["module"]:
                        info["module"] = line[6:].strip()
                    elif "同じエラーがもう一度" in line:
                        info["repeats"] += 1
        except OSError:
            pass
        out.append(info)
    return out


def _when(eid: str) -> str:
    d, t = eid[1:9], eid[10:16]
    return f"{d[:4]}/{d[4:6]}/{d[6:]} {t[:2]}:{t[2:4]}:{t[4:]}"


def read(eid: str) -> Optional[str]:
    """1件の中身。番号の形が違えば `None`(パスを組み立てるので、形を確かめる)。"""
    if not ID_RE.match(str(eid or "")):
        return None
    try:
        return path_for(eid).read_text(encoding="utf-8")
    except OSError:
        return None


def cleanup_old(keep_days: Optional[int] = None, *, today: Optional[date] = None) -> list[str]:
    """残す日数を過ぎた記録を消す。**番号の形のファイルだけ**消す。"""
    from . import local_settings
    if keep_days is not None:
        days = keep_days
    else:
        try:
            days = local_settings.log_keep_days(strict=True)
        except local_settings.Unreadable:
            # 残す日数を確かめられない。既定の 180 日で消すと、長く残す設定の PC で共有の記録を消してしまう
            # (統合 1.2.61。読めないのを「設定が無い」と取り違えない)。今回は消さない
            return []
    limit = (today or date.today()) - timedelta(days=days)
    removed = []
    for target, match in _places():                      # この端末の・このツールの物だけ
        try:
            names = os.listdir(target)
        except OSError:
            continue
        for name in names:
            eid = match(name)
            if not eid:
                continue
            try:
                day = date(int(eid[1:5]), int(eid[5:7]), int(eid[7:9]))
            except ValueError:
                continue
            if day < limit:
                try:
                    (target / name).unlink()
                    removed.append(name)
                except OSError:
                    pass
    return removed
