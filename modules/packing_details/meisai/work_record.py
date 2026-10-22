"""作業の記録の「消えない控え」と、記録に添える素性(VER 0.13.30。ペナラベルと同じ)

出力・出力し直し・修正の取り消し・全クリア・印刷画面・紙面を開いた、を 1 件ごとに
**月ごとのファイルへ 1 行ずつ足す**(JSON Lines。足すだけで、消さない・書き換えない):

    この PC     <ローカル領域>/records/梱包明細記録_YYYY-MM.jsonl      (いつも)
    設定の場所  <記録の保存先フォルダー>/<PC名>/梱包明細記録_YYYY-MM.jsonl
                (設定「記録の保存先フォルダー」= record_dir があるときだけ。PC ごとに
                 ファイルを分けるので、複数のラインが同じファイルを取り合わない)

記録に添える素性(どの記録からも言えるように):

- **ソフトの版**: 統合ツールの版・梱包明細の版。「切り替えの前か後か」「どの版で出したか」を
  記録から言うため
- **PC 名**と**ライン名**(設定「ライン名」。この PC)、ログインID
- テストモードなら「テストモード」の印(`mode`)

書けなくても出力・印刷は止めない(理由はログと記録の画面に出す)。
"""
from __future__ import annotations

import datetime
import json
import re
import threading
from pathlib import Path
from typing import Any, Optional

from . import app_config, config, shared_settings, user_settings
from .logging_utils import get_logger

log = get_logger("work_record")

_lock = threading.Lock()
#: 最後に書けなかった理由(記録の画面に出す)
_problems: dict[str, str] = {"local": "", "shared": ""}

FILE_PREFIX = "梱包明細記録"

#: テストモードのときの印(記録・明細の履歴・編集履歴に残す)
TEST_MODE_MARK = "テストモード"


# ==================================================================
# 設定(この PC)
# ==================================================================
def test_mode() -> bool:
    """テストモードか。**文字の "false" は切**(手で直した設定ファイル)。"""
    return user_settings.get(config.KEY_TEST_MODE) is True


def mode_mark() -> str:
    """テストモードなら「テストモード」。切なら空。"""
    return TEST_MODE_MARK if test_mode() else ""


def line_name() -> str:
    value = user_settings.get(config.KEY_LINE_NAME)
    return value.strip() if isinstance(value, str) else ""


def record_dir_setting() -> str:
    value = user_settings.get(config.KEY_RECORD_DIR)
    return value.strip() if isinstance(value, str) else ""


def pc_folder_name() -> str:
    """PC 名をフォルダの名前に使える形にする。"""
    name = shared_settings.pc_name() or "PC"
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "PC"


# ==================================================================
# 素性
# ==================================================================
def who() -> dict[str, str]:
    """どの PC(ライン)・誰の記録か。テストモードなら ``mode`` に「テストモード」。"""
    out = {"pc": shared_settings.pc_name(), "line": line_name(),
           "user": shared_settings.login_id()}
    if test_mode():
        out["mode"] = TEST_MODE_MARK
    return out


def versions() -> dict[str, str]:
    """ソフトの版(統合ツール・梱包明細)。"""
    out = {"details": app_config.version()}
    try:
        from common import versions as common_versions
        out["app"] = common_versions.all_versions().get("app", "")
    except Exception:                                   # noqa: BLE001 - 版が読めなくても記録は残す
        pass
    return out


# ==================================================================
# 書く
# ==================================================================
def _month_file(folder: Path, when: datetime.datetime) -> Path:
    return folder / f"{FILE_PREFIX}_{when:%Y-%m}.jsonl"


def folders() -> dict[str, Optional[Path]]:
    """書く場所(この PC・設定の場所)。設定の場所は設定が無ければ None。"""
    shared: Optional[Path] = None
    text = record_dir_setting()
    if text:
        try:
            shared = config.resolve_dir(text) / pc_folder_name()
        except ValueError as exc:
            log.warning("記録の保存先フォルダーが読めません: %s", exc)
    return {"local": Path(config.RECORDS_DIR), "shared": shared}


def append(kind: str, payload: Optional[dict[str, Any]] = None, *,
           when: Optional[datetime.datetime] = None) -> dict[str, str]:
    """1 件を月ごとのファイルへ足す。戻り値は書けなかった理由(書けたら空)。

    **例外にしない。** 記録が書けなくても出力・印刷は止めない。
    """
    when = when or datetime.datetime.now()
    try:
        entry = {"at": when.strftime("%Y-%m-%d %H:%M:%S"), "kind": kind,
                 **(payload or {}), "version": versions(), "who": who()}
        line = json.dumps(entry, ensure_ascii=False) + "\n"
        where = folders()
    except Exception as exc:                            # noqa: BLE001 - 止めない
        log.warning("作業の記録を組めません(続けます): %s", exc)
        return {"local": str(exc), "shared": ""}
    out: dict[str, str] = {}
    with _lock:
        for key, folder in where.items():
            if folder is None:
                out[key] = ""
                _problems[key] = ("記録の保存先フォルダーが読めません"
                                  if key == "shared" and record_dir_setting() else "")
                continue
            try:
                folder.mkdir(parents=True, exist_ok=True)
                with open(_month_file(folder, when), "a", encoding="utf-8", newline="\n") as f:
                    f.write(line)
                out[key] = ""
            except OSError as exc:
                out[key] = f"{folder} に書けません({exc.strerror or exc})"
                log.warning("作業の記録の控えを書けません: %s", out[key])
            _problems[key] = out[key]
    log.info("[記録] %s %s", kind, json.dumps(payload or {}, ensure_ascii=False))
    return out


def status() -> dict[str, str]:
    """記録の画面に出す: どこに控えを書いているか・書けていない理由。"""
    where = folders()
    with _lock:
        return {"local": str(where["local"]),
                "shared": str(where["shared"]) if where["shared"] else "",
                "shared_setting": record_dir_setting(),
                "local_problem": _problems.get("local", ""),
                "shared_problem": _problems.get("shared", "")}


# ==================================================================
# 読む(記録の画面)
# ==================================================================
def recent(limit: int = 300, *, today: Optional[datetime.date] = None) -> list[dict[str, Any]]:
    """この PC の記録の新しいものから ``limit`` 件。今月と先月のファイルを読む。"""
    today = today or datetime.date.today()
    first = today.replace(day=1)
    months = [first, (first - datetime.timedelta(days=1)).replace(day=1)]
    folder = Path(config.RECORDS_DIR)
    out: list[dict[str, Any]] = []
    for month in months:
        path = folder / f"{FILE_PREFIX}_{month:%Y-%m}.jsonl"
        try:
            lines = path.read_text(encoding="utf-8").splitlines()
        except FileNotFoundError:
            continue
        except OSError as exc:
            log.warning("作業の記録を読めません: %s (%s)", path, exc)
            continue
        for text in reversed(lines):
            if not text.strip():
                continue
            try:
                item = json.loads(text)
            except ValueError:
                continue                                # 書きかけの行は飛ばす
            if isinstance(item, dict):
                out.append(item)
            if len(out) >= limit:
                return out
    return out


def reset_for_tests() -> None:
    """試験用。"""
    with _lock:
        _problems.update(local="", shared="")
