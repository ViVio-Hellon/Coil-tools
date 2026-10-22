# -*- coding: utf-8 -*-
"""計算・印刷の記録の「消えない控え」と、記録に添える素性(統合 1.2.39・現場の指示)

画面の「計算・印刷の記録」は、この PC の状態DB(`work_log`)に 5,000 件まで持ち、古いものから消える。
PC の入れ替えや状態DB の破損でも消える。そこで 1 件ごとに、**月ごとのファイルへ 1 行ずつ足す**
(JSON Lines。足すだけで、消さない・書き換えない):

    この PC     <ローカル領域>/records/ペナラベル記録_YYYY-MM.jsonl      (いつも)
    設定の場所  <記録の保存先フォルダー>/<PC名>/ペナラベル記録_YYYY-MM.jsonl
                (設定「計算・印刷の記録の保存先フォルダー」= record_dir があるときだけ。PC ごとに
                 ファイルを分けるので、複数のラインが同じファイルを取り合わない)

記録に添える素性(どの記録からも言えるように):

- **ソフトの版**: 統合ツールの版・ペナラベルの版・中身の指紋(動いているファイルの中身から作った値)。
  「切り替えの前か後か」「どの版で刷ったか」を記録から言うため
- **PC 名**(Windows の COMPUTERNAME)と**ライン名**(設定「ライン名」。この PC)、Windows の利用者名

書けなくても計算・印刷は止めない(理由はログと記録の画面に出す)。
"""
from __future__ import annotations

import datetime
import getpass
import json
import logging
import os
import re
import socket
import threading
from typing import Any, Dict, Optional

log = logging.getLogger(__name__)

_lock = threading.Lock()
#: 最後に書けなかった理由(記録の画面に出す)。{"local": .., "shared": ..}
_problems: Dict[str, str] = {"local": "", "shared": ""}

FILE_PREFIX = "ペナラベル記録"


def pc_name() -> str:
    name = os.environ.get("COMPUTERNAME") or socket.gethostname() or "PC"
    return re.sub(r'[\\/:*?"<>|]', "_", name).strip() or "PC"


def user_name() -> str:
    try:
        return os.environ.get("USERNAME") or getpass.getuser() or ""
    except Exception:                                   # noqa: BLE001
        return ""


#: テストモードのときの印(記録・編集履歴に残す)
TEST_MODE_MARK = "テストモード"


def who(cfg) -> Dict[str, str]:
    """どの PC(ライン)・誰の記録か。テストモードなら ``mode`` に「テストモード」(統合 1.2.40)。"""
    out = {"pc": pc_name(), "line": str(getattr(cfg, "line_name", "") or "").strip(),
           "user": user_name()}
    if getattr(cfg, "test_mode", False):
        out["mode"] = TEST_MODE_MARK
    return out


def versions() -> Dict[str, str]:
    """ソフトの版(統合ツール・ペナラベル)と中身の指紋。"""
    from ..config import APP_VERSION, code_stamp
    out = {"pena": APP_VERSION}
    try:
        from common import versions as common_versions
        out["app"] = common_versions.all_versions().get("app", "")
    except Exception:                                   # noqa: BLE001 - 単体のペナラベル
        pass
    try:
        out["stamp"] = code_stamp()
    except Exception:                                   # noqa: BLE001
        pass
    return out


def _month_file(folder: str, when: datetime.datetime) -> str:
    return os.path.join(folder, "%s_%s.jsonl" % (FILE_PREFIX, when.strftime("%Y-%m")))


def folders(cfg) -> Dict[str, str]:
    """書く場所(この PC・設定の場所)。設定の場所は設定が無ければ空。"""
    local = os.path.join(str(cfg.local_dir), "records")
    root = os.path.expandvars(str(getattr(cfg, "record_dir", "") or "").strip())
    shared = os.path.join(root, pc_name()) if root else ""
    return {"local": local, "shared": shared}


def append(cfg, entry: Dict[str, Any], when: Optional[datetime.datetime] = None) -> Dict[str, str]:
    """1 件を月ごとのファイルへ足す。戻り値は書けなかった理由(書けたら空)。"""
    when = when or datetime.datetime.now()
    line = json.dumps(entry, ensure_ascii=False, sort_keys=False) + "\n"
    out = {}
    with _lock:
        for key, folder in folders(cfg).items():
            if not folder:
                out[key] = ""
                _problems[key] = ""
                continue
            try:
                os.makedirs(folder, exist_ok=True)
                with open(_month_file(folder, when), "a", encoding="utf-8", newline="\n") as f:
                    f.write(line)
                out[key] = ""
            except OSError as exc:
                out[key] = "%s に書けません(%s)" % (folder, exc.strerror or exc)
                log.warning("計算・印刷の記録の控えを書けません: %s", out[key])
            _problems[key] = out[key]
    return out


def status(cfg) -> Dict[str, str]:
    """記録の画面に出す: どこに控えを書いているか・書けていない理由。"""
    where = folders(cfg)
    with _lock:
        return {"local": where["local"], "shared": where["shared"],
                "localProblem": _problems.get("local", ""),
                "sharedProblem": _problems.get("shared", "")}
