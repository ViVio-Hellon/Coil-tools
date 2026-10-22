"""資材マスタの編集履歴(VER 0.13.30。ペナラベルと同じ表・同じ列)

このアプリが共有へ書くのは 2 つだけ:

- 紙面の右上の文字(梱包資材マスタの表「梱包明細打ち出し」。`qa_mark`)
- 管理者パスワード(控えのJSON。`admin_password`)

変えたときに、**誰が(PC・ライン・ログインID)・いつ・何を何から何に**変えたかを、
梱包資材マスタの中の表「資材マスタ編集履歴」に 1 行足す(ペナラベルが書く表と同じ。
全ラインで 1 か所)。表・列が無ければ作る。

**管理者パスワードの値は書かない。** 変えた・既定に戻した、だけを残す
(前・後は「(伏せる)」)。

書けなくても変更そのものは止めない(変更はもう済んでいる。理由はログへ)。
梱包資材マスタが無いときは書く場所が無いので、ログにだけ残す。
"""
from __future__ import annotations

import sqlite3
import time
from pathlib import Path
from typing import Any, Iterable, Optional

from . import shared_settings, source_db, work_record
from .logging_utils import get_logger

log = get_logger("edit_history")

TABLE = "資材マスタ編集履歴"
COLUMNS = ("日時", "PC", "ライン", "利用者", "表", "行", "行の名前", "操作", "列", "前", "後",
           "モード")
#: 値を残さない欄(管理者パスワード)
HIDDEN = "(伏せる)"


def _q(name: str) -> str:
    return source_db.quote_identifier(name)


def _ensure(conn) -> None:
    """表と列をそろえる(無ければ作る・足す)。"""
    conn.execute("CREATE TABLE IF NOT EXISTS %s (%s)" % (
        _q(TABLE), ", ".join("%s TEXT" % _q(c) for c in COLUMNS)))
    have = {r["name"] for r in conn.query(f"PRAGMA table_info({_q(TABLE)})")}
    for c in COLUMNS:
        if c not in have:
            conn.execute(f"ALTER TABLE {_q(TABLE)} ADD COLUMN {_q(c)} TEXT")


def write(table: str, row_id: Any, row_name: str, action: str,
          changes: Iterable[tuple[str, Any, Any]], *,
          path: Optional[Path] = None, keep_same: bool = False) -> str:
    """編集履歴を足す。戻り値は書けなかった理由(書けたら空)。**例外にしない。**

    ``changes`` は (列, 前, 後) の並び。値が変わっていない列は書かない
    (``keep_same`` なら書く。伏せた値は前後とも同じ字になるため)。
    """
    try:
        who = work_record.who()
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        rows = []
        for col, before, after in changes:
            b = "" if before is None else str(before)
            a = "" if after is None else str(after)
            if b == a and not keep_same:
                continue
            rows.append((stamp, who.get("pc", ""), who.get("line", ""), who.get("user", ""),
                         table, str(row_id), row_name, action, col, b, a, who.get("mode", "")))
        if not rows:
            return ""
        target = path or shared_settings.master_path()
        if target is None:
            problem = "梱包資材マスタが無いので編集履歴を書けません"
            log.warning("%s: %s %s %s", problem, table, action, row_name)
            return problem
        with source_db.connect(target) as conn:
            _ensure(conn)
            for row in rows:
                conn.insert(TABLE, dict(zip(COLUMNS, row)))
    except Exception as exc:                            # noqa: BLE001 - 変更は済んでいる。止めない
        log.warning("編集履歴を書けません(変更は済んでいます): %s", exc)
        return f"編集履歴を書けません: {exc}"
    log.info("編集履歴を書きました: %s %s %s(%s行)", table, row_name, action, len(rows))
    return ""


def write_secret(table: str, row_id: Any, row_name: str, action: str, column: str) -> str:
    """値を残さない変更(管理者パスワード)。前・後は「(伏せる)」。"""
    return write(table, row_id, row_name, action, [(column, HIDDEN, HIDDEN)], keep_same=True)


def read(limit: int = 300) -> dict[str, Any]:
    """編集履歴(新しい順)。表が無ければ空。"""
    try:
        path = shared_settings.master_path()
    except Exception as exc:                            # noqa: BLE001
        return {"rows": [], "path": "", "note": f"梱包資材マスタを探せません: {exc}"}
    if path is None:
        return {"rows": [], "path": str(shared_settings.shared_dir()),
                "note": "梱包資材マスタがありません"}
    try:
        conn = source_db.open_read_only(path)
        try:
            conn.row_factory = sqlite3.Row
            cur = conn.execute(f"SELECT * FROM {_q(TABLE)} ORDER BY rowid DESC LIMIT ?",
                               (int(limit),))
            rows = [{k: ("" if r[k] is None else str(r[k])) for k in r.keys()}
                    for r in cur.fetchall()]
        finally:
            conn.close()
    except (sqlite3.Error, source_db.SourceError, OSError) as exc:
        if "no such table" in str(exc):
            return {"rows": [], "path": str(path), "note": ""}
        return {"rows": [], "path": str(path), "note": f"編集履歴を読めません: {exc}"}
    return {"rows": rows, "path": str(path), "note": ""}
