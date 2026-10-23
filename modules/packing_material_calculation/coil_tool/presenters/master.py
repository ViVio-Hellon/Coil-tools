"""マスタ管理の画面に出すもの

判断は `master_admin` が済ませてあり、ここは**並べ方と言葉**だけを持つ。

【共有フォルダに触るのは、画面を開いたときだけ】
最初の描画では**共有に触らない分**(直せるかどうか・どこへ書くか)だけ
を出し、中身は開いてから読む(`browse`)。共有が遠い日に、画面を開く
だけで待たされないようにするため。
"""
from __future__ import annotations

import sqlite3
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any, Optional

from .. import admin_session, config, data_sync, master_admin, outer_condition


@dataclass
class MasterViewModel:
    # --- 共有に触らずに分かること ---
    can_edit: bool = False
    edit_why: str = ""
    source_dir: str = ""
    # 認証の状態。**パスワードそのものは入らない**
    auth: Optional[admin_session.State] = None

    # --- 中身(開いてから読む) ---
    loaded: bool = False
    source: str = ""
    tables: list[master_admin.TableInfo] = field(default_factory=list)
    table: str = ""
    query: str = ""
    columns: list[master_admin.Column] = field(default_factory=list)
    page: Optional[master_admin.Page] = None
    message: str = ""
    # 共有のマスタと手元の写しの食い違い(統合 1.2.49)。この画面は共有を直に読むので、
    # 手元(計算が読むほう)が追いついていないと「画面は新しいのに計算は古い」になる
    sync_note: str = ""


def frame(conn: Optional[sqlite3.Connection]) -> MasterViewModel:
    """最初の描画ぶん。**共有フォルダには触らない。**

    `touch=False` ── 画面を出すだけでは認証の時間切れを延ばさない。
    延ばすと、タブを開きっぱなしにするだけで永久に開いたままになる。
    """
    allowed, why = master_admin.can_edit(conn, touch=False)
    return MasterViewModel(can_edit=allowed, edit_why=why,
                           source_dir=str(config.master_db_dir()),
                           auth=admin_session.state())


def browse(conn: sqlite3.Connection, *, table: str = "", query: str = "",
           sort: str = "", sort_dir: str = "asc",
           path: Optional[Path] = None, message: str = "") -> MasterViewModel:
    """画面を開いたとき / 直したあとに返す、まるごとの状態。

    直したあとも同じものを返す。「保存しました」だけを返すと、
    **本当に入ったかどうか**を画面が別に確かめに行くことになる。
    """
    view = frame(conn)
    view.message = message
    view.sync_note = _sync_note(conn)
    found = path or data_sync.find_master_db()
    view.loaded = True
    view.source = str(found) if found else ""
    view.tables = master_admin.tables(found)

    view.table = _pick(view.tables, table)
    view.query = query or ""
    if not view.table:
        view.page = master_admin.page(None, "", query=view.query)
        return view

    # 表が決まってから引き直す。`frame()` は表を知らない時点の判定。
    # ここも見るだけなので時間切れは延ばさない
    view.can_edit, view.edit_why = master_admin.can_edit(
        conn, view.table, touch=False)

    # 札の出し分けは手元(計算が読むほう)に外径条件の表があるかで見る(統合 1.2.50)。
    # `_sync_note` が先に手元を追いつかせているので、共有で作られた表もここで分かる
    view.page = master_admin.page(found, view.table, query=view.query,
                                  sort=sort, sort_dir=sort_dir,
                                  outer_table=outer_condition.present(conn))
    # 打ち込める欄は、**取り込み元に本当にある列**だけにする。
    # 上流がまだ足していない列を出すと、保存の瞬間に断られる
    view.columns = master_admin.columns(conn, view.table, view.page.columns)
    if (not view.columns and not view.page.missing and not view.page.error
            and view.page.editable):
        # 表はある(missing=False)のに1つも打ち込めない ── たいてい
        # 列名が想定と違う。空の編集窓を出すだけでは分からないので、
        # ここで理由を足す(既存の `why` は「見るだけの表」用なので、
        # 直せる表のこの状況では元々空)
        mismatch = master_admin.column_mismatch_why(view.table, view.page.columns)
        if mismatch:
            view.page.why = mismatch
    return view


def _sync_note(conn: sqlite3.Connection) -> str:
    """共有のマスタに、手元へまだ取り込んでいない変更があるか(統合 1.2.49・現場の指摘)。

    変わっていれば、その場で手元へ取り込み直す(小さい表なので)。取り込み直せなければ一言出す。
    """
    try:
        fresh = data_sync.freshness(conn)
    except Exception:                          # noqa: BLE001 - 見られなくても画面は出す
        return ""
    if fresh.importing:
        return "取り込み中です。終わるまで、計算は前の写しで動きます。"
    if fresh.check.master:
        return ("共有のマスタに、手元へまだ取り込んでいない変更があります(計算は手元の写しで動きます)。"
                "上の「いま取り込む」を押してください。")
    done = [t for t in fresh.refreshed if t in data_sync.master_tables()]
    if done:
        return ("共有のマスタが変わっていたので、手元へ取り込み直しました(" + "・".join(done)
                + ")。資材計算をやり直すと新しい値になります。")
    return ""


def _pick(tables: list[master_admin.TableInfo], wanted: str) -> str:
    """出す表を決める。

    指定が無い / 取り込み元に無いときは**先頭**にする ── 空の画面を出して
    「自分で選べ」とするより、いちばん触る表を開いておくほうが早い。
    """
    names = [t.table for t in tables]
    if wanted and wanted in names:
        return wanted
    return names[0] if names else ""


def to_dict(view: MasterViewModel) -> dict[str, Any]:
    return {
        "can_edit": view.can_edit,
        "edit_why": view.edit_why,
        "source_dir": view.source_dir,
        "auth": {"open": view.auth.open, "remains": view.auth.remains,
                 "custom": view.auth.custom} if view.auth else None,
        "loaded": view.loaded,
        "source": view.source,
        "tables": [t.to_dict() for t in view.tables],
        "table": view.table,
        "query": view.query,
        # 型の言い方はサーバが持つ。画面側で「整数」と書き分けない
        "columns": [{**c.to_dict(),
                     "kind_label": master_admin.KIND_LABEL.get(c.kind, c.kind)}
                    for c in view.columns],
        "page": view.page.to_dict() if view.page else None,
        "row_key": master_admin.ROW_KEY,
        "message": view.message,
        "sync_note": view.sync_note,
    }
