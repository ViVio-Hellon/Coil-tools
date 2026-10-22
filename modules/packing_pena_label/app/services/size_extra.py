# -*- coding: utf-8 -*-
"""あとから足したサイズ(共有マスタの ``ペナラベル追加サイズ`` 表。統合 1.2.16)

これまでサイズを増やすには ``data/size_master.json`` を手で直すしかなかった(現場の指摘:
型番は足せてもサイズを足す仕組みが無い)。新しいサイズが来たら、設定の「サイズ」から足す。

**書き先は共有の資材マスタ(SQLite)。** 全ラインで同じサイズを使うため。
JSON(同梱・設定のサイズ構成マスタ)は土台のまま、この表のサイズを後ろに足す。
表が無い・読めない端末では、何も足さずに今までどおり動く。

    管理番号 … 丈1 の台紙番号(obIdx。奇数)。丈2 はその次の番号
    サイズ名 … 「1.2mm×45.0mm」(表記は normalize_size_name で揃える)
    巾       … mm(サイズ名の後ろの数。積み高さ・HB の計算に使う)
    型番     … ラベルに刷る型番
    記号     … 型番マスタの記号(画面に出すだけ。計算には使わない)
    備考     … 指定サイズの画面に赤字で出す注記
    登録日時 … 足した日時
    削除日時 … マスタ管理で消した日時(統合 1.2.38)。入っている行は読まない

サイズ名と巾は**あとから直さない**(台紙・状態がこの番号とサイズで結び付く)。
間違えたときは、マスタ管理で消して(または画面に出さないようにして)、正しいものを足し直す。
型番・記号・備考はマスタ管理から直せる。

**消すのは論理削除**(行は残し、削除日時を入れる。統合 1.2.38・現場の指摘)。管理番号を新しいサイズで
使い回すと、各 PC の手元に残った台紙の中身(前のサイズの検番・重量・型番)が新しいサイズに出てしまう。
消したサイズと同じ寸法は、もう一度足せる(番号は新しくなる)。
"""

from __future__ import annotations

import datetime
import logging
import os
import re
import sqlite3
import time
import unicodedata
from decimal import Decimal, InvalidOperation
from typing import Dict, List, Optional, Tuple

from . import master_admin as MA
from .size_master import FIRST_EXTRA_OB, normalize_size_name

log = logging.getLogger(__name__)

TABLE_NAME = "ペナラベル追加サイズ"
COLUMNS = ("管理番号", "サイズ名", "巾", "型番", "記号", "備考", "登録日時")
#: 消した日時の列(論理削除。master_admin の Managed.soft_delete と同じ名前)
DELETED_COLUMN = "削除日時"

#: 足したサイズの台紙番号はここから(同梱の 1..14 と重ねない)
FIRST_OB = FIRST_EXTRA_OB

#: 更新の確認を省く時間(秒)。共有フォルダーへの stat を減らす
DEFAULT_REFRESH_SEC = 5.0

_KATABAN_RE = re.compile(r"[A-Z0-9][A-Z0-9\-]{2,29}")


class SizeExtra:
    """``ペナラベル追加サイズ`` を読む(更新日時が変わったときだけ読み直す)。"""

    def __init__(self, refresh_sec: float = DEFAULT_REFRESH_SEC):
        self.refresh_sec = float(refresh_sec)
        self.last_error = ""
        self._rows: List[dict] = []
        self._stamp: Optional[Tuple[str, float, int]] = None
        self._at = 0.0

    def invalidate(self) -> None:
        self._stamp = None
        self._at = 0.0

    def rows(self) -> List[dict]:
        return [dict(r) for r in self._rows]

    def load(self, db_path: str, force: bool = False) -> List[dict]:
        path = str(db_path or "")
        if not path or not os.path.exists(path):
            self._rows, self._stamp = [], None
            return []
        now = time.monotonic()
        if not force and self._stamp is not None and (now - self._at) < self.refresh_sec:
            return self.rows()
        try:
            st = os.stat(path)
            stamp = (os.path.abspath(path), st.st_mtime, st.st_size)
        except OSError:
            stamp = None
        if not force and stamp is not None and stamp == self._stamp:
            self._at = now
            return self.rows()
        self._rows = read_rows(path)
        self._stamp, self._at = stamp, now
        return self.rows()


def read_rows(path: str) -> List[dict]:
    """表の行。表が無い・読めないときは空(今までどおり動く)。"""
    out: List[dict] = []
    try:
        conn = MA._connect(path)
        try:
            cur = conn.execute('SELECT * FROM "%s" ORDER BY "管理番号"' % TABLE_NAME)
            for r in cur.fetchall():
                if str(dict(r).get(DELETED_COLUMN) or "").strip():
                    continue                    # 消したサイズ(統合 1.2.38)
                row = _row(dict(r))
                if row:
                    out.append(row)
        finally:
            conn.close()
    except sqlite3.Error as exc:
        if "no such table" not in str(exc):
            log.info("%s を読めません(足したサイズは出ません): %s", TABLE_NAME, exc)
    return out


def _row(r: dict) -> Optional[dict]:
    """1 行を使える形に。壊れた行は飛ばす(業務を止めない)。"""
    try:
        ob = int(r.get("管理番号"))
        width = float(r.get("巾"))
    except (TypeError, ValueError):
        return None
    name = normalize_size_name(str(r.get("サイズ名") or "").strip())
    if ob < FIRST_OB or ob % 2 == 0 or not name or width <= 0:
        return None
    return {"obIdx": ob, "baseName": name, "widthMm": width,
            "kataban": str(r.get("型番") or "").strip(),
            "symbol": str(r.get("記号") or "").strip(),
            "note": str(r.get("備考") or "").strip()}


# ---------------------------------------------------------------- 足す
def _number(raw, label: str, low: Decimal, high: Decimal, places: int) -> Decimal:
    text = unicodedata.normalize("NFKC", str(raw or "")).strip().replace("ｍｍ", "").replace("mm", "")
    if not re.fullmatch(r"[0-9]+(\.[0-9]+)?", text):
        raise MA.Refused(MA.REFUSE_BAD_VALUE, "%sは数で入れてください(例: %s)" % (label, "1.2" if places == 3 else "45.5"))
    num = Decimal(text)
    if not (low < num < high):
        raise MA.Refused(MA.REFUSE_BAD_VALUE, "%sが範囲の外です(%s より大きく %s より小さく)" % (label, low, high))
    if -num.as_tuple().exponent > places:
        raise MA.Refused(MA.REFUSE_BAD_VALUE, "%sは小数 %d 桁までです" % (label, places))
    return num


def _plain(num: Decimal) -> str:
    """53.50 → 53.5、2 → 2(このあと normalize_size_name が 2.0 にする)。"""
    text = format(num.normalize(), "f")
    return text


def size_name(thickness, width) -> Tuple[str, float]:
    t = _number(thickness, "厚み", Decimal("0"), Decimal("10"), 3)
    w = _number(width, "巾", Decimal("0"), Decimal("1000"), 2)
    return normalize_size_name("%smm×%smm" % (_plain(t), _plain(w))), float(w)


def _dims(name: str) -> Tuple[float, float]:
    nums = re.findall(r"(\d+(?:\.\d+)?)mm", name)
    return tuple(float(n) for n in nums[:2]) if len(nums) >= 2 else (0.0, 0.0)


def add(repo, master, thickness, width, kataban, symbol="", note="") -> dict:
    """サイズを 1 つ足す。共有の資材マスタ(SQLite)に書く。

    断るとき: 書き先が無い・読み取り専用 / 数の形が違う / 同じ寸法のサイズがもうある /
    型番が空・形が違う。
    """
    why = MA.editable_why_not(repo)
    if why:
        raise MA.Refused(MA.REFUSE_NO_SOURCE, why)
    path = MA.source_path(repo)
    name, width_mm = size_name(thickness, width)
    kata = unicodedata.normalize("NFKC", str(kataban or "")).strip().upper()
    if not _KATABAN_RE.fullmatch(kata):
        raise MA.Refused(MA.REFUSE_BAD_VALUE,
                         "型番を入れてください(英数字とハイフン。例: BJB7610400QR)")
    sym = unicodedata.normalize("NFKC", str(symbol or "")).strip()
    if len(sym) > 10:
        raise MA.Refused(MA.REFUSE_BAD_VALUE, "記号は 10 文字までです")
    memo = str(note or "").strip()
    if len(memo) > 60:
        raise MA.Refused(MA.REFUSE_BAD_VALUE, "備考は 60 文字までです")

    have = {_dims(c.base_name): c.base_name for c in master.all()}
    if _dims(name) in have:
        raise MA.Refused(MA.REFUSE_BAD_VALUE,
                         "同じ寸法のサイズがもうあります: %s\n"
                         "画面に出ていないだけなら、表示のチェックを入れてください。" % have[_dims(name)])

    conn = MA._connect(path, write=True)
    try:
        conn.execute(
            'CREATE TABLE IF NOT EXISTS "%s" ("管理番号" INTEGER, "サイズ名" TEXT, "巾" REAL, '
            '"型番" TEXT, "記号" TEXT, "備考" TEXT, "登録日時" TEXT, "%s" TEXT)'
            % (TABLE_NAME, DELETED_COLUMN))
        # 番号は消したサイズの分も数える(使い回さない。統合 1.2.38)
        used = [int(r[0]) for r in conn.execute('SELECT "管理番号" FROM "%s"' % TABLE_NAME)
                if str(r[0]).lstrip("-").isdigit()]
        has_deleted = DELETED_COLUMN in MA.columns(conn, TABLE_NAME)
        alive = ('SELECT "サイズ名" FROM "%s" WHERE COALESCE("%s", \'\') = \'\''
                 % (TABLE_NAME, DELETED_COLUMN)) if has_deleted else \
            'SELECT "サイズ名" FROM "%s"' % TABLE_NAME
        for r in conn.execute(alive):
            other = normalize_size_name(str(r[0] or ""))
            if _dims(other) == _dims(name):
                raise MA.Refused(MA.REFUSE_BAD_VALUE, "同じ寸法のサイズがもうあります: %s" % other)
        top = max([FIRST_OB - 2] + list(master.ob_indexes()) + used)
        ob = top + 1 if top % 2 == 0 else top + 2
        backup = MA._backup(path)
        stamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute('INSERT INTO "%s" (%s) VALUES (?,?,?,?,?,?,?)'
                     % (TABLE_NAME, ", ".join('"%s"' % c for c in COLUMNS)),
                     (ob, name, width_mm, kata, sym, memo, stamp))
        conn.commit()
    except MA.Refused:
        conn.rollback()
        raise
    except sqlite3.Error as exc:
        conn.rollback()
        raise MA.Refused(MA.REFUSE_WRITE_FAILED, "書き込めませんでした: %s" % exc)
    finally:
        conn.close()
    log.info("サイズを足しました: %s(台紙 %d/%d・型番 %s)", name, ob, ob + 1, kata)
    return {"ok": True, "baseName": name, "obIdx": [ob, ob + 1], "kataban": kata,
            "backup": backup}
