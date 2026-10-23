"""パレット種類_外径条件 ── 包装仕様Noごとに、外径でパレット種類を決める表(統合 1.2.50・現場の指示)

【なぜ表にしたのか】
1C1103(ｲｽﾞﾐﾒﾀﾙ)は、パレット種類をマスタ(包装仕様)の値のあとで外径によって上書きして
いた(VBA `PalletType` の末尾。`special_rules.pallet_type_override`)。現場の見立ては
「本来マスタ優先であるべきだが、マスタに外径の条件が無かったのでコードで対応した」。
そこで梱包資材マスタに外径の条件を置く表を足し、**表があれば表(マスタ)で決める**。

    表がまだ無い(共有に作っていない) … いままでどおりプログラムの規則で決める(動きを変えない)
    表がある                          … どの包装仕様No も表で決める。プログラムの規則は使わない
        その包装仕様No の行が無い     … マスタ(包装仕様)のパレット種類のまま
        行はあるが外径がどれにも入らない … マスタ(包装仕様)のパレット種類のまま

【決め方】
その包装仕様No の行を 順番 の小さい順(同じ順番なら取り込み元の行の順)に見て、

    (外径_以上 が空 または 外径 >= 外径_以上) かつ (外径_以下 が空 または 外径 <= 外径_以下)

に**最初に当たった行**のパレット種類にする(先勝ち。VBA の `ElseIf` と同じ)。
外径が空のときは 0 として当てる ── プログラムの規則も `vba.val("")` = 0 で見ていて
(0 < 800 なので「全面」)、作った直後の表でも同じ結果になる。

【表が有るか無いか】
手元の作業DBにこの表が**あるか**で見る。手元の表は `schema.sql` に置かず、取り込みのときに
作る(共有に表があれば作り、無ければ消す。`data_sync._import_optional`)。

【表の作り方】
Access から直に変換はしない(現場の運用の決め)。マスタ管理の「表を作る」で、共有の
梱包資材マスタにこの表を作り、いまのプログラムの規則と同じ結果になる行を入れる
(`master_admin.create_table`)。
"""
from __future__ import annotations

import sqlite3
from typing import Any, Optional

from . import config, special_rules
from .logging_utils import get_logger

log = get_logger("outer_condition")

TABLE = config.TBL_OUTER_CONDITION

# 手元の表。取り込みのときに作る(`schema.sql` には置かない。上の【表が有るか無いか】)。
# 外径_以上・外径_以下 は空(NULL)が「決めていない」なので、NOT NULL にしない
LOCAL_DDL = f"""
CREATE TABLE IF NOT EXISTS "{TABLE}" (
    id            INTEGER PRIMARY KEY,
    包装仕様No    TEXT NOT NULL DEFAULT '',
    順番          INTEGER NOT NULL DEFAULT 0,
    外径_以上     REAL,
    外径_以下     REAL,
    パレット種類  TEXT NOT NULL DEFAULT '',
    備考          TEXT NOT NULL DEFAULT ''
)
"""

# 共有の梱包資材マスタに作る表(マスタ管理の「表を作る」)
SOURCE_DDL = (f'CREATE TABLE "{TABLE}" ('
              '"包装仕様No" TEXT, "順番" INTEGER, "外径_以上" REAL, "外径_以下" REAL,'
              ' "パレット種類" TEXT, "備考" TEXT)')

# 表を作るときに入れる行。**いまのプログラムの規則(1C1103)と同じ結果になる**ようにしてある。
#
# VBA の並びは  外径 < 800 → 全面 / 780〜890 → 強度UP / 890〜990 → 強度UP  で、境目が
# 重なっている(780〜799 は最初の条件にも当たり、先勝ちで「全面」)。実際の動きは
#     外径 < 800       → 全面
#     800 〜 990       → 強度UP(両端を含む)
#     990 を超える     → マスタの値のまま
# 外径は小数2桁にそろえて入るので、「800 未満」は「799.99 以下」と書ける
SEED_ROWS: tuple[dict[str, Any], ...] = (
    {"包装仕様No": special_rules.IZUMI_SPEC_NO, "順番": 1, "外径_以上": None, "外径_以下": 799.99,
     "パレット種類": "全面", "備考": "VBA の 外径<800"},
    {"包装仕様No": special_rules.IZUMI_SPEC_NO, "順番": 2, "外径_以上": 800.0, "外径_以下": 990.0,
     "パレット種類": "強度UP", "備考": "VBA の 780〜990(780〜799 は先の条件が勝つ)"},
)


def present(conn: Optional[sqlite3.Connection]) -> bool:
    """手元の作業DBにこの表があるか(取り込まれているか)。あれば**マスタ優先**。"""
    if conn is None:
        return False
    try:
        row = conn.execute(
            "SELECT 1 FROM sqlite_master WHERE type = 'table' AND name = ?", (TABLE,)).fetchone()
    except sqlite3.Error:
        return False
    return row is not None


def ensure_local(conn: sqlite3.Connection) -> None:
    """手元の表を作る(取り込みの直前に呼ぶ)。"""
    with conn:
        conn.execute(LOCAL_DDL)


def drop_local(conn: sqlite3.Connection) -> None:
    """手元の表を消す(共有に表が無くなったとき)。消すと、プログラムの規則に戻る。"""
    with conn:
        conn.execute(f'DROP TABLE IF EXISTS "{TABLE}"')


def local_columns_info() -> list[sqlite3.Row]:
    """手元の表の列(型・NULL を許すか)。手元にまだ表が無いときのマスタ管理の欄に使う。"""
    mem = sqlite3.connect(":memory:")
    mem.row_factory = sqlite3.Row
    try:
        mem.execute(LOCAL_DDL)
        return mem.execute(f'PRAGMA table_info("{TABLE}")').fetchall()
    finally:
        mem.close()


def matches(low: Optional[float], high: Optional[float], outer: float) -> bool:
    """外径が 以上〜以下 に入るか。空(None)の端は「決めていない」(どこまでも入る)。"""
    return (low is None or outer >= float(low)) and (high is None or outer <= float(high))


def pallet_type(conn: sqlite3.Connection, spec_no: str, outer: float, current: str) -> str:
    """表で決めたパレット種類。当たる行が無ければ `current`(マスタの値)のまま。

    表があるときだけ呼ぶ(`present`)。外径が空(0)でも、そのまま 0 として当てる。
    """
    try:
        rows = conn.execute(
            f'SELECT 外径_以上, 外径_以下, パレット種類 FROM "{TABLE}"'
            " WHERE 包装仕様No = ? ORDER BY 順番, id", (str(spec_no or "").strip(),)).fetchall()
    except sqlite3.Error as exc:                   # pragma: no cover - 表はあると確かめてから呼ぶ
        log.warning("%s を引けません: %s", TABLE, exc)
        return current
    for low, high, kind in (tuple(r) for r in rows):
        if matches(low, high, outer) and str(kind or "").strip():
            return str(kind).strip()
    return current


def decide(conn: sqlite3.Connection, spec_no: str, outer: float, current: str) -> str:
    """パレット種類の上書き。**表があれば表(マスタ優先)、無ければプログラムの規則。**"""
    if present(conn):
        return pallet_type(conn, spec_no, outer, current)
    return special_rules.pallet_type_override(spec_no, outer, current)
