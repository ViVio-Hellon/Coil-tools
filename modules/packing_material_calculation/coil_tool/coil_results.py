"""コイル実績 ── 実績の形式のファイルからコイル外径を出す(統合 1.2.29・現場の指示)

梱包課共有の仕掛のファイル名が **LS4LOT 以外**(実績の形式。例: 実績_LS4.sqlite3)のときだけ読む。
名前が LS4LOT のままなら読まない(内径・肉厚が無いので、いままでどおり受注の外径 MAX を使う)。

    外径 = BOX実績_ｺｲﾙ内径 + 2 × BOX実績_ｺｲﾙ肉厚
    検入数(本) = BOX実績_枚本数(無ければ BOX設計_横割数 × BOX設計_縦割数。統合 1.2.30)

- BOX実績_作業開始_日付 が**いまの年月**の行だけを使い、同じロットは作業開始がいちばん新しい行
  (形式の決まりは梱包明細と同じ部品 `common/ls4_source.py`。仕様は `docs/仕掛ファイルの仕様.md`)
- 第1パス(梱包課共有のフォルダ・名前) → 第2パス(仕掛台帳の第2パスのフォルダ・LS4LOT の代わりの名前)。
  第1パスに無いロットは第2パスから足す。LS4LOT の生成時刻と合わないときの受け皿。
  **第1パスにあっても、外径(内径・肉厚)・検入数が空なら、第2パスの同じロットの行で空の項目だけ埋める**
  (統合 1.2.45・現場の指摘: 第1パスが読めても検入数が埋まらないことがあった。以前は第1パスの行だけを使った)
- 外径が出せない(内径・肉厚が空・0)ロットは、いままでの参照物(受注の外径 MAX)を使う
- 共有のファイルは手元へ写してから読む(`source_db` → `common/source_copy.py`)
"""
from __future__ import annotations

import sqlite3
from pathlib import Path
from typing import Any, Optional

from common import ls4_source

from . import config, db, source_db
from .logging_utils import get_logger

log = get_logger("coil_results")

TABLE = "コイル実績"
#: 第2パスの取り込みを覚える名前(取り込み履歴)
SECOND_NOTE = "コイル実績(第2パス)"


def sources() -> list[Path]:
    """読むファイル(第1パス → 第2パス)。名前が LS4LOT のパスは入れない。"""
    out: list[Path] = []
    first_name = config.konpo_db_file()
    if ls4_source.is_results_name(first_name):
        found = source_db.find_named(config.konpo_db_dir(), first_name)
        if found is not None:
            out.append(found)
    dirs = config.lot_db_dirs()
    second_name = config.ls4_db2_file()
    if len(dirs) > 1 and ls4_source.is_results_name(second_name):
        found = source_db.find_named(dirs[1], second_name)
        if found is not None and all(_key(found) != _key(p) for p in out):
            out.append(found)
    return out


def wanted() -> bool:
    """実績の形式を読む設定になっているか(どちらかのパスの名前が LS4LOT 以外)。"""
    return (ls4_source.is_results_name(config.konpo_db_file())
            or (len(config.lot_db_dirs()) > 1 and ls4_source.is_results_name(config.ls4_db2_file())))


def _key(path: Path) -> str:
    import os
    return os.path.normcase(os.path.abspath(str(path)))


def missing_columns(path: Path) -> list[str]:
    """そのファイルに実績の形式として足りない列(設定の「確かめる」用)。読めなければ空。"""
    try:
        tables = {t: source_db.columns(Path(path), t) for t in source_db.list_tables(Path(path))}
    except (source_db.SourceError, OSError):
        return []
    if not tables or ls4_source.choose_table(tables, "仕掛"):
        return []
    return ls4_source.missing_required(tables.get("仕掛") or next(iter(tables.values())))


def read_rows(path: Path) -> tuple[list[dict], int]:
    """いまの年月の行(ロットごとに1行)と、ファイルの全行数。要る列が無ければ SourceError。"""
    tables = {t: source_db.columns(Path(path), t) for t in source_db.list_tables(Path(path))}
    picked = ls4_source.choose_table(tables, "仕掛")
    if picked is None:
        best = tables.get("仕掛") or (next(iter(tables.values())) if tables else [])
        raise source_db.SourceError(
            f"{Path(path).name} に実績の形式の列がありません: "
            f"{', '.join(ls4_source.missing_required(best))}")
    rows = source_db.read_table(Path(path), picked)
    return ls4_source.pick_rows(rows), len(rows)


def import_results(conn: Any):
    """コイル実績を取り込む。読む設定でなければ None(画面の結果に出さない)。"""
    from .data_sync import TableResult, source_stamp, _bump   # 循環を避ける

    if not wanted():
        return None
    files = sources()
    if not files:
        return TableResult(table=TABLE, ok=False,
                           error="実績の形式のファイルが見つかりません: "
                                 f"{config.konpo_db_dir()} に {config.konpo_db_file()}"
                                 + (f" / 第2パスに {config.ls4_db2_file()}"
                                    if len(config.lot_db_dirs()) > 1 else ""))
    # 目印は**読む前に**取る(統合 1.2.49。data_sync._import_table と同じ理由)
    stamps = [source_stamp(path) for path in files]
    merged: dict[str, dict] = {}
    filled = 0
    notes: list[str] = []
    errors: list[str] = []
    for index, path in enumerate(files):
        try:
            rows, total = read_rows(path)
        except source_db.SourceError as exc:
            # 片方が読めなくても、もう片方で埋める(統合 1.2.45。以前は先のファイルが読めないと止めた)
            errors.append(str(exc))
            notes.append(f"{path.name} を読めません: {exc}")
            continue
        added = 0
        for row in rows:
            lot = str(row.get(ls4_source.COL_LOT) or "").strip()
            if not lot:
                continue
            if lot not in merged:
                merged[lot] = _entry(row, path)
                added += 1
            elif _fill_from(merged[lot], row, path):
                filled += 1
        year, month = ls4_source.this_month()
        notes.append(f"{path.name}: {year}-{month:02d} の行 {added:,}ロット(全 {total:,}行)"
                     + ("を第2パスから足しました" if index else ""))
    if filled:
        notes.append(f"第1パスで外径・検入数が空だった {filled:,}ロットを第2パスで埋めました")
    if len(errors) == len(files):
        return TableResult(table=TABLE, ok=False, source=str(files[0]), error=" / ".join(errors))
    try:
        ensure_columns(conn)
        with conn:
            conn.execute(f'DELETE FROM "{TABLE}"')
            for lot, e in merged.items():
                conn.execute(
                    f'INSERT INTO "{TABLE}" (ロット番号, 内径, 肉厚, 外径, 枚本数, 作業開始, 元ファイル)'
                    " VALUES (?, ?, ?, ?, ?, ?, ?)",
                    (lot, e["inner"], e["thick"], e["outer"], e["count"], e["start"], e["source"]))
            now = db.now_db_string()
            conn.execute(
                "INSERT OR REPLACE INTO 取り込み履歴"
                " (テーブル名, 取り込み日時, 元ファイル, 件数, 元の目印, 確かめた日時)"
                " VALUES (?, ?, ?, ?, ?, ?)", (TABLE, now, str(files[0]), len(merged), stamps[0], now))
            if len(files) > 1:
                conn.execute(
                    "INSERT OR REPLACE INTO 取り込み履歴"
                    " (テーブル名, 取り込み日時, 元ファイル, 件数, 元の目印, 確かめた日時)"
                    " VALUES (?, ?, ?, ?, ?, ?)", (SECOND_NOTE, now, str(files[1]), 0, stamps[1], now))
    except sqlite3.Error as exc:
        log.exception("コイル実績の取り込みに失敗")
        return TableResult(table=TABLE, ok=False, source=str(files[0]), error=str(exc))
    _bump()                                     # 取り込みの世代を進める(統合 1.2.49)
    log.info("コイル実績: %d ロット (%s)", len(merged), " / ".join(notes))
    return TableResult(table=TABLE, ok=True, rows=len(merged), source=str(files[0]),
                       note="・".join(notes))


def _entry(row: dict, path: Path) -> dict:
    """1 ロットの値(外径・検入数と、どのファイルから取ったか)。"""
    outer = ls4_source.outer_diameter(row)
    return {"inner": _num(row.get(ls4_source.COL_INNER)), "thick": _num(row.get(ls4_source.COL_THICK)),
            "outer": outer, "count": ls4_source.coil_count(row),
            "start": str(row.get(ls4_source.COL_START) or ""), "source": str(path)}


def _fill_from(entry: dict, row: dict, path: Path) -> bool:
    """第1パスで空だった外径・検入数を、第2パスの同じロットの行で埋める。埋めたら True。

    **値のある項目は第1パスのまま**(第2パスで上書きしない)。外径は内径・肉厚をそろえて取る。
    """
    other = _entry(row, path)
    took = []
    if entry["outer"] is None and other["outer"] is not None:
        entry.update(inner=other["inner"], thick=other["thick"], outer=other["outer"])
        took.append("外径")
    if entry["count"] is None and other["count"] is not None:
        entry["count"] = other["count"]
        took.append("検入数")
    if took:
        entry["source"] = f'{entry["source"]} ({"・".join(took)}は {path})'
    return bool(took)


def _num(value: Any) -> Optional[float]:
    try:
        return float(str(value).strip())
    except (TypeError, ValueError):
        return None


def ensure_columns(conn: sqlite3.Connection) -> None:
    """1.2.29 で作った手元の表に、あとから足した列(枚本数。統合 1.2.30)を足す。"""
    try:
        have = {r[1] for r in conn.execute(f'PRAGMA table_info("{TABLE}")')}
        if have and "枚本数" not in have:
            with conn:
                conn.execute(f'ALTER TABLE "{TABLE}" ADD COLUMN 枚本数 INTEGER')
    except sqlite3.Error:
        log.warning("コイル実績に列を足せませんでした", exc_info=True)


def count_for(conn: sqlite3.Connection, lot_no: str) -> Optional[int]:
    """そのロットの検入数(本)(実績から)。無ければ None(いままでどおり手で入れる)。"""
    if not lot_no:
        return None
    try:
        row = conn.execute(f'SELECT 枚本数 FROM "{TABLE}" WHERE ロット番号 = ?',
                           (str(lot_no).strip(),)).fetchone()
    except sqlite3.Error:
        return None
    if row is None or row[0] is None or int(row[0]) <= 0:
        return None
    return int(row[0])


def outer_for(conn: sqlite3.Connection, lot_no: str) -> Optional[float]:
    """そのロットのコイル外径(実績から)。無ければ None(いままでの参照物を使う)。"""
    if not lot_no:
        return None
    try:
        row = conn.execute(f'SELECT 外径 FROM "{TABLE}" WHERE ロット番号 = ?',
                           (str(lot_no).strip(),)).fetchone()
    except sqlite3.Error:
        return None
    if row is None or row[0] is None or float(row[0]) <= 0:
        return None
    return float(row[0])
