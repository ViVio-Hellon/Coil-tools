"""LS4 の仕掛ファイルの2つの形式(統合 1.2.29・現場の指示)。**決まりはここ1か所。**

梱包明細(縦割数・横割数 → 前工程実績数・丈数)と資材計算(コイル外径)が使う。
仕様と列の対応は `docs/仕掛ファイルの仕様.md`。

**ファイル名で形式を決める。**

    名前が LS4LOT(LS4LOT.sqlite3 など)  … いままでの形式(梱包課共有の仕掛)
        ﾛｯﾄ番号 / 当工程設計_縦割数 / 当工程設計_横割数
    名前が LS4LOT 以外(例: 実績_LS4.sqlite3) … 実績の形式
        ﾛｯﾄ番号
        BOX設計_縦割数 = 当工程設計_縦割数
        BOX設計_横割数 = 当工程設計_横割数
        BOX実績_作業開始_日付 … **いまの年月の行だけ**使う(2026-10-21 00:01:56 → 2026-10)
        BOX実績_ｺｲﾙ内径・BOX実績_ｺｲﾙ肉厚 … コイル外径 = 内径 + 2 × 肉厚(資材計算)
        BOX実績_枚本数 … 検入数(本)(資材計算)。無ければ BOX設計_横割数 × BOX設計_縦割数

同じロットが何行もあれば、作業開始がいちばん新しい行を使う。
"""
from __future__ import annotations

import re
from datetime import date
from pathlib import Path
from typing import Any, Iterable, Optional

#: いままでの形式のファイル名(拡張子を除く。大文字小文字は見ない)
CLASSIC_STEM = "LS4LOT"

COL_LOT = "ﾛｯﾄ番号"
# いままでの形式
COL_TATE = "当工程設計_縦割数"
COL_YOKO = "当工程設計_横割数"
# 実績の形式
COL_BOX_TATE = "BOX設計_縦割数"
COL_BOX_YOKO = "BOX設計_横割数"
COL_START = "BOX実績_作業開始_日付"
COL_INNER = "BOX実績_ｺｲﾙ内径"
COL_THICK = "BOX実績_ｺｲﾙ肉厚"
COL_COUNT = "BOX実績_枚本数"

#: 実績の形式で**無いと読まない**列(ロット・割数・月の判定)
RESULTS_REQUIRED = (COL_LOT, COL_BOX_TATE, COL_BOX_YOKO, COL_START)
#: 実績の形式で、あれば使う列(コイル外径・検入数)
RESULTS_OPTIONAL = (COL_INNER, COL_THICK, COL_COUNT)

_MONTH = re.compile(r"^\s*(\d{4})[-/.](\d{1,2})")


def is_results_name(name: Any) -> bool:
    """そのファイル(名前)は実績の形式か。名前が LS4LOT なら いままでの形式(False)。"""
    stem = Path(str(name or "")).stem.strip()
    return bool(stem) and stem.upper() != CLASSIC_STEM


def year_month(text: Any) -> Optional[tuple[int, int]]:
    """`2026-10-21 00:01:56` → (2026, 10)。`/` 区切りも読む。読めなければ None。"""
    m = _MONTH.match(str(text or ""))
    if not m:
        return None
    month = int(m.group(2))
    if not 1 <= month <= 12:
        return None
    return int(m.group(1)), month


def this_month(today: Optional[date] = None) -> tuple[int, int]:
    today = today or date.today()
    return today.year, today.month


def pick_rows(rows: Iterable[dict], today: Optional[date] = None) -> list[dict]:
    """実績の行から、**いまの年月の行だけ**を、ロットごとに1行(作業開始がいちばん新しい行)。"""
    want = this_month(today)
    best: dict[str, dict] = {}
    for row in rows:
        if year_month(row.get(COL_START)) != want:
            continue
        lot = str(row.get(COL_LOT) or "").strip()
        if not lot:
            continue
        known = best.get(lot)
        if known is None or str(row.get(COL_START) or "") > str(known.get(COL_START) or ""):
            best[lot] = row
    return list(best.values())


def _number(value: Any) -> Optional[float]:
    try:
        n = float(str(value).strip())
    except (TypeError, ValueError):
        return None
    return n if n == n else None              # NaN を落とす


def outer_diameter(row: dict) -> Optional[float]:
    """コイル外径 = 内径 + 2 × 肉厚。どちらかが無い・0 以下なら None(いままでの参照物を使う)。"""
    inner = _number(row.get(COL_INNER))
    thick = _number(row.get(COL_THICK))
    if inner is None or thick is None or inner <= 0 or thick <= 0:
        return None
    return round(inner + 2 * thick, 2)


def coil_count(row: dict) -> Optional[int]:
    """検入数(本) = BOX実績_枚本数。無い・0 以下なら BOX設計_横割数 × BOX設計_縦割数(同じ数のはず)。
    どちらも出せなければ None(いままでどおり手で入れる)。"""
    count = _number(row.get(COL_COUNT))
    if count is not None and count > 0 and count == int(count):
        return int(count)
    tate, yoko = _number(row.get(COL_BOX_TATE)), _number(row.get(COL_BOX_YOKO))
    if tate and yoko and tate > 0 and yoko > 0 and tate == int(tate) and yoko == int(yoko):
        return int(tate) * int(yoko)
    return None


def choose_table(tables: dict[str, Iterable[str]], preferred: str = "") -> Optional[str]:
    """実績のファイルの中で、要る列(`RESULTS_REQUIRED`)をそろえた表。`preferred` を先に見る。"""
    order = ([preferred] if preferred in tables else []) + [t for t in tables if t != preferred]
    for name in order:
        if set(RESULTS_REQUIRED) <= set(tables[name]):
            return name
    return None


def missing_required(columns: Iterable[str]) -> list[str]:
    """実績の形式として足りない列。"""
    have = set(columns)
    return [c for c in RESULTS_REQUIRED if c not in have]
