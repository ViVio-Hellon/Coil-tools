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
        BOX実績_作業開始_日付 … 同じロットは**いちばん新しい行**(月を問わない。統合 1.2.80。以前はいまの年月の行だけ)
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


#: 何か月前までの行を使うか(統合 1.2.80・現場の指示)。None なら月を問わない。
#: 以前は**いまの年月の行だけ**だったので、先月 LS4 を通ったロットは割数・外径・検入数が出なかった。
#: 実ファイル(2026-08〜11 の 769 行・765 ロット)では、月を問わなくても読み込みの時間は変わらなかった
#: (読む 83ms・選ぶ 1ms 未満)。大きくなって遅くなったら、ここで絞る(例: 1 = 今月と先月)
MONTHS_BACK: Optional[int] = None


def _start_key(row: dict) -> tuple:
    """新しさの順。年月で比べ、同じ年月なら作業開始の文字(区切りを - にそろえる)で比べる。読めない日付はいちばん古い。"""
    text = str(row.get(COL_START) or "").strip()
    return (year_month(text) or (0, 0), text.replace("/", "-"))


def pick_rows(rows: Iterable[dict], today: Optional[date] = None,
              months_back: Optional[int] = None) -> list[dict]:
    """実績の行から、ロットごとに1行(**作業開始がいちばん新しい行** = いまの月に近いほう)。

    `months_back` を渡せば、いまの年月からその月数前までの行だけ(0 = 今月だけ)。省けば `MONTHS_BACK`。
    """
    limit = MONTHS_BACK if months_back is None else months_back
    if limit is not None:
        y, m = this_month(today)
        first = (y * 12 + m - 1) - limit                  # 使う最初の月(月の通し番号)
    best: dict[str, dict] = {}
    for row in rows:
        if limit is not None:
            ym = year_month(row.get(COL_START))
            if ym is None or not first <= ym[0] * 12 + ym[1] - 1 <= first + limit:
                continue
        lot = str(row.get(COL_LOT) or "").strip()
        if not lot:
            continue
        known = best.get(lot)
        if known is None or _start_key(row) > _start_key(known):
            best[lot] = row
    return list(best.values())


def period_text(today: Optional[date] = None) -> str:
    """どの月の行を使ったか(取り込みの知らせ用)。"""
    if MONTHS_BACK is None:
        return "月を問わず、ロットごとにいちばん新しい行"
    y, m = this_month(today)
    if MONTHS_BACK == 0:
        return f"{y}-{m:02d} の行だけ"
    return f"{y}-{m:02d} から {MONTHS_BACK} か月前までの行"


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
