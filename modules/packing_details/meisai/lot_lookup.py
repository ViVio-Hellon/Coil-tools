"""ロット番号から、仕掛(LS4LOT)の同じ行の検査番号・コイル外径を引く(統合 1.2.57・現場の指示)

現場の指示:
- ペナラベルに LOTNO の入力欄を足す。LOTNO を探すデータは**コイル梱包明細打ち出しシステム**を参照する。
  ﾛｯﾄ番号がヒットしたら、**同じ行の検査番号**を新検番に入れる。そのとき外径計算も行う
- 資材計算のコイル外径は、データにあっても計算する(データが無ければ常に計算する)

そこで、梱包明細が読む仕掛のファイル(第1パスの梱包課共有 LS4LOT → 第2パス)を**列で**読む。
以前(統合 1.2.29〜1.2.56)は**ファイルの名前**で形式を決め、名前が LS4LOT のままなら内径・肉厚の列を
読まなかった。ファイルに `BOX実績_ｺｲﾙ内径`・`BOX実績_ｺｲﾙ肉厚` の列があれば、名前に関わらず外径を出す。

- 表は `ﾛｯﾄ番号` の列を持つもの(`仕掛` という表があればそれを先に見る)
- 同じロットが何行もあれば、`BOX実績_作業開始_日付` がいちばん新しい行(その列が無ければ最後の行)
- 検査番号の列は名前で探す(`検査番号` → 名前に「検査番号」を含む列 → 「検番」を含む列)。列が見つからなければ
  そう知らせる(どんな列があったかも添える)
- 外径 = 内径 + 2 × 肉厚(`common/ls4_source.outer_diameter`。どちらも 0 より大きいときだけ)
- 第1パスの行に検査番号・外径が無ければ、第2パスの同じロットの行で**空の項目だけ**埋める(統合 1.2.45 と同じ考え)
- 共有のファイルは手元へ写してから読む(`source_db` が写す)。ファイルの更新日時と大きさが同じあいだは覚えておく
"""
from __future__ import annotations

import os
import threading
from pathlib import Path
from typing import Any, Optional

from common import ls4_source

from . import config, data_sync, source_db
from .logging_utils import get_logger

log = get_logger("lot_lookup")

#: 検査番号の列の名前(この順で探す)
KENSA_EXACT = ("検査番号", "検査NO", "検査No", "検査ＮＯ")
KENSA_PARTS = ("検査番号", "検番")

_lock = threading.Lock()
#: (ファイル, 更新日時 ns, 大きさ) → (表, 列)
_tables: dict[tuple, tuple[str, list[str]]] = {}


def sources() -> list[Path]:
    """読むファイル(第1パス → 第2パス)。梱包明細の「仕掛当工程」(LS4LOT)と同じもの。"""
    out: list[Path] = []
    first = data_sync.find_sources().get("仕掛当工程")
    if first is not None:
        out.append(Path(first))
    second = data_sync.find_second_lots().get("仕掛当工程")
    if second is not None and all(not _same(second, p) for p in out):
        out.append(Path(second))
    return out


def _same(a: Path, b: Path) -> bool:
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def kensa_column(columns: list[str]) -> Optional[str]:
    """検査番号の列の名前。無ければ None。"""
    for name in KENSA_EXACT:
        if name in columns:
            return name
    for part in KENSA_PARTS:
        for name in columns:
            if part in str(name):
                return name
    return None


def _table_of(path: Path) -> tuple[str, list[str]]:
    """ﾛｯﾄ番号の列を持つ表と、その列。無ければ ("", [])。"""
    st = os.stat(path)
    key = (str(path), st.st_mtime_ns, st.st_size)
    with _lock:
        hit = _tables.get(key)
    if hit is not None:
        return hit
    found = ("", [])
    tables = source_db.list_tables(path)
    order = sorted(tables, key=lambda t: (t != "仕掛", t))
    for table in order:
        cols = source_db.columns(path, table)
        if ls4_source.COL_LOT in cols:
            found = (table, cols)
            break
    with _lock:
        for old in [k for k in _tables if k[0] == key[0]]:
            del _tables[old]
        _tables[key] = found
    return found


def _row_in(path: Path, lot: str) -> tuple[Optional[dict], list[str]]:
    """そのファイルの、そのロットの行(いちばん新しいもの)と、表の列。"""
    table, cols = _table_of(path)
    if not table:
        return None, []
    rows = source_db.read_query(
        path, f"SELECT * FROM {source_db.quote_identifier(table)} "
              f"WHERE TRIM({source_db.quote_identifier(ls4_source.COL_LOT)}) = ?", (lot,))
    if not rows:
        return None, cols
    if ls4_source.COL_START in cols:
        rows.sort(key=lambda r: str(r.get(ls4_source.COL_START) or ""))
    return rows[-1], cols


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def lookup(lot: str) -> dict:
    """ロット番号の行から、検査番号とコイル外径を引く。

    戻り値:
        found    そのロットの行があったか
        kensa    検査番号(無ければ "")
        outer    コイル外径(内径 + 2 × 肉厚。出せなければ None)
        inner / thick  内径・肉厚(元の値)
        file     行を見つけたファイル
        notes    知らせる文の並び(検査番号の列が無い・外径を出せない など)
        reason   見つからないときの理由
    """
    lot = _text(lot).upper()
    out: dict[str, Any] = {"found": False, "lot": lot, "kensa": "", "outer": None,
                           "inner": None, "thick": None, "file": "", "notes": [], "reason": ""}
    if not lot:
        out["reason"] = "LOTNO が空です"
        return out
    files = sources()
    if not files:
        out["reason"] = (f"仕掛のファイル({config.konpo_db_files().get('仕掛当工程', 'LS4LOT.sqlite3')})"
                         "が見つかりません。梱包明細の設定(置き場所・取り込み)を確かめてください")
        return out
    problems: list[str] = []
    kensa_missing_cols: list[str] = []
    for path in files:
        try:
            row, cols = _row_in(path, lot)
        except (source_db.SourceError, OSError, Exception) as exc:   # noqa: BLE001 - 次のファイルへ
            log.warning("仕掛のファイルを読めません: %s", path, exc_info=True)
            problems.append(f"{path.name} を読めません({exc})")
            continue
        if row is None:
            continue
        if not out["found"]:
            out.update(found=True, file=str(path))
        if not out["kensa"]:
            col = kensa_column(cols)
            if col is None:
                kensa_missing_cols = cols
            else:
                out["kensa"] = _text(row.get(col)).upper()
        if out["outer"] is None:
            outer = ls4_source.outer_diameter(row)
            if outer is not None:
                out.update(outer=outer, inner=row.get(ls4_source.COL_INNER),
                           thick=row.get(ls4_source.COL_THICK))
        if out["kensa"] and out["outer"] is not None:
            break
    if not out["found"]:
        out["reason"] = "・".join(problems) or f"仕掛のファイルにロット {lot} がありません"
        return out
    if not out["kensa"]:
        if kensa_missing_cols:
            shown = "・".join(str(c) for c in kensa_missing_cols[:12])
            out["notes"].append(f"検査番号の列が見つかりません(列: {shown}…)")
        else:
            out["notes"].append(f"ロット {lot} の行の検査番号が空です")
    if out["outer"] is None:
        out["notes"].append(f"ロット {lot} の内径・肉厚が入っていないので、コイル外径を出せません")
    return out


def forget() -> None:
    """覚えた表の形を捨てる(試験用)。"""
    with _lock:
        _tables.clear()
