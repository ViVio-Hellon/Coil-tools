"""ロット番号から、検査番号(仕掛ロットなど)とコイル外径(LS4LOT)を引く(統合 1.2.57・1.2.58・現場の指示)

現場の指示:
- ペナラベルに LOTNO の入力欄を足す。LOTNO を探すデータは**コイル梱包明細打ち出しシステム**を参照する。
  ﾛｯﾄ番号がヒットしたら、**同じ行の検査番号**を新検番に入れる。そのとき外径計算も行う
- 資材計算のコイル外径は、データにあっても計算する(データが無ければ常に計算する)

そこで、梱包明細が読むファイルを**列で**読む。

- **検査番号**: 梱包明細が読む4ファイルのうち `ﾛｯﾄ番号` の列を持つもの ── 仕掛ロット(SIKALOT)→ 仕掛引当(SIKAHIKI)
  → 仕掛当工程(LS4LOT)の順に、それぞれ第1パス → 第2パス。**LS4LOT には検査番号の列が無い**(現場の確認・統合 1.2.58)
  ので、仕掛台帳の側を先に見る。仕掛受注は受注番号の表なので見ない。列名は SIKALOT・SIKAHIKI とも「検査番号」
  (統合 1.2.59 で実ファイルで確かめた)。SIKALOT は大半の行が「0000000000」なので、番号の入っている行から取る
- **コイル外径**: 仕掛当工程の**第2パス(「LS4LOT の代わり」。実績_LS4)→ 第1パス(梱包課共有 LS4LOT)**の
  `BOX実績_ｺｲﾙ内径`・`BOX実績_ｺｲﾙ肉厚`(統合 1.2.60 から第2パスを先に見る。現場の指示)。
  以前(統合 1.2.29〜1.2.56)は**ファイルの名前**で形式を決め、名前が LS4LOT のままなら内径・肉厚の列を
  読まなかった。列があれば名前に関わらず外径を出す。実際の LS4LOT には内径・肉厚の列が無く、第2パスの実績_LS4 から出る

- 表は `ﾛｯﾄ番号` の列を持つもの(`仕掛` という表があればそれを先に見る)
- 同じロットが何行もあれば、`BOX実績_作業開始_日付` がいちばん新しい行(その列が無ければ最後の行)
- 検査番号の列は名前で探す。半角・全角をそろえてから(NFKC)、`検査番号`・`検査NO` → 名前に「検査番号」「検査NO」
  「検番」「ケンサ」を含む列。どのファイルにも見つからなければそう知らせ、名前に「検」を含む列があれば添える
- 外径 = 内径 + 2 × 肉厚(`common/ls4_source.outer_diameter`。どちらも 0 より大きいときだけ)
- 外径は第2パスを先に見て、出せなければ第1パスの同じロットの行で出す
- 共有のファイルは手元へ写してから読む(`source_db` が写す)。ファイルの更新日時と大きさが同じあいだは覚えておく
"""
from __future__ import annotations

import os
import threading
import unicodedata
from pathlib import Path
from typing import Any, Optional

from common import ls4_source

from . import config, data_sync, source_db
from .logging_utils import get_logger

log = get_logger("lot_lookup")

#: 検査番号の列の名前(NFKC でそろえ、英字は大文字にしてから比べる。この順で探す)
KENSA_EXACT = ("検査番号", "検査NO")
KENSA_PARTS = ("検査番号", "検査NO", "検番", "ケンサ")

#: 検査番号を探す表(梱包明細の手元の表名。この順)。仕掛受注は受注番号の表なので入れない
KENSA_TABLES = ("仕掛ロット", "仕掛引当", "仕掛当工程")
#: コイル外径(内径・肉厚)を探す表
OUTER_TABLE = "仕掛当工程"

_lock = threading.Lock()
#: (ファイル, 更新日時 ns, 大きさ) → (表, 列)
_tables: dict[tuple, tuple[str, list[str]]] = {}


def sources(table: str = OUTER_TABLE) -> list[Path]:
    """その表のファイル(第1パス → 第2パス)。梱包明細が取り込みで読むものと同じ。"""
    out: list[Path] = []
    first = data_sync.find_sources().get(table)
    if first is not None:
        out.append(Path(first))
    second = data_sync.find_second_lots().get(table)
    if second is not None and all(not _same(second, p) for p in out):
        out.append(Path(second))
    return out


def _norm(name: Any) -> str:
    return unicodedata.normalize("NFKC", str(name)).upper()


def _same(a: Path, b: Path) -> bool:
    return os.path.normcase(os.path.abspath(a)) == os.path.normcase(os.path.abspath(b))


def kensa_column(columns: list[str]) -> Optional[str]:
    """検査番号の列の名前。無ければ None。"""
    normed = [(_norm(c), c) for c in columns]
    for want in KENSA_EXACT:
        for n, c in normed:
            if n == want:
                return c
    for part in KENSA_PARTS:
        for n, c in normed:
            if part in n:
                return c
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
    # 読めないのを「表が無い」と取り違えて覚えない(統合 1.2.61。別のツールで、読めなかったのを「無い」と
    # 取り違えて送り直し、重複を大量に作った事案があった)。読めなければ SourceError を上げ、覚えない
    tables = source_db.list_tables_strict(path)
    order = sorted(tables, key=lambda t: (t != "仕掛", t))
    for table in order:
        cols = source_db.columns_strict(path, table)
        if ls4_source.COL_LOT in cols:
            found = (table, cols)
            break
    with _lock:
        for old in [k for k in _tables if k[0] == key[0]]:
            del _tables[old]
        _tables[key] = found
    return found


def _rows_in(path: Path, lot: str) -> tuple[list[dict], list[str]]:
    """そのファイルの、そのロットの行(古い順。いちばん新しいものが最後)と、表の列。"""
    table, cols = _table_of(path)
    if not table:
        return [], []
    rows = source_db.read_query(
        path, f"SELECT * FROM {source_db.quote_identifier(table)} "
              f"WHERE TRIM({source_db.quote_identifier(ls4_source.COL_LOT)}) = ?", (lot,))
    if ls4_source.COL_START in cols:
        rows.sort(key=lambda r: str(r.get(ls4_source.COL_START) or ""))
    return rows, cols


def kensa_value(value: Any) -> str:
    """検査番号として使える値(大文字)。空・**0 だけ**(「0000000000」= まだ番号が無い)なら ""。

    実データ(2026-10 の SIKALOT)は 11,450 行のうち 9,754 行が「0000000000」。同じロットの何行かのうち、
    番号が入っているのは1行だけだった。
    """
    text = _text(value).upper()
    return "" if not text or set(text) == {"0"} else text


def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def lookup(lot: str, *, kensa: bool = True) -> dict:
    """ロット番号の行から、検査番号とコイル外径を引く。

    `kensa=False` … 外径だけ(資材計算)。検査番号のファイル(仕掛ロット・仕掛引当)は読まない。

    戻り値:
        found    どれかのファイルにそのロットの行があったか
        kensa    検査番号(無ければ "")
        kensa_file  検査番号を見つけたファイル
        outer    コイル外径(内径 + 2 × 肉厚。出せなければ None)
        inner / thick  内径・肉厚(元の値)
        file     外径の行を見つけたファイル(無ければ最初に行を見つけたファイル)
        notes    知らせる文の並び(検査番号の列が無い・外径を出せない など)
        reason   見つからないときの理由
    """
    lot = _text(lot).upper()
    out: dict[str, Any] = {"found": False, "lot": lot, "kensa": "", "kensa_file": "", "outer": None,
                           "inner": None, "thick": None, "file": "", "notes": [], "reason": ""}
    if not lot:
        out["reason"] = "LOTNO が空です"
        return out
    problems: list[str] = []
    seen: list[Path] = []
    looked: list[str] = []          # 検査番号の列が無かったファイル
    empty = False                   # 検査番号の列はあったが、そのロットの行では空だった
    hints: list[str] = []           # 名前に「検」を含む列(どれが検査番号かを聞くため)
    rows: dict[str, tuple[Optional[dict], list[str]]] = {}

    def row_of(path: Path):
        key = os.path.normcase(os.path.abspath(path))
        if key not in rows:
            try:
                rows[key] = _rows_in(path, lot)
            except Exception as exc:   # noqa: BLE001 - 次のファイルへ
                log.warning("仕掛のファイルを読めません: %s", path, exc_info=True)
                problems.append(f"{path.name} を読めません({exc})")
                rows[key] = ([], [])
        return rows[key]

    # 検査番号: 仕掛ロット → 仕掛引当 → 仕掛当工程(それぞれ第1パス → 第2パス)
    for table in (KENSA_TABLES if kensa else ()):
        for path in sources(table):
            if any(_same(path, p) for p in seen):
                continue
            seen.append(path)
            found_rows, cols = row_of(path)
            if not found_rows:
                continue
            if not out["found"]:
                out.update(found=True, file=str(path))
            col = kensa_column(cols)
            if col is None:
                log.info("検査番号の列がありません: %s(列: %s)", path.name, "・".join(map(str, cols)))
                looked.append(path.name)
                hints += [c for c in cols if "検" in _norm(c) and c not in hints]
                continue
            # 番号の入っている行のうち、いちばん新しいもの(「0000000000」の行は飛ばす)
            values = [v for v in (kensa_value(r.get(col)) for r in found_rows) if v]
            if values:
                out.update(kensa=values[-1], kensa_file=str(path))
                break
            empty = True
        if out["kensa"]:
            break

    # コイル外径: 仕掛当工程の**第2パス → 第1パス**(統合 1.2.60・現場の指示「外径は初めから第2パスで見に行ってよい」。
    # 実際の LS4LOT(第1パス)には内径・肉厚の列が無く、第2パスの実績_LS4 にある)
    outer_files = list(reversed(sources(OUTER_TABLE)))
    for path in outer_files:
        found_rows, _cols = row_of(path)
        if not found_rows:
            continue
        if not out["found"]:
            out.update(found=True)
        # 内径・肉厚の入っている行のうち、いちばん新しいもの
        hit = next((r for r in reversed(found_rows) if ls4_source.outer_diameter(r) is not None), None)
        if hit is not None:
            out.update(outer=ls4_source.outer_diameter(hit), inner=hit.get(ls4_source.COL_INNER),
                       thick=hit.get(ls4_source.COL_THICK), file=str(path))
            break

    if not out["found"]:
        if not seen and not outer_files:
            out["reason"] = ("梱包明細の取り込み元(仕掛ロット・仕掛引当・"
                             f"{config.konpo_db_files().get('仕掛当工程', 'LS4LOT.sqlite3')})"
                             "が見つかりません。梱包明細の設定(置き場所・取り込み)を確かめてください")
        else:
            out["reason"] = "・".join(problems) or f"梱包明細の取り込み元にロット {lot} がありません"
        return out
    if kensa and not out["kensa"]:
        if not empty:
            text = f"検査番号の列が見つかりません(見たファイル: {'・'.join(looked) or 'なし'})"
            if hints:
                text += f"。名前に「検」を含む列: {'・'.join(str(h) for h in hints[:8])}"
            out["notes"].append(text)
        else:
            out["notes"].append(f"ロット {lot} の検査番号はまだ入っていません(空か 0000000000)")
    if out["outer"] is None:
        out["notes"].append(f"ロット {lot} の内径・肉厚が入っていないので、コイル外径を出せません")
    return out


def forget() -> None:
    """覚えた表の形を捨てる(試験用)。"""
    with _lock:
        _tables.clear()
