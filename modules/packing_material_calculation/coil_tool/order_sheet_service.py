"""発注票 (VBA `転記` → `Ripla_Count` → `注文票入力`)

チェックリストに溜めたものを、**倉庫へ出す形**に組み替える。

    チェックリスト(ロット単位・1セルに2つ)
        │  転記 … 改行で入っている2つを2レコードに戻す
        ▼
    リプラ / LVS ごとの明細(Lot, 種類, 角サイズ, 長さ, 本数)
        │  集計 … 角サイズ + 長さ をキーに本数を足す
        ▼
    発注票(角サイズごとに1枚。長さの行に本数とロットを並べる)

【角サイズは3種類まで】
30×40 / 40×60 / 80×80。VBA も最大3枚しか作らない。
**80×80 は長さを持たない**ので、長さの行は使わない。

【倉庫の定尺に無い長さ】
発注票の原紙は定尺8種+空欄4行でできていて、定尺に無い長さは空欄へ
書き足す。**5件以上あると入らない**ので、そこで断る。

【定尺は梱包資材マスタの表「リプラサイズ」から】(統合 1.2.13。現場の依頼)
VBA は原紙に 700〜1150 の8種を直に書いていた。マスタ管理で新しい長さを足しても発注票に
出るように、表「リプラサイズ」の リプラ長さ を小さい順に並べて行にする(行の数も表に合わせる)。
表が空のときだけ、設定(`stock_lengths`)か原紙の8種を使う。

【LOTで分ける】(資材計算 0.2.35・現場の依頼)
ふだんは同じ角サイズ・同じ長さを1行にまとめるので、1行に2つのロットが並ぶことがある
(例: 900 「L6241W0 L7160C0」 60本)。それだと**ロットごとの本数が票から分からない**。
「LOTで分ける」にすると、**チェックリストの行に戻って**ロットごとに本数を数え直し、
ロットごとに票を起こす(1ロット = 1ページ。同じロットにリプラと LVS、角サイズが2つ
あればそのぶん増える)。

    800 L6241W0 20本 / 900 L6241W0 30本 / 1150 L7160C0 20本
      → L6241W0 の票(800 20本・900 30本)と L7160C0 の票(1150 20本)
"""
from __future__ import annotations

import sqlite3
from dataclasses import dataclass, field
from typing import Iterable, Optional

from . import checklist_service, config, db, order_history, vba
from .logging_utils import get_logger

log = get_logger("order_sheet_service")

LF = "\n"

# 倉庫在庫の定尺(mm)。VBA は発注票の原紙(A5:A12)に直接書いていた。
# **ふだんは梱包資材マスタの表「リプラサイズ」から作る**(`stock_lengths`)。これは表が空のときの予備
DEFAULT_STOCK_LENGTHS: tuple[int, ...] = (700, 800, 900, 950, 1000, 1050, 1100, 1150)

# 定尺に無い長さを書き足せる行数(原紙の A13:A16)
EXTRA_ROWS = 4

# 発注票1枚の行数(原紙: 定尺8 + 書き足し4)。定尺の数が変われば行の数も変わる(`sheet_rows`)
SHEET_ROWS = len(DEFAULT_STOCK_LENGTHS) + EXTRA_ROWS

# 1回で作れる発注票の枚数。角サイズの種類ぶん
MAX_SHEETS = 3

# コイル間に入るもの。この2つだけが発注票になる
KINDS = ("リプラ", "LVS")

REASON_TOO_MANY_SIZES = "too_many_extra_lengths"
REASON_DUPLICATE_LOT = "duplicate_lot"
REASON_EMPTY = "nothing_to_order"


def stock_lengths(conn: Optional[sqlite3.Connection] = None) -> list[float]:
    """いま使う定尺(小さい順・重なりなし)。

    1. 梱包資材マスタの表「リプラサイズ」の リプラ長さ(取り込んだ写し。マスタ管理で直すと
       すぐ取り込み直される)。リプラの長さの選び方(`ripla_service`)と同じ表
    2. 表が空なら設定(`stock_lengths`。設定ファイルで入れたとき)
    3. それも無ければ原紙の8種
    """
    if conn is not None:
        try:
            rows = db.fetch_all(conn, "SELECT リプラ長さ FROM リプラサイズ",
                                caller_name="order_stock_lengths") or []
        except sqlite3.Error as exc:
            log.warning("リプラサイズを読めません(原紙の定尺で作ります): %s", exc)
            rows = []
        found = sorted({vba.val(r["リプラ長さ"]) for r in rows} - {0.0})
        found = [n for n in found if n > 0]
        if found:
            return found
    from . import user_settings
    configured = user_settings.get(config.KEY_STOCK_LENGTHS)
    if isinstance(configured, list) and configured:
        return sorted({vba.val(n) for n in configured if vba.val(n) > 0})
    return [float(n) for n in DEFAULT_STOCK_LENGTHS]


def sheet_rows(lengths: list[float]) -> int:
    """発注票1枚の行数。定尺ぶん + 書き足し4行。"""
    return len(lengths) + EXTRA_ROWS


# ==================================================================
# 転記 ── 1セルに2つ入っているものを2レコードに戻す
# ==================================================================
@dataclass
class Detail:
    """明細1件。"""
    LotNo: str = ""
    種類: str = ""       # リプラ / LVS
    角サイズ: str = ""   # "30×40" など
    長さ: str = ""
    本数: str = ""


def expand(rows: list[dict]) -> list[Detail]:
    """チェックリストの行を明細に開く (VBA `転記`)。

    リプラ長さ・リプラ本数は改行で2つ入っていることがある。
    **改行が無ければ1件のまま**にする ── 以前は必ず2行書いていて、
    改行の無い行が2重に数えられていた(コード冒頭のコメントにある修正)。
    """
    out: list[Detail] = []
    for row in rows:
        kind = str(row.get("コイル間", "") or "")
        if kind not in KINDS:
            continue

        lengths = str(row.get("リプラ長さ", "") or "").split(LF)
        counts = str(row.get("リプラ本数", "") or "").split(LF)
        lot = str(row.get("LotNo", "") or "")
        size = str(row.get("リプラサイズ", "") or "")

        out.append(Detail(LotNo=lot, 種類=kind, 角サイズ=size,
                          長さ=lengths[0].strip() if lengths else "",
                          本数=counts[0].strip() if counts else ""))

        if len(lengths) > 1 or len(counts) > 1:
            out.append(Detail(
                LotNo=lot, 種類=kind, 角サイズ=size,
                長さ=lengths[-1].strip() if len(lengths) > 1 else "",
                本数=counts[-1].strip() if len(counts) > 1 else ""))
    return out


# ==================================================================
# 集計 ── 角サイズ + 長さ で本数を足す
# ==================================================================
@dataclass
class Aggregated:
    角サイズ: str = ""
    長さ: str = ""
    本数: int = 0
    lots: list[str] = field(default_factory=list)


def aggregate(details: Iterable[Detail], kind: str) -> list[Aggregated]:
    """同じ角サイズ・同じ長さのものを1件にまとめる (VBA `Ripla_Count`)。

    ロットは**出てきた順のまま重複を落とす**。発注票のロット欄は
    倉庫が現物と突き合わせるためのものなので、並びが動くと読みにくい。
    """
    table: dict[tuple[str, str], Aggregated] = {}
    for d in details:
        if d.種類 != kind:
            continue
        if not d.長さ and not d.本数:
            continue
        key = (d.角サイズ, d.長さ)
        entry = table.setdefault(key, Aggregated(角サイズ=d.角サイズ, 長さ=d.長さ))
        entry.本数 += int(vba.val(d.本数))
        if d.LotNo and d.LotNo not in entry.lots:
            entry.lots.append(d.LotNo)
    return list(table.values())


# ==================================================================
# 発注票を組む
# ==================================================================
@dataclass
class SheetRow:
    長さ: str = ""
    LotNo: str = ""
    数量: str = ""


@dataclass
class OrderSheet:
    """発注票1枚。"""
    種類: str = ""        # リプラ / LVS
    角サイズ: str = ""    # "30×40" など
    提出日付: str = ""
    依頼者: str = ""
    rows: list[SheetRow] = field(default_factory=list)
    # 「LOTで分ける」で作ったときの、その票のロット(まとめて作ったときは空)
    ロット: str = ""

    @property
    def has_content(self) -> bool:
        return any(r.数量 for r in self.rows)


@dataclass
class BuildResult:
    sheets: list[OrderSheet] = field(default_factory=list)
    duplicates: list[order_history.Duplicate] = field(default_factory=list)
    ok: bool = False
    reason: str = ""
    # 定尺に無くて書き足した長さ。画面に出して倉庫へ伝える
    extra_lengths: list[str] = field(default_factory=list)
    # 発注履歴を読めず、2週間以内の重複を確かめていない(統合 1.2.61)
    history_unchecked: bool = False


def _sizes_in(entries: list[Aggregated]) -> list[str]:
    """出てきた角サイズを、出てきた順のまま重複なしで。最大3種。"""
    seen: list[str] = []
    for e in entries:
        if e.角サイズ and e.角サイズ not in seen:
            seen.append(e.角サイズ)
    return seen[:MAX_SHEETS]


def _build_sheet(entries: list[Aggregated], *, kind: str, size: str,
                 worker: str, day: str,
                 lengths: Optional[list[float]] = None) -> tuple[Optional[OrderSheet], list[str], str]:
    """角サイズ1つぶんの発注票。戻り値は `(票, 書き足した長さ, 断る理由)`。"""
    mine = [e for e in entries if e.角サイズ == size]
    sheet = OrderSheet(種類=kind, 角サイズ=size, 提出日付=day, 依頼者=worker)

    # 80×80 は長さを持たない。枠だけ作って長さ行は使わない
    if lengths is None:
        lengths = stock_lengths()
    total = sheet_rows(lengths)
    if size == "80×80":
        sheet.rows = [SheetRow() for _ in range(total)]
        return sheet, [], ""

    wanted = [e.長さ for e in mine if e.長さ]
    known = set(lengths)
    extra = [n for n in dict.fromkeys(wanted) if vba.val(n) not in known]

    if len(extra) > EXTRA_ROWS:
        return None, extra, REASON_TOO_MANY_SIZES

    slots = [vba.num_text(n) for n in lengths] + extra
    slots += [""] * (total - len(slots))

    rows: list[SheetRow] = []
    for slot in slots[:total]:
        row = SheetRow(長さ=slot)
        if slot:
            for e in mine:
                if e.長さ and vba.val(e.長さ) == vba.val(slot):
                    # 0 本は空欄にする。VBA も「空白をサイズとして
                    # とらえてしまうため」0 を消していた (2022.09.21)
                    row.数量 = str(e.本数) if e.本数 else ""
                    row.LotNo = " ".join(e.lots)
        rows.append(row)
    sheet.rows = rows
    return sheet, extra, ""


def lots_in(details: Iterable[Detail]) -> list[str]:
    """発注票になる明細のロットを、チェックリストに出てきた順のまま重複なしで。"""
    seen: list[str] = []
    for d in details:
        if d.種類 in KINDS and (d.長さ or d.本数) and d.LotNo not in seen:
            seen.append(d.LotNo)
    return seen


def build(conn: sqlite3.Connection, *, worker: str,
          check_history: bool = True, by_lot: bool = False) -> BuildResult:
    """チェックリストから発注票を起こす。

    **履歴の重複はここで返すだけ**で、止めはしない。続けるかどうかは
    人が決める(VBA も「このまま重複して発注しますか？」と聞いていた)。
    続けるときは `check_history=False` で呼び直す。

    `by_lot=True`(「LOTで分ける」)… ロットごとに票を起こす。本数はチェックリストの
    そのロットの行だけで数え直す(まとめた行からは割り振れないので)。
    """
    rows = checklist_service.load_all(conn)
    if not rows:
        return BuildResult(ok=False, reason=REASON_EMPTY)

    details = expand(rows)
    day = checklist_service.today_label()
    lengths = stock_lengths(conn)

    groups: list[tuple[str, list[Detail]]]
    if by_lot:
        groups = [(lot, [d for d in details if d.LotNo == lot]) for lot in lots_in(details)]
    else:
        groups = [("", details)]

    result = BuildResult()
    for lot, mine in groups:
        for kind in KINDS:
            entries = aggregate(mine, kind)
            if not entries:
                continue
            for size in _sizes_in(entries):
                sheet, extra, reason = _build_sheet(
                    entries, kind=kind, size=size, worker=worker, day=day, lengths=lengths)
                if reason:
                    return BuildResult(ok=False, reason=reason, extra_lengths=extra)
                if sheet is not None and (sheet.has_content or size == "80×80"):
                    sheet.ロット = lot if by_lot else ""
                    result.sheets.append(sheet)
                    result.extra_lengths.extend(n for n in extra if n not in result.extra_lengths)

    if not result.sheets:
        return BuildResult(ok=False, reason=REASON_EMPTY)

    if check_history:
        lots = [r.LotNo for s in result.sheets for r in s.rows if r.LotNo]
        try:
            result.duplicates = order_history.find_duplicates(conn, lots)
        except order_history.HistoryUnreadable as exc:
            # 読めなかったのを「重複なし」にしない。票は出すが、確かめていないことを画面に出す(統合 1.2.61)
            log.warning("発注履歴を確かめられませんでした: %s", exc)
            result.history_unchecked = True

    result.ok = True
    return result


def commit(conn: sqlite3.Connection, sheets: list[OrderSheet], *,
           worker: str) -> int:
    """出した内容を履歴へ残す (VBA `SaveToHistory` の呼び出し)。

    **長さと数量の両方がある行だけ**。空欄の行を残しても、次の重複
    チェックで引っかかるだけで意味が無い。
    """
    day = checklist_service.today_label()
    entries = [
        {"処理日": day, "担当者": worker, "品名": s.種類, "サイズ": s.角サイズ,
         "長さ": r.長さ, "LotNo": r.LotNo, "数量": r.数量}
        for s in sheets for r in s.rows
        if str(r.LotNo).strip() and str(r.数量).strip()
    ]
    return order_history.save_many(conn, entries)
