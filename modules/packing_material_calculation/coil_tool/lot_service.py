"""ロットと受注を引く (VBA `LOT検索` / `フォーム展開`)

    LotNo(7桁) ──仕掛引当──▶ 受注番号の候補
    受注番号(8桁) ──仕掛受注──▶ 計算に要る値をぜんぶ

【候補は「両方にあるもの」だけ】
VBA は仕掛引当の行と仕掛受注の行を突き合わせて、**どちらにも居る
受注番号だけ**をコンボへ入れていました。引当はあるが受注が消えている
番号を選ばせても、そのあと何も展開できないためです。

【`"データなし"` をそのまま持ち込む】
値が 0 の欄には文字列 `"データなし"` を入れます。`台数計算` が
`IsNumeric` で「指定なし」と読むので、数値へ寄せると優先順位が変わります。
"""
from __future__ import annotations

import sqlite3
from typing import Optional

from . import coil_results, db, vba
from .logging_utils import get_logger
from .models import CalcState, OrderInfo

log = get_logger("lot_service")

# VBA の入力制限。LotNo は7桁、受注番号は8桁でそれぞれ検索が走る
LOT_NO_LENGTH = 7
ORDER_NO_LENGTH = 8

# 包装仕様NO は先頭6桁だけを使う (VBA `Left(..., 6)`)
SPEC_NO_LENGTH = 6


def normalize_lot_no(text: str) -> str:
    """LotNo の入力整形 (VBA `Lot_KeyPress`)。

    英数字だけを通し、小文字は大文字に直す。現場は英字を小文字で
    打つことがあり、そのままだと引当に当たらない。
    """
    out = []
    for ch in str(text or ""):
        if ch.isdigit() or ("A" <= ch <= "Z"):
            out.append(ch)
        elif "a" <= ch <= "z":
            out.append(ch.upper())
    return "".join(out)


def find_order_numbers(conn: sqlite3.Connection, lot_no: str) -> list[str]:
    """LotNo から受注番号の候補を返す (VBA `LOT検索`)。

    **仕掛引当と仕掛受注の両方にある番号だけ。** 重複は落とす
    (VBA は引当の行数だけコンボへ追加していたので、同じ番号が
    複数回並ぶことがあった)。
    """
    lot_no = normalize_lot_no(lot_no)
    if not lot_no:
        return []
    rows = db.fetch_all(
        conn,
        "SELECT DISTINCT h.受注番号 AS no FROM 仕掛引当 h"
        " JOIN 仕掛受注 o ON o.受注番号 = h.受注番号"
        " WHERE h.ロット番号 = ? ORDER BY h.受注番号",
        (lot_no,), caller_name="find_order_numbers") or []
    return [r["no"] for r in rows]


def shipping_date(conn: sqlite3.Connection, lot_no: str) -> str:
    """そのロットの出荷日 (VBA `LOT検索` の `.出荷`)。

    **あとで `expand_order` が営業納期で上書きする。** VBA も同じ順で
    書いていて、帳票の「営業納期」欄に出るのは受注側の営業納期のほう。
    ここはロットだけ引いた段階の表示に使う。
    """
    lot_no = normalize_lot_no(lot_no)
    if not lot_no:
        return ""
    row = db.fetch_one(
        conn, "SELECT 出荷日 FROM 仕掛引当 WHERE ロット番号 = ? LIMIT 1",
        (lot_no,), caller_name="shipping_date")
    return str(row["出荷日"]) if row is not None else ""


def _or_no_data(value: float, decimals: Optional[int] = None) -> str:
    """0 なら「データなし」、それ以外は値。

    VBA `フォーム展開` の
    `If TempHiki(1, C_…) <> 0 Then … Else "データなし"` の移植。
    """
    if value == 0:
        return OrderInfo.NO_DATA
    # 小数を切り捨てない(統合 1.2.6)。VBA は値をそのままテキストボックスへ入れる
    # (`.梱包単位_重量 = …`)。以前は整数にしていて、外径 1160.5 が 1160、重量 1000.5 が 1000 になり、
    # パレットの選び方・積数・チェックリスト・発注履歴へそのまま乗っていた。整数の値は今までどおり
    return vba.fmt(value, decimals) if decimals is not None else vba.num_text(value)


def expand_order(conn: sqlite3.Connection, order_no: str) -> Optional[OrderInfo]:
    """受注番号から計算に要る値を展開する (VBA `フォーム展開`)。"""
    order_no = str(order_no or "").strip()
    if not order_no:
        return None
    row = db.fetch_one(
        conn, "SELECT * FROM 仕掛受注 WHERE 受注番号 = ?", (order_no,),
        caller_name="expand_order")
    if row is None:
        return None

    data = dict(row)
    info = OrderInfo(
        受注番号=str(data.get("受注番号", "")),
        受注材質=str(data.get("受注材質", "")),
        受注調質=str(data.get("受注調質", "")),
        受注板厚=vba.fmt(data.get("受注板厚"), 3),
        受注板幅=vba.fmt(data.get("受注板幅"), 1),
        用途コード=str(data.get("用途コード", "")),
        用途名=str(data.get("用途名", "")),
        包装仕様NO=str(data.get("包装仕様NO", ""))[:SPEC_NO_LENGTH],
        納入先名称=str(data.get("納入先名称", "")),
        取引先名称=str(data.get("取引先名称", "")),
        製品単重=vba.fmt(data.get("製品単重"), 2),
        梱包単位_重量=_or_no_data(vba.val(data.get("梱包単位_重量"))),
        梱包単位_枚数=_or_no_data(vba.val(data.get("梱包単位_枚数"))),
        コイル外径_MAX=_or_no_data(vba.val(data.get("コイル外径_MAX"))),
        コイル外径_目標=_or_no_data(vba.val(data.get("コイル外径_目標"))),
        コイル外径_MIN=_or_no_data(vba.val(data.get("コイル外径_MIN"))),
        コイル内径_目標=_or_no_data(vba.val(data.get("コイル内径_目標"))),
        工場用コメント=str(data.get("工場用コメント", "")),
        # 帳票の「営業納期」欄はこれ。VBA も `.出荷` へ入れ直している
        営業納期=str(data.get("営業納期", "")),
        比重=vba.int_text(vba.val(data.get("材質_比重")))
        if vba.val(data.get("材質_比重")) == int(vba.val(data.get("材質_比重")))
        else str(vba.val(data.get("材質_比重"))),
        梱包コード=str(data.get("梱包コード", "")),
    )
    info.出荷日 = info.営業納期
    return info


def load_into(conn: sqlite3.Connection, state: CalcState, order_no: str) -> bool:
    """受注情報を画面の状態へ入れる。外径の既定値もここで入れる。

    VBA は `ｺｲﾙ外径_MAX` が 0 でなければ外径欄へ入れていた。現場は
    そのまま使うことも、打ち直すこともある。
    """
    info = expand_order(conn, order_no)
    if info is None:
        return False
    state.order = info
    # コイル外径は**計算する**(外径 = BOX実績_ｺｲﾙ内径 + 2 × BOX実績_ｺｲﾙ肉厚)。受注の外径MAX があっても
    # 計算できれば計算した値を入れる(資材計算 0.2.35・現場の指示「データにあっても計算する。データが
    # 無ければ常に計算」)。先に梱包明細が読む仕掛(LS4LOT。**名前に関わらず列で読む**)、次に実績の形式の
    # ファイル(統合 1.2.29)。どちらでも出せなければ、いままでどおり受注の外径 MAX
    outer, where = _computed_outer(conn, state.LOT)
    if outer is not None:
        state.外径 = vba.fmt(outer, 2)
        state.note(f"コイル外径は{where}から計算しました(内径 + 2 × 肉厚 = {state.外径}mm)。"
                   "違うときは打ち直してください")
    elif info.コイル外径_MAX != OrderInfo.NO_DATA:
        state.外径 = info.コイル外径_MAX
        if where:
            state.note(f"コイル外径を計算できないので受注の外径MAXを入れました({where})。"
                       "違うときは打ち直してください")
        elif coil_results.wanted():
            # 実績を読む設定なのに、このロットが実績に無い。**黙って外径MAXにしない**(統合 1.2.49・現場の指摘)。
            # 実績は取り込んだ時点の写しなので、いつの写しかも添える
            state.note(f"実績に無いので受注の外径MAXを入れました(実績の取り込み: {_results_imported_at(conn)})。"
                       "違うときは打ち直してください")
    elif coil_results.wanted():
        state.note(f"実績に無く、受注の外径MAXも無いので、コイル外径を入れてください"
                   f"(実績の取り込み: {_results_imported_at(conn)})")
    # 検入数(本)も実績から(BOX実績_枚本数、無ければ 横割数 × 縦割数。統合 1.2.30)。
    # **先に手で入れてあれば上書きしない**
    count = coil_results.count_for(conn, state.LOT)
    if count is not None and not str(state.検入数).strip():
        state.検入数 = str(count)
        state.note(f"検入数は実績から入れました({count}本)。違うときは打ち直してください")
    return True


def _computed_outer(conn: sqlite3.Connection, lot: str) -> tuple:
    """(計算したコイル外径 または None, どこから/なぜ出せないか)。

    1. 梱包明細が読む仕掛(LS4LOT。`modules/packing_details/meisai/lot_lookup.py`)のそのロットの行の内径・肉厚
    2. 資材計算に取り込んだコイル実績(実績の形式のファイル。`coil_results.outer_for`)
    """
    why = ""
    try:
        from modules.packing_details.meisai import lot_lookup
        hit = lot_lookup.lookup(lot)
    except Exception as exc:                            # noqa: BLE001 - 引けなくても計算は止めない
        hit = {"found": False, "reason": f"仕掛を読めません: {exc}"}
    if hit.get("outer") is not None:
        return hit["outer"], "仕掛(LS4LOT)の内径・肉厚"
    if hit.get("found"):
        why = "仕掛(LS4LOT)の行に内径・肉厚がありません"
    outer = coil_results.outer_for(conn, lot)
    if outer is not None:
        return outer, "実績"
    return None, why


def _results_imported_at(conn: sqlite3.Connection) -> str:
    """コイル実績を取り込んだ日時(無ければ「まだ」)。"""
    try:
        row = conn.execute("SELECT 取り込み日時 FROM 取り込み履歴 WHERE テーブル名 = ?",
                           (coil_results.TABLE,)).fetchone()
    except sqlite3.Error:
        row = None
    return str(row[0]) if row is not None and row[0] else "まだ取り込んでいません"


# 受注の参照値の呼び名(取り込み直しで変わったことを知らせるとき)
_ORDER_LABELS = {
    "コイル外径_MAX": "外径MAX", "コイル外径_目標": "外径目標", "コイル外径_MIN": "外径MIN",
    "コイル内径_目標": "内径目標", "梱包単位_重量": "梱包単位_重量", "梱包単位_枚数": "梱包単位_枚数",
}


def refresh_order(conn: sqlite3.Connection, state: CalcState) -> tuple[bool, str]:
    """選んだ受注を、手元の仕掛受注から引き直す(計算Start・積数の手直しの前。統合 1.2.49・現場の指摘)。

    受注を選んだあとに仕掛受注が取り込み直されても、以前は作業状態に写した**前の値のまま**
    計算していた。引き直して違えば新しい値にし、何が変わったかを返す。

    **受注の参照値だけ**を新しくする。外径・検入数のように利用者が打ち直す欄は勝手に変えない。

    戻り値: `(受注がまだあるか, 変わった項目の説明)`。受注を選んでいなければ `(True, "")`。
    """
    order_no = str(state.order.受注番号 or "").strip()
    if not order_no:
        return True, ""
    info = expand_order(conn, order_no)
    if info is None:
        return False, ""
    before = state.order
    changes = []
    for name in OrderInfo.__dataclass_fields__:
        old, new = str(getattr(before, name, "")), str(getattr(info, name, ""))
        if old != new and name != "出荷日":             # 出荷日は営業納期の写し(同じことを2度言わない)
            changes.append(f"{_ORDER_LABELS.get(name, name)}: {old or '(空)'} → {new or '(空)'}")
    if not changes:
        return True, ""
    state.order = info
    log.info("受注 %s が取り込み直しで変わりました: %s", order_no, " / ".join(changes))
    return True, "受注の内容が取り込み直しで変わりました(" + "、".join(changes) + ")"
