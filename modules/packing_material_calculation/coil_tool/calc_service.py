"""計算Start ── パレットから員数までを順に通す

VBA `UF_Material.CommandButton5_Click` の移植。**順番がすべて**です。

    ソフトクリア
      → パレット種類   (PalletType)    … 指定パレットフラグ・特殊フラグが立つ
      → パレット名称   (PalletSize)    … 指定パレットフラグを読む
      → 積数・台数     (台数計算)      … パレットサイズを高さの上限に使う
      → リプラ員数     (リプラ計算)    … 積数・台数を読む。最後

前の段が書いた値を次の段が読むので、入れ替えると結果が変わります。
`docs/VBA解析.md` §4 に、どの段が何を書くかの表があります。
"""
from __future__ import annotations

import sqlite3
from dataclasses import dataclass
from typing import Optional

from . import pallet_service, ripla_service, stack_service, vba
from .logging_utils import get_logger
from .models import FLAG_NO_SINGLE, FLAG_SINGLE_SPLIT, CalcState

log = get_logger("calc_service")


@dataclass
class CalcResult:
    """計算の結果。画面はこれを見て表示を組む。"""
    ok: bool = False
    reason: str = ""
    total_ripla: int = 0
    shape: ripla_service.ShapeView = None  # type: ignore[assignment]

    def __post_init__(self) -> None:
        if self.shape is None:
            self.shape = ripla_service.ShapeView()


def refuse_inspection_count(state: CalcState) -> Optional[CalcResult]:
    """検入数が空・1未満なら断る(計算Start・積数の手直しで共通)。

    VBA は 0 や負の検入数でもそのまま計算し、台数 0・-2 などをチェックリストへ積めた
    (統合 1.2.26・構造点検)。
    """
    if not str(state.検入数).strip():
        state.note("検入数を入力してください")
        return CalcResult(ok=False, reason="no_inspection_count")
    if state.検入数値 < 1:
        state.note("検入数には1以上を入れてください")
        return CalcResult(ok=False, reason="bad_inspection_count")
    return None


def calculate(conn: sqlite3.Connection, state: CalcState) -> CalcResult:
    """計算Start (VBA `CommandButton5_Click`)。

    検入数が無いと何も出せないので、そこだけ先に断る。
    """
    # 先に消す。断ったときに前の結果が残ると、そのままチェックリストへ積めてしまう(統合 1.2.26)
    state.soft_clear()
    refused = refuse_inspection_count(state)
    if refused:
        return refused

    pallet_service.pallet_type(conn, state)
    pallet_service.pallet_size(conn, state)
    if pallet_service.refuse_without_pallet(state):
        # ﾎｲｰﾙで収まるパレットが無い。VBA は負の台数を出すが、移植は断る
        # (不明点 Q24・回答済み)。理由の文言は pallet_size が入れている
        return CalcResult(ok=False, reason=pallet_service.REASON_NO_PALLET)

    stack = stack_service.calculate(conn, state)
    if not stack.ok:
        # パレットまでは出ている。そこまでを見せて、理由を添える
        return CalcResult(ok=False, reason=stack.reason)

    total = ripla_service.calculate(conn, state)

    shape = ripla_service.shape_view(
        state.パレット種類, int(vba.val(state.最下部本数)))
    return CalcResult(ok=True, total_ripla=total, shape=shape)


def _units_for_stack(conn: sqlite3.Connection, state: CalcState) -> int:
    """手で入れた積数から台数を出す。1本積み不可の品は、最後の1本の台を2台に分ける。"""
    ins, stack = state.検入数値, state.積数値
    spec = pallet_service.get_spec(conn, state.order.包装仕様NO) if state.order.包装仕様NO else None
    if spec is None or not stack_service.no_single_by_rule(state.order, spec):
        return vba.round_up(ins / stack)
    state.mark(FLAG_NO_SINGLE)
    count, _stack, equal = stack_service._split_for_no_single(state, stack, ins)
    if count > 0 and FLAG_SINGLE_SPLIT in state.flags:
        rest = int(ins - (0 if equal else count * stack))
        big, small = (rest + 1) // 2, rest // 2
        state.note("1本積み不可のため、最後の台が1本にならないように分けました("
                   + (f"{rest}本を {big}本・{small}本 の2台)" if equal else
                      f"{vba.int_text(stack)}本 × {count}台 + 別枠 {big}本・{small}本 の2台)"))
    return count


def recalculate_from_stack(conn: sqlite3.Connection, state: CalcState, *,
                           units: Optional[int] = None) -> CalcResult:
    """積数を手で直したときの再計算 (VBA `積数_AfterUpdate`)。

    パレットは引き直すが、**台数は積数から出し直す**(`台数計算` は
    呼ばない)。手で入れた積数を、規則で上書きしてしまわないため。

        台数 = 切り上げ(検入数 ÷ 積数)   (`units` を渡すとそれ。バリ揃えの結果転送)

    【1本積み不可の品(統合 1.2.27・現場の指示)】手で入れた積数でも、最後の台が1本だけに
    なるなら、計算Start と同じ振り分け(`stack_service._split_for_no_single`、VBA `台数計算`)
    で1台戻して2台に分ける。例: 検入数13・積数4 → 4・4・4・1 ではなく 4・4 + 別枠 3・2(4台)。

    【VBA のバグを直した(統合 1.2.26・現場で確認)】VBA は前の計算の別枠(Re_積数・Re_台数)・
    印・断りを消さずに計算し直していた。1本積み不可で分けたあとに積数を直すと、チェック
    リスト・HB・リプラで台数を二重に数えた(検入数13・枚数4 → 2台+別枠2台 のあと積数5 にすると、5・5・3 の3台のはずが5台)。
    手で入れた積数のほかは全部消してから出し直す。高さで直すはしない(手で入れた積数が
    最後の答え)。ただし1本積み不可の品は、1本だけの台を作らないように分ける(上)
    """
    state.clear_results(keep_stack=True)
    if not str(state.積数).strip():
        state.note("積数を入力してください")
        return CalcResult(ok=False, reason="no_stack")
    refused = refuse_inspection_count(state)
    if refused:
        return refused

    pallet_service.pallet_type(conn, state)
    pallet_service.pallet_size(conn, state)
    if pallet_service.refuse_without_pallet(state):
        return CalcResult(ok=False, reason=pallet_service.REASON_NO_PALLET)

    if state.積数値 <= 0:
        state.note("積数には1以上を入れてください")
        return CalcResult(ok=False, reason="zero_stack")

    if units is None:
        units = _units_for_stack(conn, state)
        if units <= 0:
            state.note("検入数が1本のため、1本積み不可の振り分けができません。積数を手で入れてください")
            state.台数 = ""
            return CalcResult(ok=False, reason="no_single_one_coil")
    state.台数 = vba.int_text(units)

    total = ripla_service.calculate(conn, state)
    shape = ripla_service.shape_view(
        state.パレット種類, int(vba.val(state.最下部本数)))
    return CalcResult(ok=True, total_ripla=total, shape=shape)
