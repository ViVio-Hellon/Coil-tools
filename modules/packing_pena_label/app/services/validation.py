# -*- coding: utf-8 -*-
"""入力チェック（VBA の ValidateKensaNo / InpCheck / KeyPress 制限 相当）。

**VBA のメッセージ文言・判定順を一切変更していない。**

メッセージの「N桁目」は **右から数えた桁**（`Mid` は左からの位置なので
一見ずれて見えるが、7 桁の番号を右から読んだ表現で全部そろっている）。

実際の検査番号は **英字・数字・英数字・数字 4 桁**(例 N6X0945・H670129。3 文字目は英字のことも数字のこともある)。
VBA が 3 文字目を数字かどうか見ていなかったのは、そこに英字が入るから。
2026.09 の判断(解析書 B-2)で「3 文字目も数字」を足したため、正しい番号がすべて
「5桁目にアルファベット」で弾かれていた。統合 1.2.43 で戻した(現場の指摘)。
"""

from __future__ import annotations

from typing import List, Optional

from .size_master import take_of
from .vba_compat import is_numeric, to_narrow

#: ロット番号(LOTNO)の形(統合 1.2.63・現場の指示「7桁の決まりのある文字列が来たら自動で引く」)。
#: 現場から受け取った取り込み元の 8,301 ロット(SIKALOT・SIKAHIKI・LS4LOT・実績_LS4)と、現場が示した一覧で確かめた形:
#:   1 文字目 英字 / 2 文字目 英字か数字 / 3〜4 文字目 数字 / 5〜6 文字目 「数字+英字」か「英字+数字」 / 7 文字目 数字
#:   例 L7151N0・LX051L0・L815C51・HD034Y0・L0023M9・LY94Q50
#: 検査番号(N6X0945・H670129)はこの形に当たらない(3 文字目が英字、または 5〜6 文字目が両方数字)。
#: 画面はこの形がそろったら自動で検査番号を引く(`app.js`。この文字列を画面へ渡す)
LOT_PATTERN = r"^[A-Z][A-Z0-9][0-9]{2}(?:[0-9][A-Z]|[A-Z][0-9])[0-9]$"


def is_lot_no(text: str) -> bool:
    """ロット番号の形か(`LOT_PATTERN`)。"""
    import re
    return re.match(LOT_PATTERN, str(text or "")) is not None


#: VBA ValidateKensaNo の共通メッセージ前半（Chr(10) = LF）
_MSG_HEAD = "検査NOが7桁入力されていない\nもしくは"


def normalize_kensa_no(raw: str) -> str:
    """VBA ``StrConv(UCase(txtKensaNo), vbNarrow)`` と同じ正規化。"""
    return to_narrow((raw or "").upper())


def _mid(s: str, start: int, length: int) -> str:
    """VBA ``Mid(s, start, length)``（1 始まり、範囲外は空文字）。"""
    if start < 1:
        return ""
    return s[start - 1:start - 1 + length]


def _digits(s: str) -> bool:
    """数字 0〜9 だけでできているか(空は False)。

    VBA の ``IsNumeric`` は「1E11」(指数表記)・「+1」「1.1」「&H1」なども数とみなすので、
    検査番号の確かめには使わない(統合 1.2.36。「W111E11」が VBA でも Python でも通り、
    そのままラベルとバーコードに載っていた)。
    """
    return bool(s) and all("0" <= ch <= "9" for ch in s)


def _alnum(s: str) -> bool:
    """英大文字 A〜Z か数字 0〜9 だけでできているか(空は False)。"""
    return bool(s) and all("0" <= ch <= "9" or "A" <= ch <= "Z" for ch in s)


def validate_kensa_no(ken: str) -> Optional[str]:
    """VBA ``ValidateKensaNo``。OK なら None、NG ならメッセージを返す。

    判定順も VBA のまま（最初に該当したものだけを返す）。

    **3 文字目は英字でも数字でもよい**(VBA と同じく数字かどうかは見ない)。実際の検査番号は
    3 文字目が英字のもの(例 N6X0945・H5Z1404)も数字のもの(例 H670129)もある。2026.09 に足した「3 文字目も数字」の判定(解析書 B-2)は
    誤りで、正しい番号がすべて弾かれていた。統合 1.2.43 で、英字・数字以外(記号)だけを弾く形にした
    (判定は最後に置く。これまで弾いていた入力の文言は変わらない)。

    統合 1.2.36 で塞いだ穴(どれも VBA にもあった):

    - 数字の判定を ``IsNumeric`` から「0〜9 だけ」にした。「W111E11」(1E11 = 指数表記)・
      「W1+1111」などが通っていた。これまで弾かれていた入力の文言は変わらない
    - 8 文字以上を弾く。VBA の判定は 7 文字目までしか見ず、8 文字目からは何でも通った
      (VBA は欄の MaxLength=7 で入らないようにしていた。API から来る値もあるので、ここでも見る)
    - 1 文字目が英字か(記号が通っていた)
    """
    if not _digits(_mid(ken, 2, 1)):
        return _MSG_HEAD + "6桁目にアルファベットが入力されています"
    if not _digits(_mid(ken, 7, 1)):
        return _MSG_HEAD + "1桁目にアルファベットが入力されています"
    if not _digits(_mid(ken, 4, 4)):
        return _MSG_HEAD + "1～4桁目にアルファベットが入力されています"
    if is_numeric(_mid(ken, 1, 1)):
        return _MSG_HEAD + "7桁目に数字が入力されています"
    # 3 文字目(右から 5 桁目)は英字でも数字でもよい(例 N6X0945)。記号だけ弾く(統合 1.2.43)
    if not _alnum(_mid(ken, 3, 1)):
        return _MSG_HEAD + "5桁目に英字・数字以外が入力されています"
    # 統合 1.2.36 追加: 1 文字目は英字・全部で 7 文字
    if not ("A" <= _mid(ken, 1, 1) <= "Z"):
        return _MSG_HEAD + "7桁目にアルファベット以外が入力されています"
    if len(ken) != 7:
        return f"検査NOは7桁です({len(ken)}桁入力されています)"
    return None


def inp_check(state, mode: int) -> bool:
    """VBA ``InpCheck(a)``。重量が未入力なら True。

    mode 1 : 丈1（OptionButton 奇数 / lblSpec_08_53_1）
    mode 2 : 丈2（OptionButton 偶数 / lblSpec_08_53_2）
    mode 3 : 丈1,2 同時（CheckBox1..6 / ck_08_53）
    """
    w1 = state.weight1 or ""
    w2 = state.weight2 or ""

    # VBA: For i = 1 To 11 Step 2 / For i = 2 To 12 Step 2 と名前付き OB(13・14)を見ていた。
    # どれも「奇数=丈1・偶数=丈2」なので、丈の決まり(`take_of`)で見る ── 足したサイズ
    # (15〜)も同じに断る(統合 1.2.26。以前は番号を決め打ちして、足したサイズを見落とした)
    if mode == 1:
        if state.selected_ob and take_of(state.selected_ob) == 1:
            return w1 == ""
        return False

    if mode == 2:
        if state.selected_ob and take_of(state.selected_ob) == 2:
            return w2 == ""
        return False

    if mode == 3:
        if state.selected_cb:
            return w1 == "" or w2 == ""
        if state.named_cb:
            return w1 == "" or w2 == ""
        return False

    return False


def check_weight_inputs(state) -> Optional[str]:
    """VBA CommandButton1 の 3 連続 InpCheck。NG メッセージ or None。

    そのあと、使う丈の重量が **0 より大きい数** かを確かめる(統合 1.2.31・現場の指示)。
    欄は VBA と同じく「-」だけの入力を残す(打ち始めの途中)が、それで重量反映すると
    NW=0 のまま計算・印刷まで通った(VBA は計算の `CDbl("-")` で止まっていた)。
    """
    if inp_check(state, 1):
        return "重量入力がありません　丈1"
    if inp_check(state, 2):
        return "重量入力がありません　丈2"
    if inp_check(state, 3):
        return "重量入力がありません　丈1　or　丈2"
    for take, value in used_weights(state):
        if weight_number(value) is None:
            return (f"重量が数になっていません　丈{take}(「{value}」)。"
                    "0 より大きい数で入れてください")
    return None


def used_weights(state) -> List[tuple]:
    """いま選んでいるサイズで使う丈の (丈, 重量欄の値)。"""
    if state.selected_cb or state.named_cb:
        return [(1, state.weight1 or ""), (2, state.weight2 or "")]
    if state.selected_ob:
        take = take_of(state.selected_ob)
        return [(take, (state.weight1 if take == 1 else state.weight2) or "")]
    return []


def weight_number(value) -> Optional[float]:
    """重量欄・ラベルの重量の値を数に。「kg」は外す。0 以下・読めない(「-」だけ など)なら None。"""
    s = to_narrow(str(value or "")).replace("kg", "").replace(",", "").strip()
    if not s or not is_numeric(s):
        return None
    try:
        n = float(s)
    except ValueError:
        return None
    return n if n > 0 else None


def validate_settings(state) -> Optional[str]:
    """VBA ``IsValidSettings``。NG ならメッセージ、OK なら None。

    判定順は VBA のまま（チップボール → 丈フレーム表示 → サイズ選択）。

    VBA では 2 番目の「丈フレームが両方とも非表示」は **無言で中止** していた。
    これはフォームのコントロール ``Take1.Visible`` / ``Take2.Visible`` が
    サイズ選択とは独立したプロパティだったためで、
    「CheckBox のチェックを外したが Take フレームは表示のまま」という
    食い違った状態が作れたことによる。

    移植では丈フレームの表示を **選択状態から導出** している（models.FormState）。
    そのため「両方非表示」と「サイズ未選択」は常に同時に成立し、
    無言中止が起きる状況＝サイズ未選択 に一致する。
    利用者が理由の分からないまま放置されないよう、この場合は
    VBA の 3 番目のメッセージを返す。**計算を行わない点は VBA と同じ。**
    """
    from ..models import TIP_CHOICES

    if state.tip not in TIP_CHOICES:
        return "チップボールチェックがありません"
    if not (state.take1_visible or state.take2_visible):
        return "コイルサイズが選択されていません"
    if not state.any_size_selected():
        return "コイルサイズが選択されていません"
    return None


# ---------------------------------------------------------------- 入力制限
def filter_alnum_upper(raw: str) -> str:
    """VBA の KeyPress 制限（0-9, A-Z, a-z のみ / 小文字は大文字へ）。

    ``txtKensaNo`` / ``CoilH1..4`` / 全サイズの TextBox1..5 に適用。
    """
    out = []
    for ch in (raw or ""):
        code = ord(ch)
        if 48 <= code <= 57:        # 0-9
            out.append(ch)
        elif 65 <= code <= 90:      # A-Z
            out.append(ch)
        elif 97 <= code <= 122:     # a-z -> 大文字
            out.append(chr(code - 32))
        # それ以外は VBA と同じく破棄
    return "".join(out)


def filter_numeric_text(raw: str) -> str:
    """VBA の ``*_Change`` 制限。

    ``If Not IsNumeric(text) Then text = ""``（ただし "-" 単独は残す）。
    CoilH1..4 は "-" 許容の記述が無いため ``allow_minus=False`` で使う。
    """
    s = raw or ""
    if s == "-":
        return s
    return s if is_numeric(s) else ""


def filter_coil_text(raw: str) -> str:
    """``CoilH1..4_Change``: 数値でなければ空にする（"-" の例外なし）。"""
    s = raw or ""
    if not is_numeric(s):
        return ""
    # 本数は整数。11.5 は計算に 11.5 を使い、見出しは「12本」になっていた(統合 1.2.12)。
    # 違う数で計算するより、空にして入れ直してもらう
    # 「12.0」「１２」「012」はそのまま残すと、台紙・見出しに「12.0」と出た(統合 1.2.35)。整数の文字にそろえる
    try:
        n = float(to_narrow(s).replace(",", ""))
    except ValueError:
        return ""
    return str(int(n)) if n.is_integer() and n >= 0 else ""


def num_extract(value) -> str:
    """VBA ``numExtract``。数字と小数点だけを抽出。"/" があれば前半のみ。"""
    s = "" if value is None else str(value)
    if "/" in s:
        s = s.split("/")[0]
    return "".join(ch for ch in s if ch.isdigit() or ch == ".")


def convert_to_half(value) -> str:
    """VBA ``ConvertToHalf``。全角数字を半角にしてから数字だけ抽出。"""
    s = "" if value is None else str(value)
    out = []
    for ch in s:
        if "０" <= ch <= "９":
            out.append(chr(ord(ch) - 0xFEE0))
        else:
            out.append(ch)
    return num_extract("".join(out))
