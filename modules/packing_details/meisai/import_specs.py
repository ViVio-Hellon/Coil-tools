"""取り込み元の列 → 手元の列の対応表

移植元(`python-web-tools` の `import_specs.py`)と同じ形にしてある。
1行が `(手元の列名, 取り込み元の列名, 変換関数)`。

【なぜ名前を付け替えるのか】
取り込み元の列名は**半角カナ**(`ﾛｯﾄ番号` `ｵｰﾀﾞｰ板丈` `用途ｺｰﾄﾞ`)。
VBAのSQLはこれをそのまま書いていたが、手元では全角にそろえる。
半角カナは環境によって表示・入力が揺れるうえ、SQLを書くたびに
IMEを切り替えることになるため。対応表をここ1か所に置けば、
上流が名前を変えてもここだけ直せばよい。

【変換関数が効く場面】
取り込み元は**全列 TEXT**(LS4LOT に至っては型宣言すら無い)。
そのまま写すと、数値として比較できない列が残る。

    ｵｰﾀﾞｰ板丈   '0.0'  →  to_real  →  0.0   … `= 0` で引ける
    当工程設計_縦割数 '2'  →  to_int  →  2     … `= 2` で引ける

**この変換を挟まないとVBAの判定条件が成立しない。** 取り込み元を
直接引くと `ｵｰﾀﾞｰ板丈 = '0'` は1件も一致しない(実値が `'0.0'` のため)。
"""
from __future__ import annotations

import re
from typing import Any, Optional

from .logging_utils import get_logger

log = get_logger("import_specs")


# ------------------------------------------------------------------
# 変換関数
# ------------------------------------------------------------------
def _text(value: Any) -> str:
    return "" if value is None else str(value).strip()


def to_text(value: Any) -> str:
    return _text(value)


def to_int(value: Any) -> Optional[int]:
    """整数にする。読めなければ None(呼び出し側が既定値を入れる)。

    `int(float(...))` を通すのは、取り込み元が `'2'` だけでなく
    `'2.0'` の形でも来るため。`int('2.0')` は例外になる。
    """
    text = _text(value)
    if text == "":
        return None
    try:
        return int(float(text))
    except ValueError:
        return None


# VBA `Val()` が読む形: 符号・数字・小数点・指数(E/D)。空白はどこにあっても飛ばす
_VAL_RE = re.compile(r"[+-]?(?:\d+\.?\d*|\.\d+)(?:[eEdD][+-]?\d+)?", re.ASCII)


def vba_val(value: Any) -> float:
    """VBA `Val()` 相当(統合 1.2.12)。**先頭から数値として読める分だけ**を読む。

    `Val("0.8mm")` は 0.8、`Val("1,234")` は 1、`Val("abc")` は 0、`Val("1E-05")` は 0.00001。
    全角数字は読めない(`Val("１２")` は 0)。以前は取り込みで `float()` を使っていて、
    "0.8mm" を 0(読めない)にし、"１２" を 12 にしていた(VBA と違う数)。
    """
    text = re.sub(r"[ \t\r\n\u3000]", "", _text(value))
    m = _VAL_RE.match(text)          # re.ASCII: 全角数字は数字として読まない
    if not m:
        return 0.0
    num = m.group(0).replace("d", "e").replace("D", "e")
    try:
        return float(num)
    except ValueError:
        return 0.0


def to_real(value: Any) -> Optional[float]:
    """実数にする(VBA `Val()` と同じ読み方)。空なら None。"""
    text = _text(value)
    if text == "":
        return None
    # 数として読める所が1文字も無い("abc")は None(既定の 0.0 が入る。VBA の Val も 0)
    if not _VAL_RE.match(re.sub(r"[ \t\r\n\u3000]", "", text)):
        return None
    return vba_val(text)


def to_whole(value: Any) -> Optional[int]:
    """割数(縦割数・横割数)。**整数でなければ取り込まない**(None → 0 で入る。統合 1.2.12)。

    以前は `int(float())` で切り捨てていた(2.7 → 2)。VBA の読み方(Long へ入れると 2.7 → 3)が
    分からないうえ、どちらにしても前工程実績数・丈数が変わるので、違う数を入れるより 0 にして
    画面で前工程実績数を入れてもらう。ログにも残す。
    """
    text = _text(value)
    if text == "":
        return None
    number = vba_val(text)
    if number != int(number):
        log.warning("割数が整数ではないため 0 で取り込みます(前工程実績数は画面で入れる): %r", text)
        return None
    return int(number)


# 取り込み元のテーブル名。4ファイルとも「仕掛」で共通
SOURCE_TABLE = "仕掛"

# ------------------------------------------------------------------
# 対応表
# ------------------------------------------------------------------
# **VBAが読む列だけを取り込む。** 仕掛ロットは196列あるが、
# `mDB.GetLotInfo` が使うのは10列と抽出条件の1列だけ。全部写しても
# 使わない列が増えるだけで、上流の列が増減したときの当たり判定が
# 広がる。必要になったらここへ1行足す
IMPORT_SPECS: dict[str, list[tuple[str, str, object]]] = {
    "仕掛ロット": [
        ("ロット番号", "ﾛｯﾄ番号", to_text),
        ("用途コード", "用途ｺｰﾄﾞ", to_text),
        ("用途名", "用途名", to_text),
        ("製造材質", "製造材質", to_text),
        ("製造調質", "製造調質", to_text),
        ("製造板厚", "製造板厚", to_real),
        ("製造板幅", "製造板幅", to_real),
        ("オーダー板厚", "ｵｰﾀﾞｰ板厚", to_real),
        ("オーダー板幅", "ｵｰﾀﾞｰ板幅", to_real),
        # 抽出条件。元は TEXT の '0.0' なので必ず to_real を通す
        ("オーダー板丈", "ｵｰﾀﾞｰ板丈", to_real),
        ("設計_設備コース", "設計_設備ｺｰｽ", to_text),
        ("実績_設備コース", "実績_設備ｺｰｽ", to_text),
    ],
    "仕掛引当": [
        ("ロット番号", "ﾛｯﾄ番号", to_text),
        ("受注番号", "受注番号", to_text),
        # 番号であって数値ではない。数値にすると指数表記になり読めなくなる
        ("引当番号", "引当番号", to_text),
    ],
    "仕掛受注": [
        ("受注番号", "受注番号", to_text),
        ("包装仕様NO", "包装仕様NO", to_text),
        ("取引先名称", "取引先名称", to_text),
        ("納入先名称", "納入先名称", to_text),
    ],
    "仕掛当工程": [
        ("ロット番号", "ﾛｯﾄ番号", to_text),
        ("当工程設計_設備名", "当工程設計_設備名", to_text),
        # 元は型宣言なしの text("1".."6")。to_int を通して数値にする
        ("当工程設計_縦割数", "当工程設計_縦割数", to_whole),
        ("当工程設計_横割数", "当工程設計_横割数", to_whole),
        ("当工程設計_枚本数", "当工程設計_枚本数", to_int),
        ("製造板厚", "製造板厚", to_real),
        ("製造板幅", "製造板幅", to_real),
        ("製造板丈", "製造板丈", to_real),
    ],
}

# 値が `None` の行を落とす列。**鍵が空の行は入れても引けない。**
REQUIRED_KEY_COLUMNS: dict[str, tuple[str, ...]] = {
    "仕掛ロット": ("ロット番号",),
    "仕掛引当": ("ロット番号",),
    "仕掛受注": ("受注番号",),
    "仕掛当工程": ("ロット番号",),
}

# 空文字も「無し」として扱う列。`REQUIRED_KEY_COLUMNS` は `None` だけを
# 撥ねるので、`to_text` が返す空文字は通ってしまう
BLANK_IS_MISSING: dict[str, tuple[str, ...]] = {
    "仕掛ロット": ("ロット番号",),
    "仕掛引当": ("ロット番号",),
    "仕掛受注": ("受注番号",),
    "仕掛当工程": ("ロット番号",),
}

# 変換結果が None のときに NOT NULL 列へ入れる既定値。
# SQLite の DEFAULT は「列を省略したとき」にしか効かず、NULL を明示的に
# INSERT すると NOT NULL 制約で落ちる
NULL_FALLBACKS: dict[object, object] = {to_int: 0, to_whole: 0, to_real: 0.0, to_text: ""}
