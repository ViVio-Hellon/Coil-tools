"""管理者パスワード ── 梱包明細の共有のパスワード1つで確かめる(統合 1.2.75・構造点検の改善案4)

【何のためのものか】
共有フォルダのマスタと取り込み元の置き場所を**誤って書き換えない**ための関門です
(VBA版 `MaterialMasterForm` の `ADMIN_PASSWORD` から引き継いだ役目)。
本来の意味でのアクセス制御ではありません。

【全ラインで1つ】
以前はこの PC の設定ファイルに持つ**PCごとのパスワード**でした。梱包明細(全ラインで共有)・
ペナラベル(決まった値)と合わせて3系統あり、梱包明細で変えても資材計算は古いままでした。
いまは**梱包明細の共有のパスワード**(`modules.packing_details.meisai.admin_password`)で確かめます。
この画面から変えても、梱包明細の設定から変えても、全ラインの3機能に効きます。

この PC の設定ファイルに前の `admin_password` の行が残っていても、もう読みません
(切り替えの時点でどの PC も変えていなかった。初期値は3系統とも同じ)。

このモジュールは、資材計算の中の呼び出し口(`verify` / `change` / `reset` / `is_custom`)を
そのまま残し、中身を梱包明細へ渡すだけにしてある。
"""
from __future__ import annotations

from modules.packing_details.meisai import admin_password as _shared

from .logging_utils import get_logger

log = get_logger("admin_password")

# 以前この PC の設定に持っていた鍵(もう読まない)。配布設定の古い束に残っていても取り込まない
KEY = "admin_password"

MIN_LENGTH = _shared.MIN_LENGTH

# 断りの種類。**文言から推し量らない**(設計.md §1)
REFUSE_WRONG = _shared.REFUSE_WRONG
REFUSE_TOO_SHORT = _shared.REFUSE_TOO_SHORT
REFUSE_MISMATCH = _shared.REFUSE_MISMATCH
REFUSE_SAME = _shared.REFUSE_SAME
REFUSE_SHARED = _shared.REFUSE_SHARED

Result = _shared.Result


def is_custom() -> bool:
    """全ラインで変えてあるか。設定画面に出す(値そのものは出さない)。"""
    return _shared.is_custom()


def verify(password: str) -> bool:
    """合っているか。照合は梱包明細の1か所でしかしない(控えが読めなければ通さない)。"""
    return _shared.verify(str(password or ""))


def change(current: str, new: str, confirm: str) -> Result:
    """変える(全ラインで共有)。いまのパスワードを知っている人だけ。共有に届かなければ変えない。"""
    return _shared.change(current, new, confirm)


def reset(current: str) -> Result:
    """既定に戻す(全ラインで共有)。"""
    return _shared.reset(current)
