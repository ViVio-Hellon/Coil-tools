"""試験で使う資材マスタ CSV(統合 1.2.31)

同梱の data/資材重量.csv は**検証用サンプル**で、業務では使わない(取得元へ届かないときの代わりに
使わない・使っても印刷を止める)。試験では、その写しを**設定で指定した CSV**として読ませる
(業務の取得元と同じ扱い)。サンプルそのものの扱いは test_sample_master.py で確かめる。
"""
from __future__ import annotations

import os
import shutil
import tempfile

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SAMPLE = os.path.join(ROOT, "data", "資材重量.csv")
_copy = ""


def business_csv() -> str:
    """同梱サンプルの写し(指定した CSV として使う)。1 回だけ作る。"""
    global _copy
    if not _copy or not os.path.exists(_copy):
        folder = tempfile.mkdtemp(prefix="ppl_material_")
        _copy = os.path.join(folder, "資材重量.csv")
        shutil.copyfile(SAMPLE, _copy)
    return _copy
