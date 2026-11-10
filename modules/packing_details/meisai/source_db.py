"""取り込み元(sqlite3)の読み書き。**中身は `common/source_db.py` に1つ**(統合 1.2.71・構造点検の改善案2)

梱包明細と資材計算に、ほぼ同じ写しが1本ずつあり、直しが片方にしか入らないことが続いた。このモジュールは
`common.source_db` **そのもの**を指す(`from . import source_db` で読んでも同じもの。試験の差し替えも1つに効く)。
"""
import sys

from common import source_db as _impl

sys.modules[__name__] = _impl
