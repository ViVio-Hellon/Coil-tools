# -*- coding: utf-8 -*-
"""SQLite のファイルが変わったかを見る目印(統合 1.2.69)

本体だけでなく **`-wal`・`-journal`** の大きさと更新日時も入れる。WAL の形のファイルでは、書き込みはまず
`-wal` に入り、本体は変わらない(ほかのプログラムが開いたままのあいだは本体へ移らない)。本体だけを見ると、
直したのに前の中身のまま動く(統合 1.2.49 で3機能の目印に入れたが、足したサイズ・ラベル台紙一覧が漏れていた)。
"""
from __future__ import annotations

import os
from typing import Optional, Tuple

SIDE_FILES = ("-wal", "-journal")


def stamp(path: str) -> Optional[Tuple]:
    """(場所, 本体の更新日時ns・大きさ, -wal・-journal の更新日時ns・大きさ)。本体を見られなければ None。"""
    try:
        st = os.stat(path)
    except OSError:
        return None
    out = [os.path.abspath(path), st.st_mtime_ns, st.st_size]
    for suffix in SIDE_FILES:
        try:
            side = os.stat(path + suffix)
            out.append((side.st_mtime_ns, side.st_size))
        except OSError:
            out.append(None)
    return tuple(out)
