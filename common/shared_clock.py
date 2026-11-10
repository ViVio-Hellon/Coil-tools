"""共有フォルダの時計(統合 1.2.72・構造点検の改善案3)

PC の時計は PC ごとにずれる。共有のファイルの「古さ」「期限」をこの PC の時計で測ると、時計の進んだ・遅れた
PC が、ほかの PC が使っている最中のものを消したり、期限を早く・遅く決めたりする(明細の履歴の整理 1.2.62、
共有の鍵 1.2.69、管理者以外停止の期限 1.2.67 で、同じ仕組みを別々に作っていた)。

**小さな試しのファイルを共有に書き、ファイルサーバーが付けた更新日時を見る。** 書き手の PC が日時を付ける
共有ではこの PC の時計と同じになり、ずれは見つけられない(そのときはこれまでと同じ)。
"""
from __future__ import annotations

import os
import time
import uuid
from pathlib import Path
from typing import Optional

from common.logging_utils import get_logger

log = get_logger("coil_packing_tools", "shared_clock")

PROBE_PREFIX = ".時計確認_"


def now(folder: Path) -> Optional[float]:
    """共有フォルダの「いま」(エポック秒)。書けない・確かめられなければ None。試しのファイルは残さない。"""
    probe = Path(folder) / f"{PROBE_PREFIX}{os.getpid()}_{uuid.uuid4().hex[:8]}.tmp"
    try:
        probe.write_bytes(b"")
        try:
            return probe.stat().st_mtime
        finally:
            try:
                probe.unlink()
            except OSError:
                pass
    except OSError as exc:
        log.warning("共有フォルダの時計を確かめられません: %s (%s)", exc, folder)
        return None


def offset(folder: Path) -> Optional[float]:
    """共有フォルダの時計 − この PC の時計(秒)。確かめられなければ None。"""
    before = time.time()
    shared = now(folder)
    after = time.time()
    if shared is None:
        return None
    return shared - (before + after) / 2
