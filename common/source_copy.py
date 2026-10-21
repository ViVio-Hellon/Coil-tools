"""取り込み元(共有フォルダの sqlite3)は、**読む前に必ず手元へ写す**(統合 1.2.28・現場の指示)

梱包明細と資材計算の `source_db` が、読み取り専用で開くときはいつもここを通る。
以前は共有の上のファイルをそのまま `mode=ro` で開き、開けないときだけ写していた。

【なぜ写してから読むのか】
- 共有の上で開いている間、上流(仕掛台帳を作る仕組み・梱包資材マスタを差し替える人)を
  待たせない。写すのは一瞬で、あとは手元のファイルを読む
- 共有(SMB)の上の sqlite3 は、WAL・ロック・URI の書き方しだいで開けないことがある。
  手元のディスクなら同じ開き方で必ず開ける
- 読んでいる途中で上流が書き換えても、読み始めた時点の中身でそろう

【写し方】
- 本体と `-wal` / `-journal` を一緒に写す(本体だけだと、直前に書かれた分が抜ける)。
  `-shm` は写さない(WAL の索引。手元で作り直される)
- 写す前と後で大きさ・更新時刻を比べ、**写している間に書き換わったら写し直す**
- 中身が前と同じ(大きさ・更新時刻が同じ)なら写し直さない
- 写しは毎回別の名前にする。前の写しを誰かが開いていても(Windows は消せない)止まらない
- 置き場所はこの PC の一時フォルダ。アプリが終わると消す

書くとき(マスタ管理の書き戻し)は写しへ向けない ── 書いたものがどこにも残らなくなる。
"""
from __future__ import annotations

import atexit
import itertools
import logging
import os
import shutil
import tempfile
import threading
import time
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

log = logging.getLogger("coil_packing_tools.source_copy")

#: 本体と一緒に写すもの。`-shm` は WAL の索引で、手元で作り直されるので写さない
SIDECARS = ("-wal", "-journal")
#: 写している間に書き換わったときに写し直す回数
RETRIES = 3
#: 写し直す前に待つ秒
RETRY_WAIT_SEC = 0.5


class CopyError(OSError):
    """手元へ写せなかった(届かない・写している間ずっと書き換わっていた など)。"""


@dataclass
class _Known:
    stamp: tuple
    path: Path


_lock = threading.Lock()
_dir: Optional[Path] = None
_known: dict[str, _Known] = {}
_counter = itertools.count(1)


def _stamp(src: Path) -> tuple:
    """中身が変わったかを見る目印(本体と付属ファイルの大きさ・更新時刻)。"""
    parts = []
    for extra in ("",) + SIDECARS:
        target = Path(str(src) + extra)
        try:
            st = target.stat()
        except FileNotFoundError:
            parts.append(None)
            continue
        parts.append((st.st_size, st.st_mtime_ns))
    if parts[0] is None:
        raise CopyError(f"ファイルが見つかりません: {src}")
    return tuple(parts)


def copy_dir() -> Path:
    """写しの置き場所(この PC の一時フォルダの下。アプリが終わると消す)。"""
    global _dir
    if _dir is None or not _dir.exists():
        _dir = Path(tempfile.mkdtemp(prefix="cpt_src_"))
        atexit.register(shutil.rmtree, _dir, True)
    return _dir


def _remove(path: Path) -> None:
    for extra in ("",) + SIDECARS + ("-shm",):
        try:
            Path(str(path) + extra).unlink()
        except OSError:
            pass                                   # 開いている写しは消せない(Windows)。終わるときに消える


def local_copy(path: Path) -> Path:
    """`path` を手元へ写した実体の場所。中身が前と同じなら前の写しを返す。

    写している間に書き換わったら写し直す。`RETRIES` 回とも書き換わっていたら `CopyError`。
    """
    src = Path(os.path.abspath(path))
    key = os.path.normcase(str(src))
    with _lock:
        last_reason = ""
        for attempt in range(1, RETRIES + 1):
            before = _stamp(src)
            known = _known.get(key)
            if known is not None and known.stamp == before and known.path.exists():
                return known.path
            dest = copy_dir() / f"{abs(hash(key)):x}_{next(_counter)}{src.suffix}"
            try:
                for extra in SIDECARS:
                    side = Path(str(src) + extra)
                    if side.exists():
                        shutil.copyfile(side, Path(str(dest) + extra))
                shutil.copyfile(src, dest)
            except OSError as exc:
                _remove(dest)
                raise CopyError(f"{src.name} を手元へ写せませんでした: {exc}") from exc
            after = _stamp(src)
            if after == before:
                if known is not None:
                    _remove(known.path)
                _known[key] = _Known(after, dest)
                log.info("取り込み元を手元へ写しました: %s → %s", src, dest)
                return dest
            _remove(dest)
            last_reason = "写している間に書き換わりました"
            log.info("%s: %s(%d回目)。写し直します", src.name, last_reason, attempt)
            time.sleep(RETRY_WAIT_SEC)
        raise CopyError(f"{src.name} を手元へ写せませんでした: {last_reason}"
                        f"(上流が書き込み中かもしれません。少し待ってからもう一度)")


def forget(path: Optional[Path] = None) -> None:
    """覚えている写しを忘れる。次に読むときは必ず写し直す。

    `path` を渡せばそのファイルだけ(書き戻したあと。大きさも更新時刻も変わらない書き方でも
    古い写しを読まないため)。省けば全部(試験用)。
    """
    with _lock:
        if path is None:
            _known.clear()
            return
        _known.pop(os.path.normcase(os.path.abspath(str(path))), None)
