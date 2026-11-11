"""15 分おきに、取り込み元(参照パス)と手元の写しを比べ、変わっていれば写し直す(統合 1.2.77・現場の指示)

【なぜ要るのか】
梱包明細と資材計算は、仕掛台帳を**手元の写し(作業用 DB)へ取り込んでから**引く。取り込むのは起動したときと
「取り込み」を押したときだけだったので、動かしている間に台帳へ載ったロットは、取り込み直すまで引けなかった
(梱包明細では「該当するﾛｯﾄ情報が見つかりません」。上の帯に「新しい台帳があります」と出るだけだった)。
ペナラベルのロット引きは、引くたびに共有のファイルを手元へ写し直してから読むので、いつも新しかった。

現場の指示: 3 つとも、参照パスから手元へ写して、写しを引く。15 分おきに手元と参照パスを比べ、更新されていれば
自動でまた手元へ写す。

【何をするか(15 分ごと)】
- 梱包明細: 取り込み元の大きさ・更新時刻を前に取り込んだときの目印と比べ(`data_sync.newer_sources`)、
  変わった表だけ取り込み直す。取り込み中(「取り込み」ボタン・起動時の自動取り込み)なら今回は見送る。
  開いているロットはそのまま(次にロットを入れたときから新しい台帳で引く。手で直した数を消さない)
- 資材計算: 取り込み元の目印を比べ(`data_sync.freshness`。小さいマスタはここで取り込み直す)、
  仕掛台帳が変わっていれば取り込み直す(`reimport_lot`)。取り込み中なら見送る
- ペナラベル: ロット引きが読む共有のファイル(梱包明細と同じ参照パス)の写しを取り直す
  (`lot_lookup.refresh_copies`。引くときにも変わっていれば写し直す)

**読めない・届かないときは何も入れ替えない**(各機能の取り込みの決まりのまま。1つ目の参照パスに届かないときは
前の中身のまま、など)。1 機能が失敗しても、ほかの機能と次の回は続ける。
"""
from __future__ import annotations

import logging
import threading
from typing import Callable, Optional

log = logging.getLogger("coil_packing_tools.auto_refresh")

#: 比べる間隔(秒)。現場の指示で 15 分
INTERVAL_SEC = 15 * 60

_thread: Optional[threading.Thread] = None
_stop = threading.Event()
_lock = threading.Lock()


def refresh_details() -> list[str]:
    """梱包明細: 変わった表だけ取り込み直す。取り込み直した表の名前を返す。"""
    from modules.packing_details.meisai import data_sync, db

    if data_sync.is_importing():
        log.info("梱包明細: 取り込み中なので、今回の見比べは見送ります")
        return []
    marks = data_sync.source_marks()
    with db.connect() as conn:
        newer = data_sync.newer_sources(conn, marks)
        if not newer:
            return []
        log.info("梱包明細: 参照パスの台帳が変わったので手元へ取り込み直します(%s)", "・".join(newer))
        result = data_sync.import_all(conn, only=newer)
    for note in result.errors:
        log.warning("梱包明細の自動の取り込み直し: %s", note)
    log.info("梱包明細の自動の取り込み直し: %s", result.summary())
    return list(result.imported)


def refresh_material() -> list[str]:
    """資材計算: マスタ(小さい表)と仕掛台帳を、変わっていれば取り込み直す。取り込み直した表の名前を返す。"""
    from modules.packing_material_calculation.coil_tool import data_sync, db

    if data_sync.is_importing():
        log.info("資材計算: 取り込み中なので、今回の見比べは見送ります")
        return []
    done: list[str] = []
    with db.connect() as conn:
        fresh = data_sync.freshness(conn)          # 小さいマスタはここで取り込み直す
        if fresh.importing:
            return []
        done += fresh.refreshed
        if fresh.lot_changed:
            log.info("資材計算: 参照パスの仕掛台帳が変わったので手元へ取り込み直します(%s)",
                     "・".join(fresh.check.lot))
            for result in data_sync.reimport_lot(conn):
                if result.ok:
                    done.append(result.table)
                else:
                    log.warning("資材計算の自動の取り込み直し: %s: %s", result.table, result.error)
    if done:
        log.info("資材計算の自動の取り込み直し: %s", "・".join(done))
    return done


def refresh_pena() -> int:
    """ペナラベル: ロット引きが読む共有のファイルの写しを取り直す。写したファイルの数を返す。"""
    from modules.packing_details.meisai import lot_lookup

    return lot_lookup.refresh_copies()


STEPS: tuple[tuple[str, Callable[[], object]], ...] = (
    ("梱包明細", refresh_details),
    ("資材計算", refresh_material),
    ("ペナラベル", refresh_pena),
)


def run_once() -> dict[str, object]:
    """3 機能ぶんを1回。1つが失敗してもほかは続ける。機能ごとの結果(失敗なら例外の文)を返す。"""
    out: dict[str, object] = {}
    with _lock:                                    # 前の回が長引いていたら重ねない
        for name, step in STEPS:
            try:
                out[name] = step()
            except Exception as exc:               # noqa: BLE001 - 見回りは止めない
                log.warning("%s: 参照パスとの見比べに失敗しました(次の回にまた見ます): %s", name, exc)
                out[name] = f"失敗: {exc}"
    return out


def start(interval: float = INTERVAL_SEC) -> threading.Thread:
    """見回りを始める(プロセスに1つ)。最初の回は interval 秒後(起動時の取り込みと重ねない)。"""
    global _thread
    if _thread is not None and _thread.is_alive():
        return _thread
    _stop.clear()

    def loop() -> None:
        log.info("参照パスと手元の写しの見比べを %d 分おきに行います", int(interval // 60))
        while not _stop.wait(interval):
            run_once()

    _thread = threading.Thread(target=loop, name="auto-refresh", daemon=True)
    _thread.start()
    return _thread


def stop() -> None:
    _stop.set()
