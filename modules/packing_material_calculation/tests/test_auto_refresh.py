"""資材計算も、15 分おきに参照パスと手元の写しを比べて写し直す(統合 1.2.77・現場の指示)

仕掛台帳は「大きいので取り込み直さず、帯で知らせる」だけだった(「いま取り込む」を押すまで前の写し)。
"""
from __future__ import annotations

import os
import time

from common import auto_refresh
from modules.packing_material_calculation.coil_tool import data_sync
from modules.packing_material_calculation.tests._web import client, sandbox   # noqa: F401
from modules.packing_material_calculation.tests.test_lot_second_dir import (_conn, _hiki, _odr,  # noqa: F401
                                                                           dirs)


def _lots(conn) -> list[str]:
    return sorted(r[0] for r in conn.execute('SELECT "ロット番号" FROM "仕掛引当"'))


def _replace(path, *pairs):
    tmp = path.with_name(path.name + ".new")
    _hiki(tmp, *pairs)
    os.replace(tmp, path)
    later = time.time() + 5                      # 目印(更新時刻)が確かに変わるように
    os.utime(path, (later, later))


def test_あとから台帳に載ったロットも見回りのあとは引ける(dirs):
    first, _ = dirs
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"))
    _odr(first / "SIKAODR.sqlite3", ("O1", "1C0001"))
    conn = _conn()
    data_sync.import_lot(conn)
    _replace(first / "SIKAHIKI.sqlite3", ("A000001", "O1"), ("A000002", "O2"))
    assert _lots(conn) == ["A000001"], "見回りの前は前の写しのまま"
    done = auto_refresh.refresh_material()
    assert "仕掛引当" in done
    assert _lots(conn) == ["A000001", "A000002"]


def test_変わっていなければ取り込み直さない(dirs):
    first, _ = dirs
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"))
    _odr(first / "SIKAODR.sqlite3", ("O1", "1C0001"))
    conn = _conn()
    data_sync.import_lot(conn)
    assert auto_refresh.refresh_material() == []


def test_取り込み中なら見送る(dirs, monkeypatch):
    first, _ = dirs
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"))
    conn = _conn()
    data_sync.import_lot(conn)
    _replace(first / "SIKAHIKI.sqlite3", ("A000001", "O1"), ("A000002", "O2"))
    monkeypatch.setattr(data_sync, "is_importing", lambda: True)
    assert auto_refresh.refresh_material() == []
    assert _lots(conn) == ["A000001"]
