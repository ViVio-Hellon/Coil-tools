"""仕掛台帳の1つ目のフォルダに届かないとき、2つ目で手元の表を入れ替えない(統合 1.2.69)

「読めない」を「無い」と取り違えない(統合 1.2.61 の決まり)。資材計算は、1つ目のフォルダに一瞬届かないと
「ファイルが無い」とみなして2つ目のファイルを読み、**手元の表を2つ目の中身だけに総入れ替え**していた
(2つ目は名前を変えた別のファイル・一部だけのファイルのことがある)。梱包明細は 1.2.61 で直していた
(`meisai/data_sync.held_back`)。同じ決まりにする:

    前は1つ目から取り込んでいた + いま1つ目のフォルダに届かない  → 入れ替えず、前の中身のまま理由を出す
    1つ目のフォルダには届くが、ファイルが無い                   → これまでどおり2つ目から読む
    前から2つ目を使っている(1つ目から取り込んだことが無い)       → これまでどおり2つ目から読む
"""
from __future__ import annotations

from pathlib import Path

from modules.packing_material_calculation.coil_tool import data_sync
from modules.packing_material_calculation.tests._web import client, sandbox   # noqa: F401
from modules.packing_material_calculation.tests.test_lot_second_dir import (_conn, _hiki, _odr,  # noqa: F401
                                                                           dirs, use_second)


def _lots(conn) -> list[str]:
    return sorted(r[0] for r in conn.execute('SELECT "ロット番号" FROM "仕掛引当"'))


def _result(results, table):
    return next(r for r in results if r.table == table)


def _unreachable(first: Path) -> Path:
    """1つ目のフォルダに届かなくする(共有の瞬断と同じ: フォルダごと見えない)。"""
    gone = first.with_name(first.name + "_見えない")
    first.rename(gone)
    return gone


def test_1つ目に届かないときは2つ目で入れ替えず前の中身のまま(dirs):
    first, second = dirs
    use_second(second, 仕掛引当="HIKI2")
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"), ("A000002", "O2"), ("A000003", "O3"))
    _hiki(second / "HIKI2.sqlite3", ("A000001", "O1"))
    _odr(first / "SIKAODR.sqlite3", ("O1", "1C0001"))
    conn = _conn()
    data_sync.import_lot(conn)
    assert _lots(conn) == ["A000001", "A000002", "A000003"]

    _unreachable(first)
    results = data_sync.import_lot(conn)
    assert _lots(conn) == ["A000001", "A000002", "A000003"], "2つ目の中身だけに入れ替えた"
    hiki = _result(results, "仕掛引当")
    assert not hiki.ok
    assert "届かない" in hiki.error and "前の中身のまま" in hiki.error


def test_届くようになったらまた1つ目から取り込む(dirs):
    first, second = dirs
    use_second(second, 仕掛引当="HIKI2")
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"), ("A000002", "O2"))
    _hiki(second / "HIKI2.sqlite3", ("A000009", "O9"))
    conn = _conn()
    data_sync.import_lot(conn)
    gone = _unreachable(first)
    data_sync.import_lot(conn)
    gone.rename(first)
    (first / "SIKAHIKI.sqlite3").unlink()
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"), ("A000002", "O2"), ("A000003", "O3"))
    results = data_sync.import_lot(conn)
    assert _result(results, "仕掛引当").ok
    assert _lots(conn) == ["A000001", "A000002", "A000003", "A000009"], "1つ目を正に、2つ目にだけあるものを足す"


def test_1つ目のフォルダに届いてファイルだけ無いなら2つ目から読む(dirs):
    first, second = dirs
    use_second(second, 仕掛引当="HIKI2")
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"), ("A000002", "O2"))
    _hiki(second / "HIKI2.sqlite3", ("A000009", "O9"))
    conn = _conn()
    data_sync.import_lot(conn)
    (first / "SIKAHIKI.sqlite3").unlink()                   # フォルダには届く。ファイルが無い
    results = data_sync.import_lot(conn)
    assert _result(results, "仕掛引当").ok
    assert _lots(conn) == ["A000009"]


def test_前から2つ目を使っている端末はこれまでどおり(dirs):
    first, second = dirs
    use_second(second, 仕掛引当="HIKI2")
    _hiki(second / "HIKI2.sqlite3", ("A000009", "O9"))
    conn = _conn()
    _unreachable(first)
    results = data_sync.import_lot(conn)
    assert _result(results, "仕掛引当").ok
    assert _lots(conn) == ["A000009"]


def test_起動時の自動取り込みでも入れ替えない(dirs):
    first, second = dirs
    use_second(second, 仕掛引当="HIKI2")
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"), ("A000002", "O2"), ("A000003", "O3"))
    _hiki(second / "HIKI2.sqlite3", ("A000001", "O1"))
    conn = _conn()
    data_sync.import_lot(conn)
    _unreachable(first)
    data_sync.auto_import(conn)
    assert _lots(conn) == ["A000001", "A000002", "A000003"], "起動時の自動取り込みが2つ目で入れ替えた"


def test_1表だけの取り込み直しでも入れ替えない(dirs):
    first, second = dirs
    use_second(second, 仕掛引当="HIKI2")
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"), ("A000002", "O2"))
    _hiki(second / "HIKI2.sqlite3", ("A000009", "O9"))
    conn = _conn()
    data_sync.import_lot(conn)
    _unreachable(first)
    result = data_sync.reimport_table(conn, "仕掛引当")
    assert not result.ok
    assert _lots(conn) == ["A000001", "A000002", "A000009"]
