"""仕掛台帳の2つ目(統合 1.2.6。python-web-tools VER4.4.0・VER4.8.15 と同じ決まり)

    1つ目のフォルダ … 決まった名前(SIKAHIKI / SIKAODR)で探す
    2つ目のフォルダ … 2つ目の名前(表ごとに設定。空なら1つ目と同じ名前)だけを探す

    ファイルが1つ目に無い      → そのファイルは2つ目から読む
    1つ目にロット・受注が無い  → 1つ目を取り込んだあと、2つ目にだけあるものを足す
    両方にある                 → 1つ目を正とする

2つ目のフォルダは以前の「仕掛台帳の予備フォルダ」(lot_db_dir_2)をそのまま使う。
"""
from __future__ import annotations

import os
import sqlite3
import time
from pathlib import Path

import pytest

from modules.packing_material_calculation.coil_tool import (config, data_sync, db, distribution,
                                                            import_specs, user_settings)
from modules.packing_material_calculation.tests._web import client, sandbox   # noqa: F401
from modules.packing_material_calculation.tests.test_web_settings import unlock


def _source(path: Path, table: str, rows: list[dict]) -> Path:
    """取り込み元(表「仕掛」)。要る列はすべて作る。"""
    path.parent.mkdir(parents=True, exist_ok=True)
    columns = [src for _, src, _ in import_specs.LOT_IMPORT_SPECS[table]]
    conn = sqlite3.connect(path)
    conn.execute('CREATE TABLE "仕掛" (%s)' % ", ".join(f'"{c}" TEXT' for c in columns))
    for row in rows:
        conn.execute('INSERT INTO "仕掛" (%s) VALUES (%s)' % (
            ", ".join(f'"{c}"' for c in columns), ", ".join("?" for _ in columns)),
            [row.get(c, "") for c in columns])
    conn.commit()
    conn.close()
    return path


def _odr(path: Path, *orders: tuple[str, str]) -> Path:
    return _source(path, "仕掛受注", [{"受注番号": o, "包装仕様NO": s} for o, s in orders])


def _hiki(path: Path, *pairs: tuple[str, str]) -> Path:
    return _source(path, "仕掛引当", [{"ﾛｯﾄ番号": lot, "受注番号": o} for lot, o in pairs])


@pytest.fixture()
def dirs(sandbox, tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DEFAULT_LOT_DB_DIR_FALLBACK", None)
    first, second = tmp_path / "lot", tmp_path / "lot2"
    first.mkdir(exist_ok=True)
    second.mkdir(exist_ok=True)
    return first, second


def use_second(second: Path, **names: str) -> None:
    user_settings.save(config.KEY_LOT_DB_DIR_2, str(second))
    for table, name in names.items():
        user_settings.save(config.KEY_LOT_DB2_FILES[table], name)


def _conn():
    conn = db.get_connection()
    db.apply_schema(conn)
    return conn


def test_1つ目に無いファイルは2つ目の名前で読む(dirs):
    first, second = dirs
    use_second(second, 仕掛受注="ODR2")
    _odr(second / "ODR2.sqlite3", ("O1", "1C0001"))
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"))
    assert data_sync.find_lot_db("仕掛受注") == second / "ODR2.sqlite3"
    assert data_sync.find_lot_db("仕掛引当") == first / "SIKAHIKI.sqlite3"


def test_名前を変えたら2つ目の同じ名前のファイルは読まない(dirs):
    first, second = dirs
    use_second(second, 仕掛受注="ODR2")
    _odr(second / "SIKAODR.sqlite3", ("X9", "9X9999"))
    assert data_sync.find_lot_db("仕掛受注") is None


def test_名前が空なら1つ目と同じ名前(dirs):
    first, second = dirs
    use_second(second)
    _odr(second / "SIKAODR.db", ("O1", "1C0001"))
    assert data_sync.find_lot_db("仕掛受注") == second / "SIKAODR.db"
    assert config.lot_db2_files()["仕掛引当"] == "SIKAHIKI.sqlite3"


def test_1つ目に無い受注ロットを2つ目から足す_両方にあれば1つ目(dirs):
    first, second = dirs
    use_second(second, 仕掛受注="ODR2", 仕掛引当="HIKI2")
    _odr(first / "SIKAODR.sqlite3", ("O1", "1C0001"))
    _odr(second / "ODR2.sqlite3", ("O1", "2X9999"), ("O2", "1C0002"))
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "O1"))
    _hiki(second / "HIKI2.sqlite3", ("A000001", "O9"), ("B000001", "O2"))
    conn = _conn()
    results = {r.table: r for r in data_sync.import_lot(conn)}
    assert all(r.ok for r in results.values()), [r.error for r in results.values()]
    specs = {r[0]: r[1] for r in conn.execute('SELECT "受注番号", "包装仕様NO" FROM "仕掛受注"')}
    assert specs == {"O1": "1C0001", "O2": "1C0002"}
    pairs = sorted(tuple(r) for r in conn.execute('SELECT "ロット番号", "受注番号" FROM "仕掛引当"'))
    assert pairs == [("A000001", "O1"), ("B000001", "O2")], "A000001 は1つ目の引当だけ"
    assert "2つ目" in results["仕掛受注"].note


def test_2つ目だけ変わっても自動取り込みで取り込み直す(dirs):
    first, second = dirs
    use_second(second, 仕掛受注="ODR2")
    _odr(first / "SIKAODR.sqlite3", ("O1", "1C0001"))
    _odr(second / "ODR2.sqlite3", ("O1", "1C0001"))
    conn = _conn()
    data_sync.auto_import(conn)
    (second / "ODR2.sqlite3").unlink()
    _odr(second / "ODR2.sqlite3", ("O1", "1C0001"), ("O3", "1C0003"))
    later = time.time() + 5
    os.utime(second / "ODR2.sqlite3", (later, later))
    data_sync.auto_import(conn)
    assert sorted(r[0] for r in conn.execute('SELECT "受注番号" FROM "仕掛受注"')) == ["O1", "O3"]


def test_列の足りない2つ目からは足さない(dirs):
    first, second = dirs
    use_second(second, 仕掛受注="ODR2")
    _odr(first / "SIKAODR.sqlite3", ("O1", "1C0001"))
    conn2 = sqlite3.connect(second / "ODR2.sqlite3")
    conn2.execute('CREATE TABLE "仕掛" ("受注番号" TEXT)')
    conn2.execute('INSERT INTO "仕掛" VALUES (\'O2\')')
    conn2.commit()
    conn2.close()
    conn = _conn()
    results = {r.table: r for r in data_sync.import_lot(conn)}
    assert [r[0] for r in conn.execute('SELECT "受注番号" FROM "仕掛受注"')] == ["O1"]
    assert "列がありません" in results["仕掛受注"].note


def test_設定画面_名前はパスワードが要る_使えない文字は断る_確かめる(client, dirs):
    first, second = dirs
    key = config.KEY_LOT_DB2_FILES["仕掛受注"]
    assert client.post("/api/settings", json={key: "ODR2"}).status_code == 403
    unlock(client)
    r = client.post("/api/settings", json={key: "sub/ODR2"})
    assert r.status_code == 422 and r.get_json()["reason"] == "bad_name"
    assert client.post("/api/settings", json={config.KEY_LOT_DB_DIR_2: str(second),
                                              key: "ODR2"}).status_code == 200
    # 確かめる: 見つからない → 足りない列 → そろっている
    probe = lambda: client.post("/api/settings/probe", json={key: "ODR2"}).get_json()["results"][0]
    assert not probe()["ok"] and "見つかりません" in probe()["message"]
    conn2 = sqlite3.connect(second / "ODR2.sqlite3")
    conn2.execute('CREATE TABLE "仕掛" ("受注番号" TEXT)')
    conn2.commit()
    conn2.close()
    assert "足りない列" in probe()["message"]
    (second / "ODR2.sqlite3").unlink()
    _odr(second / "ODR2.sqlite3", ("O1", "1C0001"))
    assert probe()["ok"]
    view = client.get("/api/settings").get_json()["view"]
    field = next(f for f in view["fields"] if f["key"] == key)
    assert field["value"] == "ODR2" and field["resolved"] == str(second / "ODR2.sqlite3")


def test_配布設定に入る():
    keys = {k for k, _, _ in distribution.ITEMS}
    assert set(config.KEY_LOT_DB2_FILES.values()) <= keys
    assert config.KEY_LOT_DB_DIR_2 in keys
