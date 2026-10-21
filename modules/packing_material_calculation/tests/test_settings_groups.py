"""設定「取り込み」の枠(統合 1.2.30・現場の指摘)

    同じ枠で考えるもの(フォルダとそのファイル名)を、別々の欄で1つずつ保存していて間違えやすかった。
    梱包明細の「置き場所・取り込み」と同じく、**枠ごとに** 確かめる・保存して取り込み・既定に戻す。

    梱包資材マスタ / 第1パス 仕掛台帳 / 第1パス 梱包課共有(LS4LOT) / 第2パス(4つの名前)
"""
from __future__ import annotations

import sqlite3

from modules.packing_material_calculation.coil_tool import config, db, user_settings
from modules.packing_material_calculation.tests._web import client, sandbox   # noqa: F401
from modules.packing_material_calculation.tests.test_coil_results import _results, _ym
from modules.packing_material_calculation.tests.test_lot_second_dir import _hiki, _odr, dirs  # noqa: F401
from modules.packing_material_calculation.tests.test_web_settings import unlock


def _groups(c) -> dict:
    return {g["id"]: g for g in c.get("/api/settings").get_json()["view"]["groups"]}


def _count(table: str) -> int:
    conn = db.get_connection()
    try:
        return conn.execute(f'SELECT COUNT(*) FROM "{table}"').fetchone()[0]
    except sqlite3.Error:
        return -1
    finally:
        conn.close()


def test_枠は4つ_フォルダとファイル名が同じ枠に入る(client):
    groups = _groups(client)
    assert list(groups) == ["master", "lot", "konpo", "second"]
    keys = {g: [f["key"] for f in groups[g]["fields"]] for g in groups}
    assert keys["master"] == [config.KEY_MASTER_DB_DIR]
    assert keys["lot"] == [config.KEY_LOT_DB_DIR]
    assert keys["konpo"] == [config.KEY_KONPO_DB_DIR, config.KEY_KONPO_DB_FILE]
    assert keys["second"] == [config.KEY_LOT_DB_DIR_2, *config.KEY_LOT_DB2_FILES.values(),
                              config.KEY_LS4_DB2_FILE]
    # 探すファイルと、その枠で取り込む表を出す
    assert "SIKAHIKI" in groups["lot"]["files"]
    assert [t["table"] for t in groups["lot"]["tables"]] == ["仕掛引当", "仕掛受注"]
    # どの欄もどれか1つの枠に入っている(枠の外に取り残された欄が無い)
    every = sum(keys.values(), [])
    assert len(every) == len(set(every))


def test_パスワードが無ければ枠を保存できない(client, dirs):
    first, _second = dirs
    r = client.post("/api/settings/group", json={"group": "lot",
                                                 "values": {config.KEY_LOT_DB_DIR: str(first)}})
    assert r.status_code == 403
    assert user_settings.get(config.KEY_LOT_DB_DIR) in (None, "")
    r = client.post("/api/settings/group/reset", json={"group": "lot"})
    assert r.status_code == 403


def test_保存して取り込み_その枠の表だけ取り込む(client, dirs):
    first, _second = dirs
    _hiki(first / "SIKAHIKI.sqlite3", ("A000001", "10000001"))
    _odr(first / "SIKAODR.sqlite3", ("10000001", "1C0001"))
    unlock(client)
    r = client.post("/api/settings/group", json={"group": "lot",
                                                 "values": {config.KEY_LOT_DB_DIR: str(first)}})
    body = r.get_json()
    assert r.status_code == 200 and body["ok"], body
    assert config.lot_db_dir() == first
    assert {t["table"] for t in body["tables"]} == {"仕掛引当", "仕掛受注"}
    assert _count("仕掛引当") == 1 and _count("仕掛受注") == 1
    # 取り込んだ時刻が枠に出る
    lot = next(g for g in body["view"]["groups"] if g["id"] == "lot")
    assert all(t["at"] for t in lot["tables"]), lot["tables"]


def test_梱包課共有の枠はフォルダと名前をまとめて保存しコイル実績を取り込む(client, sandbox, tmp_path):
    konpo = tmp_path / "konpo"
    _results(konpo / "実績_LS4.sqlite3", [("A000001", f"{_ym()}-05 08:00:00", "508", "200", "18")])
    unlock(client)
    r = client.post("/api/settings/group", json={"group": "konpo", "values": {
        config.KEY_KONPO_DB_DIR: str(konpo), config.KEY_KONPO_DB_FILE: "実績_LS4"}})
    body = r.get_json()
    assert r.status_code == 200 and body["ok"], body
    assert [t["table"] for t in body["tables"]] == ["コイル実績"]
    assert _count("コイル実績") == 1
    assert config.konpo_db_file() == "実績_LS4"


def test_使えない名前は枠ごと保存しない(client, dirs, tmp_path):
    unlock(client)
    r = client.post("/api/settings/group", json={"group": "konpo", "values": {
        config.KEY_KONPO_DB_DIR: str(tmp_path), config.KEY_KONPO_DB_FILE: "a/b"}})
    assert r.status_code == 422
    assert r.get_json()["key"] == config.KEY_KONPO_DB_FILE
    # 同じ枠のフォルダも書いていない(半分だけ保存しない)
    assert user_settings.get(config.KEY_KONPO_DB_DIR) in (None, "")


def test_知らない枠は断る(client):
    unlock(client)
    assert client.post("/api/settings/group", json={"group": "x", "values": {}}).status_code == 400
    assert client.post("/api/settings/group/reset", json={"group": "x"}).status_code == 400


def test_既定に戻すは枠の欄をまとめて消す(client, dirs):
    first, second = dirs
    unlock(client)
    client.post("/api/settings/group", json={"group": "second", "values": {
        config.KEY_LOT_DB_DIR_2: str(second), config.KEY_LOT_DB2_FILES["仕掛受注"]: "ODR2"}})
    assert user_settings.get(config.KEY_LOT_DB2_FILES["仕掛受注"]) == "ODR2"
    r = client.post("/api/settings/group/reset", json={"group": "second"})
    assert r.status_code == 200
    assert user_settings.get(config.KEY_LOT_DB_DIR_2) in (None, "")
    assert user_settings.get(config.KEY_LOT_DB2_FILES["仕掛受注"]) in (None, "")


def test_確かめるは同じ枠で打ってあるフォルダを見る_保存前でも(client, dirs):
    """フォルダを打って保存する前に、名前を確かめられる(枠ごとに確かめる)。"""
    _first, second = dirs
    _odr(second / "ODR2.sqlite3", ("10000001", "1C0001"))
    r = client.post("/api/settings/probe", json={
        config.KEY_LOT_DB_DIR_2: str(second), config.KEY_LOT_DB2_FILES["仕掛受注"]: "ODR2"})
    results = {x["key"]: x for x in r.get_json()["results"]}
    assert results[config.KEY_LOT_DB2_FILES["仕掛受注"]]["ok"], results
    # 第2パスのフォルダは保存していない
    assert user_settings.get(config.KEY_LOT_DB_DIR_2) in (None, "")
