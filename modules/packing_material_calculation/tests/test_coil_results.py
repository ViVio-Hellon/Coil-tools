"""コイル外径を実績の形式のファイルから出す(統合 1.2.29・現場の指示。仕様は docs/仕掛ファイルの仕様.md)

    梱包課共有の仕掛のファイル名が LS4LOT 以外(例: 実績_LS4.sqlite3) … 実績の形式
        外径 = BOX実績_ｺｲﾙ内径 + 2 × BOX実績_ｺｲﾙ肉厚。BOX実績_作業開始_日付 がいまの年月の行だけ
    名前が LS4LOT のまま / 外径が出せない … いままでどおり受注の外径 MAX
"""
from __future__ import annotations

import sqlite3
from datetime import date
from pathlib import Path

import pytest

from modules.packing_material_calculation.coil_tool import (coil_results, config, data_sync, db,
                                                            lot_service, user_settings)
from modules.packing_material_calculation.coil_tool.models import CalcState, OrderInfo
from modules.packing_material_calculation.tests._web import client, sandbox   # noqa: F401
from modules.packing_material_calculation.tests.test_web_settings import unlock

COLUMNS = ("ﾛｯﾄ番号", "BOX設計_横割数", "BOX設計_縦割数", "BOX実績_作業開始_日付",
           "BOX実績_ｺｲﾙ内径", "BOX実績_ｺｲﾙ肉厚")


def _ym(months_ago: int = 0) -> str:
    today = date.today()
    y, m = today.year, today.month - months_ago
    while m < 1:
        y, m = y - 1, m + 12
    return f"{y}-{m:02d}"


def _results(path: Path, rows: list[tuple]) -> Path:
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.execute('CREATE TABLE "実績" (%s)' % ", ".join(f'"{c}" TEXT' for c in COLUMNS))
    for lot, started, inner, thick in rows:
        conn.execute('INSERT INTO "実績" VALUES (?, ?, ?, ?, ?, ?)', (lot, "12", "2", started, inner, thick))
    conn.commit()
    conn.close()
    return path


@pytest.fixture()
def place(sandbox, tmp_path, monkeypatch):
    monkeypatch.setattr(config, "DEFAULT_LOT_DB_DIR_FALLBACK", None)
    konpo = tmp_path / "konpo"
    konpo.mkdir()
    user_settings.save(config.KEY_KONPO_DB_DIR, str(konpo))
    conn = db.get_connection()
    db.apply_schema(conn)
    return konpo, conn


def test_名前がLS4LOTのままなら読まない(place):
    konpo, conn = place
    _results(konpo / "LS4LOT.sqlite3", [("A000001", f"{_ym()}-01 08:00:00", "508", "200")])
    assert coil_results.wanted() is False
    assert coil_results.import_results(conn) is None
    assert coil_results.outer_for(conn, "A000001") is None


def test_実績の形式なら外径は内径と2倍の肉厚_いまの年月の行だけ(place):
    konpo, conn = place
    user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
    _results(konpo / "実績_LS4.sqlite3", [
        ("A000001", f"{_ym()}-21 00:01:56", "508", "200.5"),
        ("B000001", f"{_ym(1)}-28 08:00:00", "508", "100"),     # 先月は使わない
        ("C000001", f"{_ym()}-03 08:00:00", "", "100"),          # 内径が空 → 外径は出さない
    ])
    result = coil_results.import_results(conn)
    assert result.ok and result.rows == 2, result
    assert coil_results.outer_for(conn, "A000001") == 909.0
    assert coil_results.outer_for(conn, "B000001") is None
    assert coil_results.outer_for(conn, "C000001") is None


def test_第2パスは第1パスに無いロットだけ足す(place, tmp_path):
    konpo, conn = place
    second = tmp_path / "lot2"
    user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
    user_settings.save(config.KEY_LOT_DB_DIR_2, str(second))
    _results(konpo / "実績_LS4.sqlite3", [("A000001", f"{_ym()}-05 08:00:00", "508", "200")])
    _results(second / "実績_LS4.sqlite3", [("A000001", f"{_ym()}-05 08:00:00", "600", "300"),
                                         ("D000001", f"{_ym()}-05 08:00:00", "500", "150")])
    coil_results.import_results(conn)
    assert coil_results.outer_for(conn, "A000001") == 908.0        # 第1パスの値
    assert coil_results.outer_for(conn, "D000001") == 800.0


def test_受注を選ぶと実績の外径を入れる_無ければ外径MAX(place, monkeypatch):
    konpo, conn = place
    user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
    _results(konpo / "実績_LS4.sqlite3", [("A000001", f"{_ym()}-05 08:00:00", "508", "200.5")])
    coil_results.import_results(conn)
    monkeypatch.setattr(lot_service, "expand_order",
                        lambda c, o: OrderInfo(受注番号=o, コイル外径_MAX="1000.00"))
    state = CalcState(LOT="A000001")
    assert lot_service.load_into(conn, state, "O1")
    assert state.外径 == "909.00"
    assert any("実績から" in m for m in state.messages)
    other = CalcState(LOT="Z999999")
    assert lot_service.load_into(conn, other, "O1")
    assert other.外径 == "1000.00"


def test_取り込みの結果に出る_実績の列が無ければ断る(place):
    konpo, conn = place
    user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
    bad = konpo / "実績_LS4.sqlite3"
    c = sqlite3.connect(bad)
    c.execute('CREATE TABLE "仕掛" ("ﾛｯﾄ番号" TEXT)')
    c.commit()
    c.close()
    result = data_sync.import_all(conn)
    coil = [t for t in result.tables if t.table == coil_results.TABLE][0]
    assert not coil.ok and "BOX設計_横割数" in coil.error
    assert "BOX実績_作業開始_日付" in coil_results.missing_columns(bad)


def test_設定画面に欄があり_確かめるで形式を言う(client, place):
    konpo, _conn = place
    _results(konpo / "実績_LS4.sqlite3", [("A000001", f"{_ym()}-05 08:00:00", "508", "200")])
    keys = {f["key"] for f in client.get("/api/settings").get_json()["view"]["fields"]}
    assert {config.KEY_KONPO_DB_DIR, config.KEY_KONPO_DB_FILE, config.KEY_LS4_DB2_FILE} <= keys
    unlock(client)
    r = client.post("/api/settings/probe", json={config.KEY_KONPO_DB_FILE: "実績_LS4"})
    assert r.get_json()["results"][0]["ok"], r.get_json()
    r = client.post("/api/settings/probe", json={config.KEY_KONPO_DB_FILE: ""})
    assert "LS4LOT" in r.get_json()["results"][0]["message"]
