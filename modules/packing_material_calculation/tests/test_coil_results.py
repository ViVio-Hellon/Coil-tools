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
           "BOX実績_ｺｲﾙ内径", "BOX実績_ｺｲﾙ肉厚", "BOX実績_枚本数")


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
    for row in rows:
        lot, started, inner, thick = row[:4]
        count = row[4] if len(row) > 4 else ""
        yoko, tate = (row[5], row[6]) if len(row) > 6 else ("12", "2")
        conn.execute('INSERT INTO "実績" VALUES (?, ?, ?, ?, ?, ?, ?)',
                     (lot, yoko, tate, started, inner, thick, count))
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


def test_第1パスで外径_検入数が空なら第2パスで空の項目だけ埋める(place, tmp_path, monkeypatch):
    """統合 1.2.45・現場の指摘: 第1パスが読めても検入数が埋まらないことがあった。"""
    konpo, conn = place
    second = tmp_path / "lot2"
    user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
    user_settings.save(config.KEY_LOT_DB_DIR_2, str(second))
    now = f"{_ym()}-05 08:00:00"
    _results(konpo / "実績_LS4.sqlite3", [
        ("A000001", now, "508", "200", "", "", ""),     # 検入数が出せない(枚本数・横割縦割が空)
        ("B000001", now, "", "", "18"),                  # 外径が出せない(内径・肉厚が空)
        ("C000001", now, "", "", "", "", ""),            # どちらも出せない
        ("E000001", now, "", "", "", "", ""),            # 第2パスにも無い
    ])
    _results(second / "実績_LS4.sqlite3", [
        ("A000001", now, "600", "300", "16"),
        ("B000001", now, "500", "150", "99"),
        ("C000001", now, "400", "100", "", "5", "4"),
    ])
    result = coil_results.import_results(conn)
    assert result.ok, result
    assert "外径・検入数が空だった 3ロットを第2パスで埋めました" in result.note
    assert coil_results.outer_for(conn, "A000001") == 908.0, "値のある外径は第1パスのまま"
    assert coil_results.count_for(conn, "A000001") == 16, "空の検入数は第2パスで"
    assert coil_results.outer_for(conn, "B000001") == 800.0, "空の外径は第2パスで"
    assert coil_results.count_for(conn, "B000001") == 18, "値のある検入数は第1パスのまま"
    assert (coil_results.outer_for(conn, "C000001"), coil_results.count_for(conn, "C000001")) == (600.0, 20)
    assert (coil_results.outer_for(conn, "E000001"), coil_results.count_for(conn, "E000001")) == (None, None)
    src = conn.execute('SELECT 元ファイル FROM "コイル実績" WHERE ロット番号 = ?', ("A000001",)).fetchone()[0]
    assert "検入数は" in src and str(second) in src
    # 画面: 受注を選ぶと第2パスで埋めた検入数が入る
    monkeypatch.setattr(lot_service, "expand_order",
                        lambda c, o: OrderInfo(受注番号=o, コイル外径_MAX="1000.00"))
    state = CalcState(LOT="A000001")
    assert lot_service.load_into(conn, state, "O1")
    assert (state.外径, state.検入数) == ("908.00", "16")


def test_第1パスが読めなくても第2パスで埋める(place, tmp_path):
    konpo, conn = place
    second = tmp_path / "lot2"
    user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
    user_settings.save(config.KEY_LOT_DB_DIR_2, str(second))
    c = sqlite3.connect(konpo / "実績_LS4.sqlite3")
    c.execute('CREATE TABLE "仕掛" ("ﾛｯﾄ番号" TEXT)')      # 実績の列が無い
    c.commit()
    c.close()
    _results(second / "実績_LS4.sqlite3", [("A000001", f"{_ym()}-05 08:00:00", "600", "300", "16")])
    result = coil_results.import_results(conn)
    assert result.ok, result
    assert "読めません" in result.note
    assert coil_results.count_for(conn, "A000001") == 16


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
    assert any("実績から計算しました" in m for m in state.messages)
    other = CalcState(LOT="Z999999")
    assert lot_service.load_into(conn, other, "O1")
    assert other.外径 == "1000.00"


def test_外径MAXがあっても仕掛の内径と肉厚から計算する_名前がLS4LOTでも(place, monkeypatch, tmp_path):
    """資材計算 0.2.35・現場の指示「コイル外径はデータにあっても計算する(データが無ければ常に計算)」。
    梱包明細が読む仕掛(LS4LOT)の同じロットの行に内径・肉厚があれば、名前が LS4LOT のままでも計算する。"""
    from modules.packing_details.meisai import lot_lookup
    ls4 = tmp_path / "share" / "LS4LOT.sqlite3"
    ls4.parent.mkdir()
    c = sqlite3.connect(ls4)
    c.execute('CREATE TABLE "仕掛" ("ﾛｯﾄ番号" TEXT, "検査番号" TEXT, "BOX実績_ｺｲﾙ内径" REAL, "BOX実績_ｺｲﾙ肉厚" REAL)')
    c.execute('INSERT INTO "仕掛" VALUES (?, ?, ?, ?)', ("A000002", "N6X0945", 508, 251.5))
    c.execute('INSERT INTO "仕掛" VALUES (?, ?, ?, ?)', ("A000003", "N6X0946", None, None))
    c.commit()
    c.close()
    lot_lookup.forget()
    monkeypatch.setattr(lot_lookup, "sources", lambda table="仕掛当工程": [ls4] if table == "仕掛当工程" else [])
    _konpo, conn = place
    assert coil_results.wanted() is False                  # 資材計算の設定は LS4LOT のまま
    monkeypatch.setattr(lot_service, "expand_order",
                        lambda c, o: OrderInfo(受注番号=o, コイル外径_MAX="1000.00"))
    state = CalcState(LOT="A000002")
    assert lot_service.load_into(conn, state, "O1")
    assert state.外径 == "1011.00"                         # 508 + 2 × 251.5。外径MAX 1000 ではない
    assert any("仕掛(LS4LOT)の内径・肉厚から計算しました" in m for m in state.messages)
    other = CalcState(LOT="A000003")                       # 内径・肉厚が空 → 外径MAX と理由
    assert lot_service.load_into(conn, other, "O1")
    assert other.外径 == "1000.00"
    assert any("計算できないので受注の外径MAX" in m for m in other.messages)


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


def test_検入数は枚本数_無ければ横割数と縦割数の積_手で入れてあれば上書きしない(place, monkeypatch):
    konpo, conn = place
    user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4")
    _results(konpo / "実績_LS4.sqlite3", [
        ("A000001", f"{_ym()}-05 08:00:00", "508", "200", "18"),
        ("B000001", f"{_ym()}-05 08:00:00", "508", "200", ""),      # 枚本数が空 → 12 × 2
    ])
    coil_results.import_results(conn)
    assert coil_results.count_for(conn, "A000001") == 18
    assert coil_results.count_for(conn, "B000001") == 24
    monkeypatch.setattr(lot_service, "expand_order",
                        lambda c, o: OrderInfo(受注番号=o, コイル外径_MAX="1000.00"))
    state = CalcState(LOT="A000001")
    lot_service.load_into(conn, state, "O1")
    assert state.検入数 == "18"
    assert any("検入数は実績から" in m for m in state.messages)
    typed = CalcState(LOT="A000001", 検入数="7")
    lot_service.load_into(conn, typed, "O1")
    assert typed.検入数 == "7"


def test_前の版の手元の表にも枚本数の列を足す(place):
    _konpo, conn = place
    conn.execute('DROP TABLE "コイル実績"')
    conn.execute('CREATE TABLE "コイル実績" (ロット番号 TEXT PRIMARY KEY, 内径 REAL, 肉厚 REAL, 外径 REAL,'
                 " 作業開始 TEXT NOT NULL DEFAULT '', 元ファイル TEXT NOT NULL DEFAULT '')")
    db.apply_schema(conn)
    cols = {r[1] for r in conn.execute('PRAGMA table_info("コイル実績")')}
    assert "枚本数" in cols
