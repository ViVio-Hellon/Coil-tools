"""マスタを直したのに、前の中身のまま動く・表示する ことが無いか(統合 1.2.49・現場の指摘)

現場の依頼:「マスタを直しても前の中身のまま動く・表示する事がないかのチェックを
してください 例: 資材重量 変更後 前の重量のまま計算」。点検で見つかったものを、
1つずつ確かめる。

    1  起動時の自動取り込みは「更新日時が前より新しいか」でなく、取り込み元の目印が**一致するか**で見る
    2  動かしている間も、帯・計算Start で取り込み元を確かめる(マスタは取り込み直す・仕掛台帳は知らせる)
    3  受注を選んだあとに仕掛受注が取り込み直されたら、計算Start で新しい受注の値にする
    4  計算のあとにマスタ・台帳が変わったら、古い計算結果をチェックリストへ積ませない
    5  発注票を作ったあとにチェックリストが変わったら、古い票を印刷・記録させない
    6  コイル実績も確かめる・実績に無いロットは黙って外径MAXにしない
    7  「古い」は取り込み元が変わったかで出す(写した日時と確かめた日時を分ける)
    8  プログラムの規則が優先される品番には、マスタ管理で札を出す
    9  「いま取り込む」・マスタ管理の取り込み直しは、覚えている写しを捨ててから読む
    10 取り込み中は帯に出し、計算Start を断る
"""
from __future__ import annotations

import os
import shutil
import sqlite3
import time
from datetime import date
from pathlib import Path

import pytest

from common import source_copy
from modules.packing_material_calculation.coil_tool import (config, data_sync, db, import_specs,
                                                            master_admin, special_rules, user_settings)
from modules.packing_material_calculation.tests._web import client, sandbox, seed   # noqa: F401


# ==================================================================
# 共有フォルダに見立てた取り込み元
# ==================================================================
def _write_table(path: Path, table: str, columns: list[str], rows: list[list], *, wal=False) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    if wal:
        conn.execute("PRAGMA journal_mode=WAL")
    conn.execute(f'DROP TABLE IF EXISTS "{table}"')
    conn.execute(f'CREATE TABLE "{table}" (%s)' % ", ".join(f'"{c}"' for c in columns))
    conn.executemany(f'INSERT INTO "{table}" VALUES (%s)' % ", ".join("?" for _ in columns), rows)
    conn.commit()
    conn.close()


def share_from_local(sandbox: Path, *, wal=False) -> tuple[Path, Path, Path]:
    """手元DBの写し(`seed` で入れたもの)を、そのまま共有の取り込み元として書き出す。

    梱包資材マスタ(4表)と仕掛台帳(SIKAHIKI / SIKAODR)。列名は取り込み元の名前にする。
    """
    master = sandbox / "master" / config.MATERIAL_DB_NAME
    with db.connect() as conn:
        for table, specs in import_specs.MASTER_IMPORT_SPECS.items():
            dest = [d for d, _s, _c in specs]
            rows = [[r[d] for d in dest] for r in conn.execute(
                "SELECT %s FROM %s" % (", ".join(f'"{d}"' for d in dest), f'"{table}"'))]
            _write_table(master, table, [s for _d, s, _c in specs], rows, wal=wal)
        paths = {}
        for table, specs in import_specs.LOT_IMPORT_SPECS.items():
            dest = [d for d, _s, _c in specs]
            rows = [[r[d] for d in dest] for r in conn.execute(
                "SELECT %s FROM %s" % (", ".join(f'"{d}"' for d in dest), f'"{table}"'))]
            path = sandbox / "lot" / config.LOT_DB_FILES[table]
            _write_table(path, import_specs.LOT_SOURCE_TABLE, [s for _d, s, _c in specs], rows)
            paths[table] = path
    return master, paths["仕掛引当"], paths["仕掛受注"]


def nudge(path: Path) -> None:
    """書いたあと、更新時刻を確実に動かす(時計の刻みが粗い環境でも、試験の意図どおりにする)。"""
    st = path.stat()
    os.utime(path, ns=(st.st_atime_ns, st.st_mtime_ns + 7_000_000_123))


def age(path: Path, seconds: int = 3600) -> None:
    """更新時刻を過去にする(取り込んだ日時より前。以前の「新しいか」の決まりでは飛ばされる形)。"""
    when = time.time() - seconds
    os.utime(path, (when, when))


def set_on_share(path: Path, sql: str, params=()) -> None:
    """共有の取り込み元を、ほかの端末が直したことにする。"""
    conn = sqlite3.connect(path)
    conn.execute(sql, params)
    conn.commit()
    conn.close()
    nudge(path)


def local(sql: str, params=()):
    with db.connect() as conn:
        return [tuple(r) for r in conn.execute(sql, params)]


def pallet_w() -> list[float]:
    return sorted(r[0] for r in local("SELECT W FROM パレット WHERE 種類 = 'スカシ'"))


def auto() -> data_sync.SyncResult:
    with db.connect() as conn:
        return data_sync.auto_import(conn)


@pytest.fixture
def shared(sandbox):
    """共有の取り込み元があり、起動時の自動取り込みで手元へ写したところ。"""
    source_copy.forget()
    seed(sandbox)
    master, hiki, odr = share_from_local(sandbox)
    result = auto()
    assert result.ok, [t.error for t in result.tables if not t.ok]
    source_copy.forget()            # 起動し直したことにする(写しの覚えは消える)
    return master, hiki, odr


@pytest.fixture
def ready(client, shared):
    client.post("/api/settings", json={"worker": config.WORKERS[0], "line": "LS4"})
    return client


def calc(c, n="20"):
    c.post("/api/calc/lot", json={"LOT": "A123456"})
    c.post("/api/calc/order", json={"受注番号": "12345678"})
    c.post("/api/calc/input", json={"検入数": n})
    return c.post("/api/calc/run", json={})


# ==================================================================
# 1. 起動時の自動取り込みは、目印が一致するかで見る
# ==================================================================
def test_1_WALで_wal_にだけ書かれた変更も起動時に取り込む(sandbox):
    """A: WAL の取り込み元は、書いても本体が変わらず `-wal` だけが変わる。"""
    source_copy.forget()
    seed(sandbox)
    master, _h, _o = share_from_local(sandbox, wal=True)
    age(master)
    auto()
    assert 900 in pallet_w()
    holder = sqlite3.connect(master)              # ほかの端末が開いたまま(チェックポイントされない)
    holder.execute("SELECT count(*) FROM パレット").fetchone()
    before = master.stat().st_mtime_ns
    w = sqlite3.connect(master)
    w.execute("PRAGMA wal_autocheckpoint=0")
    w.execute("UPDATE パレット SET Ｗ = 1250 WHERE Ｗ = 900")
    w.commit()
    w.close()
    try:
        assert master.stat().st_mtime_ns == before, "本体は変わらず -wal だけが変わる場面"
        source_copy.forget()
        auto()
        assert 1250 in pallet_w() and 900 not in pallet_w()
    finally:
        holder.close()


def test_1_更新日時を保ったままの差し替えも取り込む(sandbox):
    """B: エクスプローラのコピーは更新日時を保つ。以前は「取り込み日時より古い」ので飛ばした。"""
    source_copy.forget()
    seed(sandbox)
    master, _h, _o = share_from_local(sandbox)
    edited = sandbox / "edited.sqlite3"
    shutil.copyfile(master, edited)
    set_on_share(edited, "UPDATE パレット SET Ｗ = 1250 WHERE Ｗ = 900")
    old = time.time() - 3600                          # 資材課が1時間前に手元で直したもの
    os.utime(edited, (old, old))
    auto()
    shutil.copy2(edited, master)                      # 共有へ置く(更新日時は1時間前のまま)
    source_copy.forget()
    auto()
    assert 1250 in pallet_w()


def test_1_古いバックアップに戻しても取り込む(sandbox):
    """C: 戻したファイルは更新日時が過去。大小でなく一致で見るので拾える。"""
    source_copy.forget()
    seed(sandbox)
    master, _h, _o = share_from_local(sandbox)
    backup = sandbox / "bak.sqlite3"
    shutil.copy2(master, backup)
    os.utime(backup, (time.time() - 86400, time.time() - 86400))
    set_on_share(master, "UPDATE パレット SET Ｗ = 1250 WHERE Ｗ = 900")   # 誤った値で取り込む
    auto()
    assert 1250 in pallet_w()
    shutil.copy2(backup, master)
    source_copy.forget()
    auto()
    assert 900 in pallet_w() and 1250 not in pallet_w()


def test_1_写したあと記録する前に共有が変わっても次で取り込む(sandbox, monkeypatch):
    """F: 目印は読む**前**に取ったものを残す。読んだあとに取ると、新しい目印に古い中身を結びつける。"""
    source_copy.forget()
    seed(sandbox)
    master, _h, _o = share_from_local(sandbox)
    age(master, 3600)
    original = data_sync.source_db.read_table

    def read_then_changed(path, table):
        rows = original(path, table)
        if table == "パレット" and Path(path).name == master.name:
            set_on_share(master, "UPDATE パレット SET Ｗ = 1250 WHERE Ｗ = 900")
            age(master, 1800)          # 時計のずれ・粗い刻みで、取り込み日時より前に見える形
        return rows

    monkeypatch.setattr(data_sync.source_db, "read_table", read_then_changed)
    auto()
    assert 900 in pallet_w()                          # 読んだのは変わる前
    monkeypatch.setattr(data_sync.source_db, "read_table", original)
    source_copy.forget()
    auto()
    assert 1250 in pallet_w()


def test_1_目印が同じなら起動時は読み直さない(shared, monkeypatch):
    """起動が遅くならないように、変わっていない表は飛ばす(いままでどおり)。"""
    calls = []
    monkeypatch.setattr(data_sync.source_db, "read_table",
                        lambda *a, **k: calls.append(a) or [])
    result = auto()
    assert result.ok and calls == []


def test_1_前の版の取り込み履歴にも目印の列を足す(sandbox):
    path = config.DB_PATH
    path.parent.mkdir(parents=True, exist_ok=True)
    conn = sqlite3.connect(path)
    conn.execute("CREATE TABLE 取り込み履歴 (テーブル名 TEXT PRIMARY KEY, 取り込み日時 TEXT NOT NULL DEFAULT '',"
                 " 元ファイル TEXT NOT NULL DEFAULT '', 件数 INTEGER NOT NULL DEFAULT 0)")
    conn.execute("INSERT INTO 取り込み履歴 VALUES ('パレット', '2026-01-01 00:00:00', 'x', 1)")
    conn.commit()
    conn.close()
    with db.connect() as c:
        db.apply_schema(c)
        have = {r[1] for r in c.execute('PRAGMA table_info("取り込み履歴")')}
        assert {"元の目印", "確かめた日時"} <= have
        assert data_sync.last_import(c)["パレット"]["件数"] == 1


# ==================================================================
# 2. 動かしている間も確かめる
# ==================================================================
def test_2_帯の問い合わせで共有のマスタの変更を取り込み直す(client, shared):
    master, _h, _o = shared
    set_on_share(master, "UPDATE パレット SET Ｗ = 1250 WHERE Ｗ = 900")
    body = client.get("/api/ribbon").get_json()
    assert 1250 in pallet_w()
    assert "取り込み直しました" in body["ribbon"]["refreshed"]
    assert body["ribbon"]["stale"] is False


def test_2_画面を開いたときも取り込み直す(client, shared):
    master, _h, _o = shared
    set_on_share(master, "UPDATE パレット SET Ｗ = 1250 WHERE Ｗ = 900")
    page = client.get("/calc?t=test-token").data.decode()
    assert 1250 in pallet_w()
    assert "取り込み直しました" in page


def test_2_計算Startでも確かめて新しいマスタで計算する(ready, shared):
    master, _h, _o = shared
    first = calc(ready).get_json()["view"]
    assert first["result"]["パレットサイズ"]
    # マスタ管理(ほかの端末)で、外径900 が収まるパレットを 1250 にした
    set_on_share(master, "UPDATE パレット SET Ｗ = 1250, 新記号 = 'P125' WHERE Ｗ = 900")
    again = ready.post("/api/calc/run", json={}).get_json()
    assert 1250 in pallet_w()
    assert any("取り込み直しました" in m for m in again["view"]["messages"])


def test_2_仕掛台帳が変わると帯に出て_いま取り込むで取り込む(client, shared, monkeypatch):
    _m, _h, odr = shared
    set_on_share(odr, "UPDATE 仕掛 SET 梱包単位_枚数 = 3 WHERE 受注番号 = '12345678'")
    rb = client.get("/api/ribbon").get_json()["ribbon"]
    assert rb["lot_changed"] is True
    assert "共有の仕掛台帳が更新されています" in rb["lot_note"]
    # 大きいので自動では取り込まない
    assert local("SELECT 梱包単位_枚数 FROM 仕掛受注 WHERE 受注番号 = '12345678'") == [(5.0,)]
    page = client.get("/calc?t=test-token").data.decode()
    assert "いま取り込む" in page and "共有の仕掛台帳が更新されています" in page

    forgotten = []
    real = source_copy.forget
    monkeypatch.setattr(source_copy, "forget", lambda p=None: (forgotten.append(p), real(p)))
    r = client.post("/api/ribbon/import-lot", json={}).get_json()
    assert r["ok"] is True
    assert r["ribbon"]["lot_changed"] is False
    assert local("SELECT 梱包単位_枚数 FROM 仕掛受注 WHERE 受注番号 = '12345678'") == [(3.0,)]
    assert odr in [Path(p) for p in forgotten if p is not None], "写しを捨ててから読む(9)"


def test_2_マスタ管理の画面に手元へ取り込んだことを出す(client, shared):
    master, _h, _o = shared
    set_on_share(master, "UPDATE パレット SET Ｗ = 1250 WHERE Ｗ = 900")
    view = client.get("/api/master/browse?table=パレット").get_json()
    assert "取り込み直しました" in view["sync_note"]
    assert 1250 in pallet_w()


def test_2_マスタ管理で取り込み直せないときは一言出す(client, shared, monkeypatch):
    master, _h, _o = shared
    set_on_share(master, "UPDATE パレット SET Ｗ = 1250 WHERE Ｗ = 900")
    monkeypatch.setattr(data_sync, "_import_table",
                        lambda conn, **k: data_sync.TableResult(table=k["dest_table"], ok=False,
                                                                error="読めません"))
    view = client.get("/api/master/browse?table=パレット").get_json()
    assert "手元へまだ取り込んでいない変更があります" in view["sync_note"]


# ==================================================================
# 3. 受注の取り込み直しを計算Start で拾う
# ==================================================================
def test_3_受注が取り込み直されたら計算Startは新しい受注の値で計算する(ready):
    calc(ready)
    ready.post("/api/calc/input", json={"外径": "880"})       # 利用者が打ち直した外径
    with db.connect() as conn:                                  # 取り込み直しで 梱包単位_枚数 5 → 3
        conn.execute("UPDATE 仕掛受注 SET 梱包単位_枚数 = 3, コイル外径_MAX = 950"
                     " WHERE 受注番号 = '12345678'")
        conn.commit()
    body = ready.post("/api/calc/run", json={}).get_json()
    view = body["view"]
    assert view["order"]["梱包単位_枚数"] == "3"
    assert view["order"]["コイル外径_MAX"] == "950"
    assert view["input"]["外径"] == "880.00", "利用者が打ち直す欄は勝手に変えない"
    joined = " ".join(view["messages"])
    assert "受注の内容が取り込み直しで変わりました" in joined
    assert "梱包単位_枚数: 5 → 3" in joined


def test_3_積数の手直しでも受注を引き直す(ready):
    calc(ready)
    with db.connect() as conn:
        conn.execute("UPDATE 仕掛受注 SET 製品単重 = 120 WHERE 受注番号 = '12345678'")
        conn.commit()
    ready.post("/api/calc/input", json={"積数": "2"})
    view = ready.post("/api/calc/restack", json={}).get_json()["view"]
    assert view["order"]["製品単重"] == "120.00"
    assert any("取り込み直しで変わりました" in m for m in view["messages"])


def test_3_受注が無くなっていたら計算Startを断る(ready):
    calc(ready)
    with db.connect() as conn:
        conn.execute("DELETE FROM 仕掛受注 WHERE 受注番号 = '12345678'")
        conn.commit()
    r = ready.post("/api/calc/run", json={})
    assert r.status_code == 422 and r.get_json()["reason"] == "order_not_found"


# ==================================================================
# 4. 計算のあとにマスタ・台帳が変わったら積ませない
# ==================================================================
def test_4_計算のあとにマスタ管理で直したら古い結果を積めない(ready, shared):
    master, _h, _o = shared
    assert calc(ready).status_code == 200
    user_settings_ok = master_admin.admin_session.open_with(config.ADMIN_PASSWORD)
    assert user_settings_ok
    rowid = sqlite3.connect(master).execute(
        "SELECT rowid FROM パレット WHERE Ｗ = 1000 AND 種類 = 'スカシ'").fetchone()[0]
    with db.connect() as conn:
        assert master_admin.save_row(conn, "パレット", rowid, {"新記号": "P10b"}).ok
    r = ready.post("/api/checklist/add", json={})
    assert r.status_code == 409
    assert r.get_json()["reason"] == "master_changed"
    assert "計算Start" in r.get_json()["message"]
    # 計算し直せば積める
    assert ready.post("/api/calc/run", json={}).status_code == 200
    assert ready.post("/api/checklist/add", json={}).status_code == 200


def test_4_計算のあとに取り込んだら古い結果を積めない(ready):
    assert calc(ready).status_code == 200
    ready.post("/api/settings/import", json={})
    r = ready.post("/api/checklist/add", json={})
    assert r.status_code == 409 and r.get_json()["reason"] == "master_changed"


# ==================================================================
# 5. 発注票を作ったあとにチェックリストが変わったら、古い票を出さない
# ==================================================================
def _build(c):
    calc(c)
    assert c.post("/api/checklist/add", json={}).status_code == 200
    b = c.post("/api/order/build", json={"確認済み": True}).get_json()
    assert b["ok"] and b["sheets"]


def test_5_チェックリストを消したら古い発注票を記録も印刷もできない(ready):
    _build(ready)
    ready.post("/api/checklist/clear", json={})
    assert ready.get("/api/order").get_json()["sheets"] == []
    c = ready.post("/api/order/commit", json={})
    assert c.status_code in (409, 422) and not c.get_json()["ok"]
    assert ready.get("/report/order?t=test-token").status_code in (404, 409)
    with db.connect() as conn:
        assert conn.execute("SELECT count(*) FROM 発注履歴").fetchone()[0] == 0


def test_5_チェックリストに積み足したら古い発注票を記録できない(ready):
    _build(ready)
    ready.post("/api/calc/run", json={})
    assert ready.post("/api/checklist/add", json={}).status_code == 200
    c = ready.post("/api/order/commit", json={})
    assert not c.get_json()["ok"]


def test_5_ほかの道でチェックリストが変わっても票の指紋で断る(ready):
    _build(ready)
    with db.connect() as conn:                      # 画面を通らずに変わった(別の機能・手で直した など)
        conn.execute("UPDATE 発注チェックリスト SET 台数 = '9'")
        conn.commit()
    c = ready.post("/api/order/commit", json={})
    assert c.status_code == 409 and c.get_json()["reason"] == "checklist_changed"
    assert "作り直して" in c.get_json()["message"]
    assert ready.get("/report/order?t=test-token").status_code in (404, 409)


# ==================================================================
# 6. コイル実績
# ==================================================================
COIL_COLUMNS = ("ﾛｯﾄ番号", "BOX設計_横割数", "BOX設計_縦割数", "BOX実績_作業開始_日付",
                "BOX実績_ｺｲﾙ内径", "BOX実績_ｺｲﾙ肉厚", "BOX実績_枚本数")


def _coil_file(path: Path, rows: list[tuple]) -> Path:
    ym = date.today().strftime("%Y-%m")
    _write_table(path, "実績", list(COIL_COLUMNS),
                 [[lot, "12", "2", f"{ym}-01 08:00", inner, thick, ""] for lot, inner, thick in rows])
    return path


@pytest.fixture
def coil(ready, sandbox):
    konpo = sandbox / "konpo"
    user_settings.save(config.KEY_KONPO_DB_DIR, str(konpo))
    user_settings.save(config.KEY_KONPO_DB_FILE, "実績_LS4.sqlite3")
    path = _coil_file(konpo / "実績_LS4.sqlite3", [("Z999999", 500, 100)])
    auto()
    source_copy.forget()
    return path


def test_6_実績に無いロットは外径MAXを入れたと言う(coil, ready):
    ready.post("/api/calc/lot", json={"LOT": "A123456"})
    view = ready.post("/api/calc/order", json={"受注番号": "12345678"}).get_json()["view"]
    assert view["input"]["外径"] == "900"
    joined = " ".join(view["messages"])
    assert "実績に無いので受注の外径MAXを入れました" in joined
    assert "実績の取り込み:" in joined


def test_6_コイル実績のファイルが変わると取り込み直す(coil, ready):
    _coil_file(coil, [("Z999999", 500, 100), ("A123456", 508, 150)])
    nudge(coil)
    ready.get("/api/ribbon")
    assert local("SELECT 外径 FROM コイル実績 WHERE ロット番号 = 'A123456'") == [(808.0,)]


def test_6_コイル実績も写しの古さの対象(coil):
    with db.connect() as conn:
        assert coil_results_in(data_sync.expected_tables())
        conn.execute("DELETE FROM 取り込み履歴 WHERE テーブル名 = 'コイル実績'")
        conn.commit()
        assert data_sync.imported_at(conn) == ""


def coil_results_in(tables) -> bool:
    return "コイル実績" in tables


# ==================================================================
# 7. 「古い」は取り込み元が変わったかで出す
# ==================================================================
def test_7_マスタが何日も変わらなければ古くない(shared):
    with db.connect() as conn:                       # 3日前に写した。取り込み元はそのまま
        conn.execute("UPDATE 取り込み履歴 SET 取り込み日時 = '2026-01-01 08:00:00',"
                     " 確かめた日時 = '2026-01-01 08:00:00'")
        conn.commit()
        assert data_sync.stale_reason(conn) == data_sync.STALE_NONE
        # 確かめた日時は今になる(写した日時はそのまま)
        assert data_sync.checked_at(conn) > "2026-01-01 08:00:00"
        assert data_sync.imported_at(conn) == "2026-01-01 08:00:00"


def test_7_同じ日でも取り込み元が変われば古い(shared):
    _m, _h, odr = shared
    set_on_share(odr, "UPDATE 仕掛 SET 用途名 = 'かわった'")
    with db.connect() as conn:
        assert data_sync.stale_reason(conn) == data_sync.STALE_SOURCE_CHANGED
        assert "変わりました" in data_sync.stale_message(conn)


def test_7_帯は写した日時と確かめた日時を分けて出す(client, shared):
    rb = client.get("/api/ribbon").get_json()["ribbon"]
    assert rb["imported_at"] and rb["checked_at"]
    page = client.get("/calc?t=test-token").data.decode()
    assert "確かめた" in page


# ==================================================================
# 8. プログラムの規則が優先される品番の札
# ==================================================================
def test_8_規則のある品番はマスタ管理で札が出る(client, sandbox):
    folder = sandbox / "master"
    specs = import_specs.MASTER_IMPORT_SPECS["包装仕様"]
    cols = [s for _d, s, _c in specs]
    rows = []
    for no, size in (("1C1103", ""), ("1C1188", ""), ("1C0001", ""), ("1C1258", "外径指定")):
        row = {c: "" for c in cols}
        row.update({"包装仕様NO": no, "サイズ": size})
        rows.append([row[c] for c in cols])
    _write_table(folder / config.MATERIAL_DB_NAME, "包装仕様", cols, rows)
    view = client.get("/api/master/browse?table=包装仕様").get_json()
    page = view["page"]
    by_no = {r["包装仕様NO"]: page["rules"].get(str(r[page["row_key"]]), "") for r in page["rows"]}
    assert "この品番はプログラムの規則が優先されます" in by_no["1C1103"]
    assert "ｲｽﾞﾐﾒﾀﾙ" in by_no["1C1103"]
    assert "積数" in by_no["1C1188"] and "リプラの角サイズ" in by_no["1C1188"]
    assert "外径指定" in by_no["1C1258"]
    assert by_no["1C0001"] == ""


def test_8_規則の品番はすべて札の対象(sandbox):
    """規則を足したのに札を足し忘れる、を見つける。"""
    nos = {r.spec_no for r in special_rules.STACK_RULES} | {
        special_rules.IZUMI_SPEC_NO, special_rules.SHIMANO_SPEC_NO, special_rules.MIYAMA_SPEC_NO,
        special_rules.IMI_COVER_SPEC_NO, special_rules.KOTOBUKI_SPEC_NO}
    for no in nos:
        assert special_rules.program_rule_notes(no), no
    assert nos <= special_rules.RULE_SPEC_NOS


# ==================================================================
# 9. 押したときは写しを捨ててから読む
# ==================================================================
def test_9_いま取り込むは写しを捨ててから読む(client, shared, monkeypatch):
    master, hiki, odr = shared
    forgotten = []
    real = source_copy.forget
    monkeypatch.setattr(source_copy, "forget", lambda p=None: (forgotten.append(p), real(p)))
    client.post("/api/settings/import", json={})
    got = {Path(p) for p in forgotten if p is not None}
    assert {master, hiki, odr} <= got


def test_9_マスタ管理の取り込み直しも写しを捨ててから読む(shared, monkeypatch):
    master, _h, _o = shared
    forgotten = []
    real = source_copy.forget
    monkeypatch.setattr(source_copy, "forget", lambda p=None: (forgotten.append(p), real(p)))
    with db.connect() as conn:
        assert data_sync.reimport_table(conn, "パレット", path=master).ok
    assert master in [Path(p) for p in forgotten if p is not None]


# ==================================================================
# 10. 取り込み中
# ==================================================================
def test_10_取り込み中は帯に出て計算Startを断る(ready):
    with data_sync._importing():
        rb = ready.get("/api/ribbon").get_json()["ribbon"]
        assert rb["importing"] is True
        page = ready.get("/calc?t=test-token").data.decode()
        assert "取り込み中" in page
        ready.post("/api/calc/lot", json={"LOT": "A123456"})
        ready.post("/api/calc/order", json={"受注番号": "12345678"})
        ready.post("/api/calc/input", json={"検入数": "20"})
        r = ready.post("/api/calc/run", json={})
        assert r.status_code == 409 and r.get_json()["reason"] == "importing"
    assert ready.post("/api/calc/run", json={}).status_code == 200


def test_10_起動時の自動取り込みが走っている間も取り込み中(sandbox, monkeypatch):
    from modules.packing_material_calculation.coil_tool import jobs

    class Running:
        kind = "import"

    monkeypatch.setattr(jobs.get_registry(), "running", lambda: Running())
    assert data_sync.is_importing() is True
