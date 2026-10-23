"""パレット種類_外径条件 ── 1C1103 の外径の規則をマスタ優先にする(統合 1.2.50・現場の指示)

現場の指示:「本来マスタ優先であるべきだが、マスタに外径条件が無かったのでコードで対応したと
思われる。マスタ条件に外径を追加し、規則はマスタ優先にするべき」。ここで確かめること:

    1  表が無いあいだは、いままでどおりプログラムの規則(境目 799.99 / 800 / 990 / 990.01・外径なし)
    2  表を作ると、プログラムの規則と同じ結果になる
    3  表の行を直すと結果が変わる(マスタ優先。プログラムの規則は使わない)
    4  表に行の無い包装仕様No は、マスタ(包装仕様)の値のまま
    5  取り込み(目印・鮮度の確かめ・取り込み直し・世代)に乗る。共有に無くても誤りにしない
    6  表を作るにはパスワードが要る
    7  マスタ管理の「規則が優先」の札は、表が無いあいだだけ 1C1103 に出す
"""
from __future__ import annotations

import sqlite3

import pytest

from modules.packing_material_calculation.coil_tool import (admin_session, config, data_sync, db,
                                                            import_specs, master_admin,
                                                            outer_condition, pallet_service,
                                                            special_rules)
from modules.packing_material_calculation.coil_tool.models import CalcState, OrderInfo
from modules.packing_material_calculation.tests import _fixture
from modules.packing_material_calculation.tests._web import client, sandbox   # noqa: F401
from modules.packing_material_calculation.tests.test_web_master import make_source, unlock

TABLE = config.TBL_OUTER_CONDITION
IZUMI = special_rules.IZUMI_SPEC_NO
SOURCE_COLS = [s for _d, s, _c in import_specs.OUTER_CONDITION_COLUMNS]

# 境目と、その前後。外径は小数2桁にそろえて入る
BOUNDARIES = [
    ("", "全面"),            # 外径なし(0 として見る。0 < 800 なので全面)
    ("0", "全面"),
    ("700", "全面"),
    ("780", "全面"),         # VBA の境目の重なり。先勝ちで全面
    ("799.99", "全面"),
    ("800", "強度UP"),
    ("889.99", "強度UP"),
    ("890", "強度UP"),
    ("990", "強度UP"),
    ("990.01", "スカシ"),    # 990 を超えたらマスタの値のまま
    ("1200", "スカシ"),
]


# ==================================================================
# 下ごしらえ
# ==================================================================
@pytest.fixture
def conn():
    c = _fixture.make_db()
    _fixture.standard_master(c)
    _fixture.add_spec(c, IZUMI, パレット="スカシ")
    _fixture.add_spec(c, "1C0001", パレット="全面")
    yield c
    c.close()


def make_local_table(conn, rows=outer_condition.SEED_ROWS) -> None:
    """手元に表を作って行を入れる(取り込んだのと同じ形)。"""
    outer_condition.ensure_local(conn)
    with conn:
        conn.execute(f'DELETE FROM "{TABLE}"')
        conn.executemany(
            f'INSERT INTO "{TABLE}" (包装仕様No, 順番, 外径_以上, 外径_以下, パレット種類, 備考)'
            " VALUES (?, ?, ?, ?, ?, ?)",
            [(r["包装仕様No"], r["順番"], r["外径_以上"], r["外径_以下"], r["パレット種類"], r["備考"])
             for r in rows])


def kind(conn, spec_no: str, outer: str) -> str:
    """計算と同じ入口(`pallet_service.pallet_type`)でパレット種類を決める。"""
    s = CalcState(検入数="10", 外径=outer)
    s.order = OrderInfo(包装仕様NO=spec_no)
    return pallet_service.pallet_type(conn, s)


def put_source_table(path, rows, cols=None) -> None:
    """共有の梱包資材マスタに外径条件の表を直に作る(ほかの端末が作った、に見立てる)。"""
    cols = cols or SOURCE_COLS
    src = sqlite3.connect(path)
    src.execute(f'DROP TABLE IF EXISTS "{TABLE}"')
    src.execute(f'CREATE TABLE "{TABLE}" (%s)' % ", ".join(f'"{c}"' for c in cols))
    src.executemany(f'INSERT INTO "{TABLE}" VALUES (%s)' % ", ".join("?" for _ in cols),
                    [[r.get(c) for c in cols] for r in rows])
    src.commit()
    src.close()
    nudge(path)


def drop_source_table(path) -> None:
    src = sqlite3.connect(path)
    src.execute(f'DROP TABLE IF EXISTS "{TABLE}"')
    src.commit()
    src.close()
    nudge(path)


def nudge(path) -> None:
    """書いたあと、更新時刻を確実に動かす(時計の刻みが粗い環境でも、試験の意図どおりにする)。"""
    import os
    st = path.stat()
    os.utime(path, ns=(st.st_atime_ns, st.st_mtime_ns + 7_000_000_123))


def source_rows(path) -> list[dict]:
    src = sqlite3.connect(path)
    src.row_factory = sqlite3.Row
    try:
        return [dict(r) for r in src.execute(f'SELECT * FROM "{TABLE}" ORDER BY rowid')]
    finally:
        src.close()


@pytest.fixture
def shared(sandbox):
    """共有の梱包資材マスタ(外径条件の表はまだ無い)と、そこから取り込んだ手元DB。"""
    path = make_source(sandbox)
    with db.connect() as c:
        db.apply_schema(c)
        data_sync.import_master(c)
    admin_session.reset()
    yield path
    admin_session.reset()


# ==================================================================
# 1. 表が無いあいだは、いままでどおりプログラムの規則
# ==================================================================
@pytest.mark.parametrize("outer,expected", BOUNDARIES)
def test_1_表が無ければプログラムの規則どおり(conn, outer, expected):
    assert not outer_condition.present(conn)
    assert kind(conn, IZUMI, outer) == expected


def test_1_外径なしは今の規則でも0として見ている(conn):
    """外径が空のときの扱いを、今のコードで確かめておく(`vba.val("")` = 0 → 0 < 800 → 全面)。"""
    assert special_rules.pallet_type_override(IZUMI, 0.0, "スカシ") == "全面"
    assert kind(conn, IZUMI, "") == "全面"


def test_1_表が無ければほかの包装仕様はマスタの値(conn):
    assert kind(conn, "1C0001", "700") == "全面"
    assert kind(conn, "1C0001", "1200") == "全面"


# ==================================================================
# 2. 表を作ると、プログラムの規則と同じ結果
# ==================================================================
@pytest.mark.parametrize("outer,expected", BOUNDARIES)
def test_2_表を作ると同じ結果(conn, outer, expected):
    make_local_table(conn)
    assert outer_condition.present(conn)
    assert kind(conn, IZUMI, outer) == expected


def test_2_作る行は細かく刻んでも規則と同じ結果(conn):
    """0〜1200 を 0.01 刻みで比べる(外径は小数2桁にそろえて入る)。"""
    make_local_table(conn)
    for i in range(0, 120001, 1):
        outer = i / 100
        want = special_rules.pallet_type_override(IZUMI, outer, "スカシ")
        got = outer_condition.pallet_type(conn, IZUMI, outer, "スカシ")
        assert got == want, outer


def test_2_作る行の中身(sandbox):
    rows = outer_condition.SEED_ROWS
    assert [(r["包装仕様No"], r["順番"], r["外径_以上"], r["外径_以下"], r["パレット種類"]) for r in rows] == [
        (IZUMI, 1, None, 799.99, "全面"),
        (IZUMI, 2, 800.0, 990.0, "強度UP"),
    ]
    assert rows[0]["備考"] == "VBA の 外径<800"
    assert rows[1]["備考"] == "VBA の 780〜990(780〜799 は先の条件が勝つ)"


# ==================================================================
# 3. 表の行を直すと結果が変わる(マスタ優先)
# ==================================================================
def test_3_行を直すと結果が変わる(conn):
    rows = [dict(r) for r in outer_condition.SEED_ROWS]
    rows[1]["外径_以下"] = 900.0          # 強度UP を 900 までにする
    rows[0]["パレット種類"] = "EXPS"     # 800 未満を EXPS にする
    make_local_table(conn, rows)
    assert kind(conn, IZUMI, "700") == "EXPS"
    assert kind(conn, IZUMI, "900") == "強度UP"
    assert kind(conn, IZUMI, "950") == "スカシ"   # プログラムの規則なら強度UP。マスタが優先


def test_3_順番の小さい行が先に勝つ(conn):
    make_local_table(conn, [
        {"包装仕様No": IZUMI, "順番": 2, "外径_以上": None, "外径_以下": None,
         "パレット種類": "EXPS", "備考": ""},
        {"包装仕様No": IZUMI, "順番": 1, "外径_以上": 850.0, "外径_以下": None,
         "パレット種類": "強度UP", "備考": ""},
    ])
    assert kind(conn, IZUMI, "900") == "強度UP"
    assert kind(conn, IZUMI, "700") == "EXPS"


def test_3_ほかの包装仕様Noも表で決められる(conn):
    make_local_table(conn, [*outer_condition.SEED_ROWS,
                            {"包装仕様No": "1C0001", "順番": 1, "外径_以上": 1000.0,
                             "外径_以下": None, "パレット種類": "強度UP", "備考": ""}])
    assert kind(conn, "1C0001", "1000") == "強度UP"
    assert kind(conn, "1C0001", "999.99") == "全面"


# ==================================================================
# 4. 表に行の無い包装仕様No はマスタの値のまま
# ==================================================================
def test_4_行の無い包装仕様はマスタの値のまま(conn):
    make_local_table(conn)
    assert kind(conn, "1C0001", "700") == "全面"
    assert kind(conn, "1C0001", "900") == "全面"


def test_4_表はあるが1C1103の行が無ければプログラムの規則は使わない(conn):
    make_local_table(conn, [])
    assert outer_condition.present(conn)
    assert kind(conn, IZUMI, "700") == "スカシ"
    assert kind(conn, IZUMI, "900") == "スカシ"


# ==================================================================
# 5. 取り込みに乗る
# ==================================================================
def test_5_共有に表が無くても取り込みの誤りにしない(shared):
    with db.connect() as c:
        results = {t.table: t for t in data_sync.import_master(c)}
        got = results[TABLE]
        assert got.ok and got.rows == 0
        assert "まだ作っていません" in got.note
        assert not outer_condition.present(c)
        # 見たことは目印で残す(共有で作られたら「変わった」と気づくため)
        record = data_sync.last_import(c)[TABLE]
        assert record["元の目印"] == data_sync.source_stamp(shared)
        # 帯の「古い」の対象にはしない
        assert TABLE not in data_sync.expected_tables()
        assert data_sync.stale_reason(c) != data_sync.STALE_SOURCE_CHANGED


def test_5_いま取り込むも誤りにしない(client, sandbox):
    make_source(sandbox)
    client.post("/api/settings", json={config.KEY_MASTER_DB_DIR: str(sandbox / "master")})
    body = client.post("/api/settings/import", json={}).get_json()
    row = next(t for t in body["tables"] if t["table"] == TABLE)
    assert row["ok"] is True
    assert "まだ作っていません" in row["note"]


def test_5_共有で表が作られたら確かめのときに取り込み直す(shared):
    put_source_table(shared, outer_condition.SEED_ROWS)
    with db.connect() as c:
        before = data_sync.generation(c)
        fresh = data_sync.freshness(c)
        assert TABLE in fresh.refreshed
        assert outer_condition.present(c)
        assert c.execute(f'SELECT COUNT(*) FROM "{TABLE}"').fetchone()[0] == 2
        assert data_sync.last_import(c)[TABLE]["元の目印"] == data_sync.source_stamp(shared)
        assert data_sync.generation(c) != before          # 計算の世代も進む
        # もう一度確かめても、変わっていなければ読み直さない
        again = data_sync.freshness(c)
        assert TABLE not in again.refreshed
        assert not again.check.optional


def test_5_共有で表が消されたら手元も消してプログラムの規則に戻る(shared):
    put_source_table(shared, outer_condition.SEED_ROWS)
    with db.connect() as c:
        data_sync.freshness(c)
        assert outer_condition.present(c)
    drop_source_table(shared)
    with db.connect() as c:
        data_sync.freshness(c)
        assert not outer_condition.present(c)


def test_5_前の版の手元DBでも1回の確かめで取り込む(shared):
    """取り込み履歴にまだこの表が無い(前の版で取り込んだ)手元DBでも、確かめのときに見に行く。"""
    put_source_table(shared, outer_condition.SEED_ROWS)
    with db.connect() as c:
        data_sync.import_master(c)
        with c:
            c.execute("DELETE FROM 取り込み履歴 WHERE テーブル名 = ?", (TABLE,))
        outer_condition.drop_local(c)
        fresh = data_sync.freshness(c)
        assert TABLE in fresh.refreshed
        assert outer_condition.present(c)
        # 任意の表だけが変わっていても「古い」とは言わない
        assert not fresh.check.changed


def test_5_起動時の自動取り込みにも乗る(shared):
    put_source_table(shared, outer_condition.SEED_ROWS)
    with db.connect() as c:
        result = data_sync.auto_import(c)
        assert TABLE in [t.table for t in result.tables]
        assert outer_condition.present(c)


def test_5_空の外径は空のまま取り込む(shared):
    put_source_table(shared, outer_condition.SEED_ROWS)
    with db.connect() as c:
        assert data_sync.reimport_table(c, TABLE).ok
        rows = c.execute(f'SELECT 外径_以上, 外径_以下 FROM "{TABLE}" ORDER BY 順番').fetchall()
        assert [tuple(r) for r in rows] == [(None, 799.99), (800.0, 990.0)]


def test_5_列が足りない表は誤りにして手元に空の表を作らない(shared):
    put_source_table(shared, [{"包装仕様No": IZUMI}], cols=["包装仕様No", "順番"])
    with db.connect() as c:
        got = data_sync.reimport_table(c, TABLE)
        assert not got.ok
        assert "列がありません" in got.error
        assert not outer_condition.present(c)       # 空の表に化けない(= 規則のまま)


def test_5_取り込み直しは写しを捨ててから読む(shared, monkeypatch):
    from common import source_copy
    forgotten = []
    real = source_copy.forget
    monkeypatch.setattr(source_copy, "forget", lambda p=None: (forgotten.append(p), real(p)))
    put_source_table(shared, outer_condition.SEED_ROWS)
    with db.connect() as c:
        assert data_sync.reimport_table(c, TABLE).ok
    assert shared in forgotten


# ==================================================================
# 6. 表を作る(パスワードが要る)
# ==================================================================
def test_6_表を作るにはパスワードが要る(client, sandbox):
    path = make_source(sandbox)
    client.post("/api/settings", json={config.KEY_MASTER_DB_DIR: str(sandbox / "master")})
    r = client.post("/api/master/table/create", json={"table": TABLE})
    assert r.status_code == 403
    assert r.get_json()["error"]["code"] == master_admin.REFUSE_NOT_ALLOWED
    with sqlite3.connect(path) as src:
        assert not src.execute("SELECT 1 FROM sqlite_master WHERE name = ?", (TABLE,)).fetchone()


def test_6_パスワードを通すと作れて手元も追いつく(client, sandbox):
    path = make_source(sandbox)
    client.post("/api/settings", json={config.KEY_MASTER_DB_DIR: str(sandbox / "master")})
    unlock(client)
    r = client.post("/api/master/table/create", json={"table": TABLE})
    assert r.status_code == 200, r.get_json()
    body = r.get_json()
    assert "作り" in body["message"]
    assert [(x["包装仕様No"], x["順番"], x["外径_以上"], x["外径_以下"], x["パレット種類"], x["備考"])
            for x in source_rows(path)] == [
        (r_["包装仕様No"], r_["順番"], r_["外径_以上"], r_["外径_以下"], r_["パレット種類"], r_["備考"])
        for r_ in outer_condition.SEED_ROWS]
    with db.connect() as c:
        assert outer_condition.present(c)
        _fixture.add_spec(c, IZUMI, パレット="スカシ")
        for outer, expected in BOUNDARIES:
            assert kind(c, IZUMI, outer) == expected, outer
    # 2回目は「もうある」
    again = client.post("/api/master/table/create", json={"table": TABLE})
    assert again.status_code == 409
    assert again.get_json()["error"]["code"] == master_admin.REFUSE_ALREADY_EXISTS


def test_6_作れる表はこの表だけ(client, sandbox):
    make_source(sandbox)
    client.post("/api/settings", json={config.KEY_MASTER_DB_DIR: str(sandbox / "master")})
    unlock(client)
    r = client.post("/api/master/table/create", json={"table": "包装仕様"})
    assert r.status_code == 422
    assert r.get_json()["error"]["code"] == master_admin.REFUSE_NOT_EDITABLE


def test_6_無いあいだは一覧に出て表を作るが出る(client, sandbox):
    make_source(sandbox)
    client.post("/api/settings", json={config.KEY_MASTER_DB_DIR: str(sandbox / "master")})
    body = client.get(f"/api/master/browse?table={TABLE}").get_json()
    info = next(t for t in body["tables"] if t["table"] == TABLE)
    assert info["missing"] is True and info["creatable"] is True
    assert body["page"]["missing"] is True and body["page"]["creatable"] is True
    assert "まだ作っていません" in body["page"]["why"]
    # 上流の表(包装仕様など)は、今までどおり作れない(作るボタンは出ない)
    spec = client.get("/api/master/browse?table=包装仕様").get_json()
    assert spec["page"]["creatable"] is False


def test_6_作った表は行を足す直す消すができ計算に効く(client, sandbox):
    path = make_source(sandbox)
    client.post("/api/settings", json={config.KEY_MASTER_DB_DIR: str(sandbox / "master")})
    unlock(client)
    client.post("/api/master/table/create", json={"table": TABLE})
    body = client.get(f"/api/master/browse?table={TABLE}").get_json()
    assert body["page"]["editable"] is True
    names = [c["name"] for c in body["columns"]]
    assert names == SOURCE_COLS
    required = {c["name"] for c in body["columns"] if c["required"]}
    assert required == {"包装仕様No", "パレット種類"}

    key = body["page"]["rows"][1][body["row_key"]]
    r = client.post("/api/master/row/save",
                    json={"table": TABLE, "key": key, "values": {"外径_以下": "900"}})
    assert r.status_code == 200, r.get_json()
    r = client.post("/api/master/row/add",
                    json={"table": TABLE, "values": {"包装仕様No": "1C0001", "順番": "1",
                                                     "外径_以上": "", "外径_以下": "",
                                                     "パレット種類": "EXPS", "備考": ""}})
    assert r.status_code == 200, r.get_json()
    assert source_rows(path)[-1]["外径_以上"] is None           # 空は空のまま(下限なし)
    with db.connect() as c:
        _fixture.add_spec(c, IZUMI, パレット="スカシ")
        _fixture.add_spec(c, "1C0001", パレット="全面")
        assert kind(c, IZUMI, "950") == "スカシ"               # マスタ優先(規則なら強度UP)
        assert kind(c, "1C0001", "700") == "EXPS"

    first = client.get(f"/api/master/browse?table={TABLE}").get_json()
    key0 = first["page"]["rows"][0][first["row_key"]]
    r = client.post("/api/master/row/delete", json={"table": TABLE, "key": key0})
    assert r.status_code == 200
    with db.connect() as c:
        _fixture.add_spec(c, IZUMI, パレット="スカシ")
        assert kind(c, IZUMI, "700") == "スカシ"               # 行を消したらマスタの値


def test_6_表を作るのは1つのまとまり(shared, monkeypatch):
    """途中で落ちたら表も残さない(空の表が残ると、1C1103 がマスタの値のままになる)。"""
    admin_session.open_with(config.ADMIN_PASSWORD)
    # 2行目を入れるところで落ちるようにする
    monkeypatch.setattr(outer_condition, "SOURCE_DDL",
                        outer_condition.SOURCE_DDL[:-1] + ', CHECK ("順番" < 2))')
    with db.connect() as c:
        got = master_admin.create_table(c, TABLE)
        assert not got.ok
        assert got.reason == master_admin.REFUSE_WRITE_FAILED
        assert not outer_condition.present(c)
    with sqlite3.connect(shared) as src:
        assert not src.execute("SELECT 1 FROM sqlite_master WHERE name = ?", (TABLE,)).fetchone()


# ==================================================================
# 7. 「規則が優先」の札
# ==================================================================
def _tags(client) -> dict[str, str]:
    view = client.get("/api/master/browse?table=包装仕様").get_json()
    page = view["page"]
    return {r["包装仕様NO"]: page["rules"].get(str(r[page["row_key"]]), "") for r in page["rows"]}


def _spec_source(sandbox):
    """包装仕様に 1C1103・1C1188 がある共有のマスタ。"""
    path = make_source(sandbox)
    cols = [s for _d, s, _c in import_specs.SPEC_COLUMNS]
    src = sqlite3.connect(path)
    for no in (IZUMI, "1C1188"):
        row = {c: "" for c in cols}
        row.update({"包装仕様NO": no, "パレット": "スカシ"})
        src.execute('INSERT INTO "包装仕様" (%s) VALUES (%s)' % (
            ", ".join(f'"{c}"' for c in cols), ", ".join("?" for _ in cols)), [row[c] for c in cols])
    src.commit()
    src.close()
    return path


def test_7_表が無いあいだは1C1103に札が出る(client, sandbox):
    _spec_source(sandbox)
    client.post("/api/settings", json={config.KEY_MASTER_DB_DIR: str(sandbox / "master")})
    tags = _tags(client)
    assert "プログラムの規則が優先されます" in tags[IZUMI]
    assert "マスタ管理で「パレット種類_外径条件」の表を作るとマスタが優先になります" in tags[IZUMI]
    assert "積数" in tags["1C1188"]


def test_7_表を作ると1C1103の札は消え他の品番の札は残る(client, sandbox):
    _spec_source(sandbox)
    client.post("/api/settings", json={config.KEY_MASTER_DB_DIR: str(sandbox / "master")})
    unlock(client)
    assert client.post("/api/master/table/create", json={"table": TABLE}).status_code == 200
    tags = _tags(client)
    assert tags[IZUMI] == ""
    assert "積数" in tags["1C1188"] and "リプラの角サイズ" in tags["1C1188"]


def test_7_札の出し分け(sandbox):
    assert special_rules.program_rule_notes(IZUMI) == [special_rules.IZUMI_NOTE]
    assert special_rules.program_rule_notes(IZUMI, outer_table=True) == []
    assert special_rules.program_rule_notes("1C1188", outer_table=True)
    assert special_rules.program_rule_notes("1C1258", size="外径指定", outer_table=True)
