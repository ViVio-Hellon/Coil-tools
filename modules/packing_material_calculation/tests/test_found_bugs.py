"""構造点検(統合 1.2.26)で見つかった不具合の再現

直す前に、**直ったら通る試験**として置く(`xfail(strict=True)`)。直したら印を外す。
印が付いたまま通ってしまったら(=気づかずに直った・別の所が変わった)、試験が落ちて知らせる。
"""
import pytest

from modules.packing_material_calculation.coil_tool import (burr_split, calc_service, checklist_service,
                                                         models)
from modules.packing_material_calculation.coil_tool.models import CalcState, OrderInfo, PackagingSpec
from modules.packing_material_calculation.tests import _fixture
from modules.packing_material_calculation.tests._web import client, sandbox, seed   # noqa: F401
from modules.packing_material_calculation.coil_tool import config


@pytest.fixture
def conn():
    c = _fixture.make_db()
    _fixture.standard_master(c)
    yield c
    c.close()


def run(conn, *, 検入数, 枚数, 枚数範囲="1不可", 単重="100.00", 外径="900", spec_no="1C0001"):
    _fixture.add_spec(conn, spec_no, 枚数範囲=枚数範囲)
    s = CalcState(検入数=str(検入数), 外径=外径)
    s.order = OrderInfo(包装仕様NO=spec_no, 受注板厚="1.000", 受注板幅="100.0",
                        製品単重=単重, 梱包単位_重量=OrderInfo.NO_DATA,
                        梱包単位_枚数=str(枚数))
    return s, calc_service.calculate(conn, s)


# ==================================================================
# 積数を手で直したとき(VBA `積数_AfterUpdate` のバグ。現場で確認済み)
# ==================================================================
def test_1本積み不可で分けたあと積数を直すと_別枠は消える(conn):
    s, r = run(conn, 検入数=13, 枚数=4)
    assert r.ok and s.Re_台数 == "2"            # 2台 + 別枠2台
    s.積数 = "5"
    again = calc_service.recalculate_from_stack(conn, s)
    assert again.ok
    assert s.台数 == "3"                         # 切り上げ(13/5)
    assert s.Re_台数 == "" and s.Re_積数 == ""
    assert checklist_service._unit_count(s) == "3"
    assert models.FLAG_SINGLE_SPLIT not in s.flags


# ==================================================================
# 検入数が小さい・0・負
# ==================================================================
@pytest.mark.parametrize("枚数", [2, 3, 4])
def test_1本積み不可で検入数1でも台数は負にならない(conn, 枚数):
    s, r = run(conn, 検入数=1, 枚数=枚数)
    assert not r.ok and r.reason == "no_single_one_coil"
    assert s.台数 == "" and s.積数 == ""
    assert any("振り分けができません" in m for m in s.messages)


@pytest.mark.parametrize("検入数", [0, -5])
def test_検入数が0や負なら計算しない(conn, 検入数):
    s, r = run(conn, 検入数=検入数, 枚数=4, 枚数範囲="以下")
    assert not r.ok and r.reason == "bad_inspection_count"


@pytest.mark.parametrize("検入数", ["0", "-5"])
def test_検入数の欄に0や負を入れると断って前の値のまま(ready, 検入数):
    r = ready.post("/api/calc/input", json={"検入数": 検入数})
    assert r.status_code == 422
    assert r.get_json()["reason"] == "not_positive"
    assert r.get_json()["view"]["input"]["検入数"] == "10"


def test_バリ揃えも検入数0なら断る(conn):
    s = CalcState(検入数="0")
    s.order = OrderInfo(包装仕様NO="1C0001", 製品単重="100.00")
    assert burr_split.calculate(conn, s, upper=0, lower=0).reason == "no_inspection_count"


# ==================================================================
# バリ揃え梱包
# ==================================================================
def test_バリ揃えの積数も整数に切り捨てる():
    order = OrderInfo(包装仕様NO="X", 製品単重="100.00", 梱包単位_重量=OrderInfo.NO_DATA,
                      梱包単位_枚数="2.5")
    assert burr_split._decide_stack(order, PackagingSpec()) == 2.0


def test_バリ揃えの結果転送で別枠は消える(conn):
    s, r = run(conn, 検入数=13, 枚数=4)
    assert s.Re_台数 == "2"
    burr = burr_split.BurrSplit(上本数=7, 下本数=6, 残り検入数=0, 積数=4,
                                上台数=2, 下台数=2, 台数=4, ok=True)
    burr_split.transfer(s, burr)
    assert checklist_service._unit_count(s) == "4"


# ==================================================================
# 古い結果をチェックリストへ積めてしまう
# ==================================================================
@pytest.fixture
def ready(client, sandbox):
    seed(sandbox)
    client.post("/api/settings", json={"worker": config.WORKERS[0], "line": "LS4"})
    client.post("/api/calc/lot", json={"LOT": "A123456"})
    client.post("/api/calc/order", json={"受注番号": "12345678"})
    client.post("/api/calc/input", json={"検入数": "10", "外径": "900"})
    assert client.post("/api/calc/run", json={}).status_code == 200
    return client


def test_検入数を変えたら計算し直すまで積めない(ready):
    ready.post("/api/calc/input", json={"検入数": "20"})
    r = ready.post("/api/checklist/add", json={})
    assert r.status_code == 422


def test_積数を直しても分けない_HBもリプラも3台ぶん(conn):
    s, r = run(conn, 検入数=13, 枚数=4)
    hb_before = s.HB枚数
    s.積数 = "5"
    assert calc_service.recalculate_from_stack(conn, s).ok
    assert s.HB枚数 != hb_before
    assert s.HB枚数 == "3"
    assert not s.messages, "前の計算の断り・知らせが残っている"


@pytest.mark.parametrize("積数", ["0", ""])
def test_積数の手直しが断られたら_前の台数で積めない(ready, 積数):
    ready.post("/api/calc/input", json={"積数": 積数})
    r = ready.post("/api/calc/restack", json={})
    assert r.status_code == 422
    assert r.get_json()["view"]["result"]["台数"] == ""
    assert ready.post("/api/checklist/add", json={}).status_code == 422


def test_外径を変えたら消す_同じ値の書き方違いなら消さない(ready):
    view = ready.post("/api/calc/input", json={"外径": "900.0"}).get_json()["view"]
    assert view["result"]["台数"] != "", "同じ値なのに消した"
    body = ready.post("/api/calc/input", json={"外径": "950"}).get_json()
    assert body["view"]["result"]["台数"] == ""
    assert any("前の計算結果を消しました" in m for m in body["view"]["messages"])


def test_バリ揃えの結果転送のあと積める(ready):
    r = ready.post("/api/calc/burr/transfer", json={"上本数": 6, "下本数": 4})
    assert r.status_code == 200, r.get_json()
    assert r.get_json()["view"]["result"]["台数"] == "3"
    assert ready.post("/api/checklist/add", json={}).status_code == 200
