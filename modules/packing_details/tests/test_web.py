"""画面とAPI (VBA `frmCoilPacking` のイベントに対応する経路)

**本物と同じ経路を通す。** トークンを付け、Flask のテストクライアントで
実際に叩く ── サービス層を直接呼ぶ試験は別にあるので、ここは
「画面から操作したときに通るか」だけを見る。
"""
from __future__ import annotations

import time
import unittest

from modules.packing_details.meisai import (app_config, config, data_sync, history_repo,
                                           meisai_service, screen, session)

from . import _db

TOKEN = "test-token-abc123"
# 画面の名前。**試験も1枚の画面として振る舞う** ── 持っていない画面の
# 操作はサーバが断るので、名乗らないと全部 409 になる
SCREEN = "test-screen-1"
HEADERS = {"X-Tool-Token": TOKEN, "X-Tool-Screen": SCREEN}


def _client(case: unittest.TestCase):
    """トークン付きのクライアントと、掴ませたDB接続。"""
    from modules.packing_details.app import routes                      # noqa: F401  (登録のため)
    from modules.packing_details import app as app_module

    conn = _db.case_db(case)
    original = app_module.get_db
    app_module.get_db = lambda: conn
    case.addCleanup(setattr, app_module, "get_db", original)

    # ルート側は `from .. import get_db` で名前を束縛済みなので、
    # そちらも差し替える。**持っているものだけ** ── `health` は
    # DBを触らないので `get_db` を import していない
    from modules.packing_details.app.routes import health as health_routes
    from modules.packing_details.app.routes import history as history_routes
    from modules.packing_details.app.routes import master as master_routes
    from modules.packing_details.app.routes import meisai as meisai_routes
    from modules.packing_details.app.routes import settings as settings_routes
    for module in (health_routes, history_routes, master_routes, meisai_routes,
                   settings_routes):
        if not hasattr(module, "get_db"):
            continue
        case.addCleanup(setattr, module, "get_db", module.get_db)
        module.get_db = lambda: conn

    flask_app = app_module.create_app(token=TOKEN)
    flask_app.config["READY"] = True
    client = flask_app.test_client()
    case.addCleanup(session.clear)
    session.clear()
    # 画面は**プロセスに1つ**。試験ごとに取り直す
    case.addCleanup(screen.reset)
    screen.reset()
    screen.claim(SCREEN)
    return client, conn


def _post(client, path, data=None):
    return client.post(path, json=data or {}, headers=HEADERS)


def _seed(conn, lot="L5160Z0", tate="2", yoko="12"):
    """1ロット分の取り込み済みデータを手で入れる。"""
    with conn:
        conn.execute(
            "INSERT INTO 仕掛ロット (ロット番号, 用途コード, 用途名, 製造材質,"
            " 製造調質, 製造板厚, 製造板幅, オーダー板厚, オーダー板幅,"
            " オーダー板丈, 設計_設備コース, 実績_設備コース)"
            " VALUES (?, 'R192', 'ｱﾝｾﾞﾝﾀｲｻﾝｺｰ', '52S', 'H34',"
            " 1.985, 104.0, 2.0, 104.0, 0, 'HOT LS4', 'HOT')", (lot,))
        conn.execute(
            "INSERT INTO 仕掛当工程 (ロット番号, 当工程設計_縦割数,"
            " 当工程設計_横割数) VALUES (?, ?, ?)", (lot, int(tate), int(yoko)))
        conn.execute(
            "INSERT INTO 仕掛引当 (ロット番号, 受注番号) VALUES (?, '62015946')",
            (lot,))
        conn.execute(
            "INSERT INTO 仕掛受注 (受注番号, 包装仕様NO) VALUES ('62015946', '1C0123')")


class SecurityTest(unittest.TestCase):
    """トークンと同一オリジンの守り。"""

    def setUp(self):
        self.client, self.conn = _client(self)

    def test_トークンが無ければ断る(self):
        res = self.client.post("/api/lot", json={"lot_no": "L5160Z0"})
        self.assertEqual(res.status_code, 401)
        self.assertEqual(res.get_json()["error"]["code"], "bad_token")

    def test_健康確認はトークン不要(self):
        res = self.client.get("/api/health")
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.get_json()["app_id"], "nlm.packing-details")

    def test_健康確認が版の出どころを名乗る(self):
        """**「入れ替えたのに古いまま」を調べるのに要る。**

        現場は端末へフォルダごとコピーして配るので、版だけ見ても
        「どの写しが動いているか」が言えない。実行フォルダとPythonを
        一緒に返す。版の書き方が壊れていれば理由もここに出る。
        """
        body = self.client.get("/api/health").get_json()
        self.assertEqual(body["version"], app_config.version())
        self.assertTrue(body["app_root"])
        self.assertRegex(body["python"], r"^\d+\.\d+\.\d+")
        self.assertTrue(body["python_exe"])
        self.assertEqual(body["version_problem"], "")   # 版の形は正しい
        self.assertIn("db_path_problem", body)          # 共有に置かれていないか

    def test_心拍はトークン不要(self):
        self.assertEqual(self.client.post("/api/alive", json={}).status_code, 200)

    def test_別オリジンからは断る(self):
        res = self.client.post("/api/lot", json={"lot_no": "L5160Z0"},
                               headers={**HEADERS,
                                        "Sec-Fetch-Site": "cross-site"})
        self.assertEqual(res.status_code, 403)

    def test_帳票もトークンが要る(self):
        """業務データを返す経路。`/api/` だけを守ると素通しになる。"""
        res = self.client.get("/report/L5160Z0/1")
        self.assertEqual(res.status_code, 401)

    def test_別のホスト名では断る(self):
        res = self.client.get("/api/health", headers={"Host": "example.com"})
        self.assertEqual(res.status_code, 400)

    def test_画面のHTMLは控えさせない(self):
        screen.reset()                     # 1枚目として開く
        res = self.client.get("/meisai")
        self.assertEqual(res.status_code, 200)
        self.assertEqual(res.headers["Cache-Control"], "no-store")


class ScreenTest(unittest.TestCase):
    """画面は1枚だけ (タブを2枚開かせない)。

    **プロセスの多重起動とは別の話。** `launch_guard` は Python が
    2つ走っていないかを見るが、タブを2枚開くのは1プロセスのままなので
    止められない。ところが作業状態(`session.py`)はプロセスに1つしか
    無いので、2枚開くと同じ盤面を奪い合う ── 自分が開いたはずのロットが
    触った瞬間に別のロットへ化ける。
    """

    def setUp(self):
        self.client, self.conn = _client(self)
        _seed(self.conn)

    def _claim(self, screen_id):
        return self.client.post("/api/screen/claim", json={"screen": screen_id},
                                headers={"X-Tool-Token": TOKEN})

    def test_2枚目の名乗りは断る(self):
        """`_client` が SCREEN を持っている＝1枚目が開いている状態。

        **器(HTML)は誰にでも返す。** ブラウザは再読込のとき古いページを
        畳む前に次の要求を送るので、開くところで断ると F5 のたびに
        断ることになる。見分けるのはブラウザ側(`screen.js`)の仕事で、
        サーバは名乗りを断る。
        """
        res = self._claim("2枚目")
        self.assertEqual(res.status_code, 409)
        self.assertEqual(res.get_json()["error"]["code"], "screen_busy")
        self.assertEqual(screen.holder().screen_id, SCREEN)

    def test_2枚目の覆いは前の画面の様子を取り直せる(self):
        """読み直さずに「前の画面からの応答」を数え直す(統合 1.2.11)。2枚目(別の名前)からも読める。"""
        res = self.client.get("/api/screen/state",
                              headers={"X-Tool-Token": TOKEN, "X-Tool-Screen": "2枚目"})
        self.assertEqual(res.status_code, 200)
        body = res.get_json()
        self.assertTrue(body["busy"])
        self.assertIn("grace_sec", body)

    def test_同じ名前の名乗り直しは通る(self):
        """**再読込(F5)がこれ。** タブの名前は sessionStorage に残る。"""
        self.assertEqual(self._claim(SCREEN).status_code, 200)
        self.assertEqual(screen.holder().screen_id, SCREEN)

    def test_器は誰にでも返す(self):
        """業務データは名乗ってからでないと触れないので、器は返してよい。"""
        res = self.client.get("/meisai")
        self.assertEqual(res.status_code, 200)

    def test_持っていない画面の操作は断る(self):
        """ここが本当の歯止め。開くところの断りはすり抜けうる。"""
        res = self.client.post("/api/lot", json={"lot_no": "L5160Z0"},
                               headers={"X-Tool-Token": TOKEN,
                                        "X-Tool-Screen": "べつの画面"})
        self.assertEqual(res.status_code, 409)
        self.assertEqual(res.get_json()["error"]["code"], "screen_taken")

    def test_画面を名乗らない操作も断る(self):
        res = self.client.post("/api/lot", json={"lot_no": "L5160Z0"},
                               headers={"X-Tool-Token": TOKEN})
        self.assertEqual(res.status_code, 409)

    def test_引き継ぐと前の画面が使えなくなる(self):
        """**2枚にするのではなく、移す。**"""
        self.assertEqual(_post(self.client, "/api/lot",
                               {"lot_no": "L5160Z0"}).status_code, 200)

        # 2枚目が引き継ぐ → 誰も持っていない状態になり、開けば取れる
        take = self.client.post("/api/screen/take-over", json={},
                                headers={"X-Tool-Token": TOKEN})
        self.assertEqual(take.status_code, 200)
        self.assertIsNone(screen.holder())
        self.assertEqual(self._claim("2枚目").status_code, 200)

        # 前の画面は断られる
        res = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        self.assertEqual(res.status_code, 409)

    def test_閉じたらすぐ空く(self):
        """猶予を待たせない。**開き直しはよくある操作。**"""
        self.client.post("/api/alive", json={"screen": SCREEN, "leaving": True})
        self.assertIsNone(screen.holder())
        self.assertEqual(self._claim("べつの画面").status_code, 200)

    def test_心拍が途切れれば空く(self):
        """タブが落ちても、誰も入れないままにはしない。"""
        self.assertIsNotNone(screen.holder())
        screen._last_seen = time.time() - screen.GRACE_SEC - 1
        self.assertIsNone(screen.holder())
        self.assertEqual(self._claim("べつの画面").status_code, 200)

    def test_持ち主でない心拍で延命しない(self):
        """延命すると、引き継がれた古いタブが裏で生き続ける。"""
        body = self.client.post(
            "/api/alive", json={"screen": "べつの画面"}).get_json()
        self.assertFalse(body["holds"])
        self.assertEqual(screen.holder().screen_id, SCREEN)

    def test_帳票は画面を持たなくても出せる(self):
        """**印刷は別のタブで開く**(`window.open`)。

        そのタブは画面を持っていない。ここで断ると印刷できなくなる。
        盤面は触らず、DBから引いて組むだけ(トークンは要る)。
        """
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        _post(self.client, "/api/strands", {})
        _post(self.client, "/api/stack-max", {"value": "1"})
        _post(self.client, "/api/stack", {"key": "1-12"})
        _post(self.client, "/api/output", {"confirm": True})

        res = self.client.get(f"/report/L5160Z0/1?t={TOKEN}")
        self.assertEqual(res.status_code, 200)
        self.assertIn("梱包明細表", res.get_data(as_text=True))


class LotTest(unittest.TestCase):
    """ロットを開く (VBA `txtLotNo_Change`)。"""

    def setUp(self):
        self.client, self.conn = _client(self)
        _seed(self.conn)

    def test_開くとロット情報と割数が返る(self):
        res = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        self.assertEqual(res.status_code, 200)
        state = res.get_json()["state"]
        self.assertTrue(state["found"])
        self.assertEqual(state["lot"]["yoto_code"], "R192")
        self.assertEqual(state["lot"]["tate_wari"], 2)
        self.assertEqual(state["zen_kotei"], 24)          # 2 × 12
        self.assertEqual(state["jou_su"], 2)              # 縦割数が入る

    def test_表示書式がVBAと同じ(self):
        state = _post(self.client, "/api/lot",
                      {"lot_no": "L5160Z0"}).get_json()["state"]
        self.assertEqual(state["lot"]["seizou_thickness"], "1.985")
        self.assertEqual(state["lot"]["seizou_width"], "104.0")
        self.assertEqual(state["lot"]["odr_thickness"], "2.000")

    def test_受注と包装仕様NOが返る(self):
        state = _post(self.client, "/api/lot",
                      {"lot_no": "L5160Z0"}).get_json()["state"]
        self.assertEqual(state["orders"],
                         [{"order_no": "62015946", "spec_no": "1C0123",
                           "customer": "", "deliver_to": ""}])

    def test_桁数が違えば断る(self):
        res = _post(self.client, "/api/lot", {"lot_no": "L516"})
        self.assertEqual(res.status_code, 400)
        self.assertEqual(res.get_json()["error"]["code"], "bad_lot_no")

    def test_見つからなければ404(self):
        res = _post(self.client, "/api/lot", {"lot_no": "ZZZZZZZ"})
        self.assertEqual(res.status_code, 404)

    def test_小文字は大文字にそろえる(self):
        res = _post(self.client, "/api/lot", {"lot_no": "l5160z0"})
        self.assertEqual(res.status_code, 200)

    def test_未印刷のまま切り替えると確認する(self):
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        history_repo.set_printed(self.conn, False)
        _seed(self.conn, lot="AAA0001")
        body = _post(self.client, "/api/lot", {"lot_no": "AAA0001"}).get_json()
        self.assertEqual(body["confirm"], "unprinted")

    def test_了解すれば切り替わり前ロットが片付く(self):
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        history_repo.set_printed(self.conn, False)
        _seed(self.conn, lot="AAA0001")
        body = _post(self.client, "/api/lot",
                     {"lot_no": "AAA0001", "confirm": True}).get_json()
        self.assertTrue(body["ok"])
        self.assertEqual(body["state"]["lot_no"], "AAA0001")
        self.assertEqual(history_repo.current_seq(self.conn), 0)   # ResetSeq

    def test_LS4LOTに無ければ理由を出す(self):
        with self.conn:
            self.conn.execute("DELETE FROM 仕掛当工程")
        state = _post(self.client, "/api/lot",
                      {"lot_no": "L5160Z0"}).get_json()["state"]
        self.assertEqual(state["zen_kotei"], 0)
        self.assertIn("LS4LOT", state["lot"]["jou_su_note"])


class FlowTest(unittest.TestCase):
    """条番号 → 積み上げ → 出力 → 印刷 の一周。"""

    def setUp(self):
        self.client, self.conn = _client(self)
        _seed(self.conn)
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})

    def _to_strands(self):
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        return _post(self.client, "/api/strands", {}).get_json()

    def test_重量が無いと条番号を作れない(self):
        res = _post(self.client, "/api/strands", {})
        self.assertEqual(res.status_code, 400)
        self.assertEqual(res.get_json()["error"]["code"], "no_weight")

    def test_条番号が配分どおりに出る(self):
        state = self._to_strands()["state"]
        self.assertEqual(state["strands"], [12, 12])
        self.assertEqual(state["rows"][0][0]["key"], "1-12")   # 左端が最大
        self.assertEqual(state["rows"][0][-1]["key"], "1-1")

    def test_条番号を作ると途中経過が保存される(self):
        self._to_strands()
        snap = history_repo.load_snapshot(self.conn, "L5160Z0")
        self.assertIsNotNone(snap)
        self.assertEqual(snap.strands, [12, 12])
        self.assertEqual(snap.weights, [250, 248])

    def test_積み上げて出力できる(self):
        self._to_strands()
        _post(self.client, "/api/stack-max", {"value": "3"})
        for key in ("1-10", "1-5"):
            _post(self.client, "/api/stack", {"key": key})
        state = _post(self.client, "/api/stack", {"key": "2-8"}).get_json()["state"]
        self.assertTrue(state["complete"])
        self.assertEqual(state["slots"], ["1-10", "1-5", "2-8"])

        body = _post(self.client, "/api/output", {}).get_json()
        self.assertTrue(body["ok"])
        self.assertEqual(body["seq_no"], 1)
        # 出力後は積み上げだけ空。条番号は残る
        self.assertEqual(body["state"]["slots"], [None, None, None])
        self.assertEqual(body["state"]["strands"], [12, 12])
        self.assertEqual(body["state"]["outputs"][0]["name"], "L5160Z0-No1")

    def test_埋まるまで出力できない(self):
        self._to_strands()
        _post(self.client, "/api/stack-max", {"value": "2"})
        _post(self.client, "/api/stack", {"key": "1-1"})
        res = _post(self.client, "/api/output", {})
        self.assertEqual(res.status_code, 400)
        self.assertEqual(res.get_json()["error"]["code"], "incomplete")

    def test_使用済みのクリックは黙って無視する(self):
        self._to_strands()
        _post(self.client, "/api/stack-max", {"value": "2"})
        _post(self.client, "/api/stack", {"key": "1-1"})
        body = _post(self.client, "/api/stack", {"key": "1-1"}).get_json()
        self.assertTrue(body["ok"])            # 断りではなく無視
        self.assertEqual(body["ignored"], "used")

    def test_解除できる(self):
        self._to_strands()
        _post(self.client, "/api/stack-max", {"value": "2"})
        _post(self.client, "/api/stack", {"key": "1-1"})
        body = _post(self.client, "/api/unstack", {"slot_no": 1}).get_json()
        self.assertEqual(body["released"], "1-1")
        self.assertEqual(body["state"]["slots"], [None, None])

    def test_廃棄を付けると保存される(self):
        self._to_strands()
        body = _post(self.client, "/api/discard", {"key": "1-3"}).get_json()
        self.assertTrue(body["discarded"])
        snap = history_repo.load_snapshot(self.conn, "L5160Z0")
        self.assertEqual(snap.discarded, ["1-3"])

    def test_廃棄済みは積めない(self):
        self._to_strands()
        _post(self.client, "/api/stack-max", {"value": "2"})
        _post(self.client, "/api/discard", {"key": "1-3"})
        body = _post(self.client, "/api/stack", {"key": "1-3"}).get_json()
        self.assertEqual(body["ignored"], "discarded")

    def test_反転しても条番号は変わらない(self):
        before = self._to_strands()["state"]["rows"][0]
        after = _post(self.client, "/api/reverse", {}).get_json()["state"]["rows"][0]
        self.assertEqual([c["key"] for c in after],
                         list(reversed([c["key"] for c in before])))

    def test_帳票が出せる(self):
        self._to_strands()
        _post(self.client, "/api/stack-max", {"value": "2"})
        _post(self.client, "/api/stack", {"key": "1-10"})
        _post(self.client, "/api/stack", {"key": "2-8"})
        _post(self.client, "/api/output", {})

        res = self.client.get(f"/report/L5160Z0/1?t={TOKEN}")
        self.assertEqual(res.status_code, 200)
        html = res.get_data(as_text=True)
        self.assertIn("L5160Z0-No1  梱包明細表", html)
        self.assertIn("250.0Kg", html)
        self.assertIn("248.0Kg", html)
        self.assertEqual(html.count('class="r-data"'), config.SHEET_ROWS)

    def test_印刷すると印刷済みになる(self):
        history_repo.set_printed(self.conn, False)
        _post(self.client, "/api/printed", {})
        self.assertTrue(history_repo.is_printed(self.conn))

    def test_丈数を変えると条番号も重量もやり直し(self):
        self._to_strands()
        state = _post(self.client, "/api/jou-su", {"value": "3"}).get_json()["state"]
        self.assertEqual(state["jou_su"], 3)
        self.assertEqual(state["strands"], [])
        self.assertEqual(state["weights"], [0, 0, 0])

    def test_丈数の上限(self):
        res = _post(self.client, "/api/jou-su",
                    {"value": str(config.JOUSU_MAX + 1)})
        self.assertEqual(res.status_code, 400)

    def test_前工程実績数を直せる(self):
        body = _post(self.client, "/api/zen-kotei", {"value": "30"}).get_json()
        self.assertEqual(body["state"]["zen_kotei"], 30)

    def test_前工程実績数は数値だけ(self):
        res = _post(self.client, "/api/zen-kotei", {"value": "あ"})
        self.assertEqual(res.status_code, 400)


class PrintTest(unittest.TestCase):
    """帳票の経路。**明細1枚につき用紙1枚**(A4横の左半分、右は空ける)。"""

    def setUp(self):
        self.client, self.conn = _client(self)
        _seed(self.conn)
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        _post(self.client, "/api/strands", {})
        _post(self.client, "/api/stack-max", {"value": "2"})
        for keys in (("1-10", "2-8"), ("1-9", "2-7"), ("1-8", "2-6")):
            for key in keys:
                _post(self.client, "/api/stack", {"key": key})
            _post(self.client, "/api/output", {"confirm": True})

    def _get(self, path):
        return self.client.get(path, headers=HEADERS)

    def test_紙面に書き足しの欄と送り先が付く(self):
        """**印刷の前に手直しできる。** 送り先には開いた窓のトークンが付く。"""
        html = self.client.get(f"/report/L5160Z0/1?t={TOKEN}").get_data(as_text=True)
        self.assertIn('data-placeholder="サイズ"', html)
        self.assertIn('data-placeholder="LOTNO"', html)
        self.assertIn(f'"/report/L5160Z0/edits?t={TOKEN}"', html)

    def _row_id(self, no=1):
        from modules.packing_details.meisai import meisai_service
        return meisai_service.find_output(self.conn, "L5160Z0", no).row_id

    def test_書き足しを送ると刷り直しても出る(self):
        rid = self._row_id()
        res = self.client.post(f"/report/L5160Z0/edits?t={TOKEN}",
                               json={"edits": {f"{rid}.size": "1.985×104.0",
                                               f"{rid}.lotno": "L5160Z0"}})
        self.assertEqual(res.status_code, 200, res.get_data(as_text=True))
        html = self._get("/report/L5160Z0/1").get_data(as_text=True)
        self.assertIn("1.985×104.0", html)

    def test_書き足しは画面を持っていなくても送れる(self):
        """**紙面は別のタブ**で、画面を持っていない。断ると直せない。"""
        rid = self._row_id()
        screen.reset()                     # 誰も画面を持っていない
        res = self.client.post(f"/report/L5160Z0/edits?t={TOKEN}",
                               json={"edits": {f"{rid}.size": "x"}})
        self.assertEqual(res.status_code, 200)

    def test_書き足しにもトークンが要る(self):
        rid = self._row_id()
        res = self.client.post("/report/L5160Z0/edits",
                               json={"edits": {f"{rid}.size": "x"}})
        self.assertEqual(res.status_code, 401)

    def test_長すぎる書き足しは理由を返す(self):
        """紙面のスクリプトはこの文言をそのまま出す。"""
        rid = self._row_id()
        res = self.client.post(f"/report/L5160Z0/edits?t={TOKEN}",
                               json={"edits": {f"{rid}.lotno": "X" * 50}})
        self.assertEqual(res.status_code, 400)
        self.assertIn("LOTNO", res.get_json()["error"]["message"])

    def test_作り直された明細への書き足しは409(self):
        rid = self._row_id()
        # 出力済み修正(VER 0.13.23)で No1 を出し直す
        _post(self.client, "/api/edit-output", {"no": 1})
        _post(self.client, "/api/output", {"confirm": True})
        res = self.client.post(f"/report/L5160Z0/edits?t={TOKEN}",
                               json={"edits": {f"{rid}.size": "x"}})
        self.assertEqual(res.status_code, 409)

    def test_1枚は左半分で右半分は空ける(self):
        html = self._get("/report/L5160Z0/1").get_data(as_text=True)
        self.assertEqual(html.count('class="page"'), 1)
        self.assertEqual(html.count('class="half blank"'), 1)

    def test_まとめて刷ると1枚ずつ別の用紙(self):
        """**A4 1枚に2枚並べる運用はしていない**(現場に確認済み)。"""
        html = self._get("/report/L5160Z0?nos=1,2").get_data(as_text=True)
        self.assertEqual(html.count('class="page"'), 2)
        self.assertEqual(html.count('class="half blank"'), 2)
        self.assertIn("L5160Z0-No1", html)
        self.assertIn("L5160Z0-No2", html)

    def test_3枚なら用紙3枚(self):
        html = self._get("/report/L5160Z0?nos=1,2,3").get_data(as_text=True)
        self.assertEqual(html.count('class="page"'), 3)
        self.assertEqual(html.count('class="half blank"'), 3)

    def test_並びはNo順に固定(self):
        """押した順で刷ると、出てくる紙の順が押し方で変わってしまう。"""
        html = self._get("/report/L5160Z0?nos=3,1").get_data(as_text=True)
        self.assertLess(html.index("L5160Z0-No1"), html.index("L5160Z0-No3"))

    def test_重複した指定はまとめる(self):
        html = self._get("/report/L5160Z0?nos=1,1,2").get_data(as_text=True)
        self.assertEqual(html.count('class="meisai"'), 2)

    def test_無いNoを混ぜたら断る(self):
        res = self._get("/report/L5160Z0?nos=1,99")
        self.assertEqual(res.status_code, 404)

    def test_数値でないNoは断る(self):
        res = self._get("/report/L5160Z0?nos=1,あ")
        self.assertEqual(res.status_code, 400)

    def test_Noを指定しなければ断る(self):
        self.assertEqual(self._get("/report/L5160Z0?nos=").status_code, 400)

    def test_まとめ刷りもトークンが要る(self):
        self.assertEqual(self.client.get("/report/L5160Z0?nos=1").status_code,
                         401)


class RestoreTest(unittest.TestCase):
    """復元 (VBA `RestoreSnapshot`)。"""

    def setUp(self):
        self.client, self.conn = _client(self)
        _seed(self.conn)

    def test_開き直すと途中経過が戻る(self):
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        _post(self.client, "/api/strands", {})
        _post(self.client, "/api/stack-max", {"value": "2"})
        _post(self.client, "/api/stack", {"key": "1-10"})
        _post(self.client, "/api/discard", {"key": "1-3"})

        # プロセスが落ちた想定。セッションだけ捨てて開き直す
        session.clear()
        body = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"}).get_json()

        state = body["state"]
        self.assertTrue(state["restored"])
        self.assertIn("復元しました", body["message"])
        self.assertEqual(state["strands"], [12, 12])
        self.assertEqual(state["weights"], [250, 248])
        states = {c["key"]: c["state"] for row in state["rows"] for c in row}
        self.assertEqual(states["1-10"], "used")
        self.assertEqual(states["1-3"], "discarded")

    def test_復元後の再表示は確認を挟む(self):
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        _post(self.client, "/api/strands", {})
        session.clear()
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})

        body = _post(self.client, "/api/strands", {}).get_json()
        self.assertEqual(body["confirm"], "overwrite_restored")

        body = _post(self.client, "/api/strands", {"confirm": True}).get_json()
        self.assertTrue(body["ok"])
        self.assertFalse(body["state"]["restored"])


def _three_outputs(case):
    """No1〜No3 を画面の経路で出す。"""
    case.client, case.conn = _client(case)
    _seed(case.conn)
    _post(case.client, "/api/lot", {"lot_no": "L5160Z0"})
    _post(case.client, "/api/weights", {"weights": ["250", "248"]})
    _post(case.client, "/api/strands", {})
    _post(case.client, "/api/stack-max", {"value": "2"})
    for keys in (("1-10", "2-8"), ("1-9", "2-7"), ("1-8", "2-6")):
        for key in keys:
            _post(case.client, "/api/stack", {"key": key})
        _post(case.client, "/api/output", {"confirm": True})


def _coil_states(state):
    return {c["key"]: c["state"] for row in state["rows"] for c in row}


class RemovedTest(unittest.TestCase):
    """副番履歴クリアと No指定は外した(VER 0.13.23・現場の判断)。"""

    def setUp(self):
        _three_outputs(self)

    def test_副番履歴クリアの口は無い(self):
        res = _post(self.client, "/api/clear-history", {})
        self.assertIn(res.status_code, (404, 405))
        self.assertEqual(history_repo.fuban_count(self.conn), 6)

    def test_No指定を送っても次の連番で出る(self):
        """古い画面が `specify_no` を送ってきても、上書きにはしない。"""
        _post(self.client, "/api/stack", {"key": "1-1"})
        _post(self.client, "/api/stack", {"key": "2-1"})
        body = _post(self.client, "/api/output", {"confirm": True, "specify_no": 1}).get_json()
        self.assertEqual(body["seq_no"], 4)
        self.assertEqual(len(meisai_service.list_outputs(self.conn, "L5160Z0")), 4)

    def test_画面にも無い(self):
        html = self.client.get(f"/meisai?t={TOKEN}").get_data(as_text=True)
        for gone in ("btnClearHistory", "chkSpecify", "specifyNo", "副番履歴クリア"):
            self.assertNotIn(f'id="{gone}"' if gone.isascii() else f">{gone}<", html)
        for new in ('id="btnEditOutput"', 'id="btnClearAll"', 'id="btnEditCancel"'):
            self.assertIn(new, html)


class EditOutputWebTest(unittest.TestCase):
    """出力済み修正(VER 0.13.23)。出力済みに1つチェック → 「出力済み修正」。"""

    def setUp(self):
        _three_outputs(self)

    def test_修正に入るとそのNoが積み上げに戻り_ほかのNoは使用済み(self):
        body = _post(self.client, "/api/edit-output", {"no": 1}).get_json()
        state = body["state"]
        self.assertEqual(state["edit"], {"no": 1, "keys": ["1-10", "2-8"]})
        self.assertEqual(state["slots"], ["1-10", "2-8"])
        states = _coil_states(state)
        for key in ("1-9", "2-7", "1-8", "2-6"):
            self.assertEqual(states[key], "used", key)
        self.assertIn("No1を修正します", body["message"])

    def test_ほかのNoのコイルを押しても積まれない(self):
        _post(self.client, "/api/edit-output", {"no": 1})
        _post(self.client, "/api/unstack", {"slot_no": 2})
        body = _post(self.client, "/api/stack", {"key": "2-7"}).get_json()   # No2 のコイル
        self.assertEqual(body["ignored"], "used")
        self.assertEqual(body["state"]["slots"], ["1-10", None])

    def test_修正して出力すると同じNoで出し直し_積み上げは空(self):
        _post(self.client, "/api/edit-output", {"no": 1})
        _post(self.client, "/api/unstack", {"slot_no": 2})
        _post(self.client, "/api/stack", {"key": "2-1"})
        body = _post(self.client, "/api/output", {}).get_json()
        self.assertEqual(body["confirm"], "output")
        self.assertEqual(body["confirms"][0]["kind"], "edit")
        body = _post(self.client, "/api/output", {"confirm": True}).get_json()
        self.assertTrue(body["ok"])
        self.assertEqual(body["seq_no"], 1)
        state = body["state"]
        self.assertIsNone(state["edit"])
        self.assertEqual(state["slots"], [None, None])
        self.assertEqual(state["current_seq"], 3)
        self.assertEqual([o["no"] for o in state["outputs"]], [1, 2, 3])
        states = _coil_states(state)
        self.assertEqual(states["2-8"], "available")     # 外したコイルは空く
        self.assertEqual(states["2-1"], "used")
        self.assertEqual(meisai_service.find_output(self.conn, "L5160Z0", 1).keys,
                         ["1-10", "2-1"])
        snap = history_repo.load_snapshot(self.conn, "L5160Z0")
        self.assertIn("2-1", snap.used)
        self.assertNotIn("2-8", snap.used)

    def test_やめると前のまま(self):
        _post(self.client, "/api/edit-output", {"no": 2})
        _post(self.client, "/api/unstack", {"slot_no": 1})
        body = _post(self.client, "/api/edit-cancel", {}).get_json()
        self.assertIsNone(body["state"]["edit"])
        self.assertEqual(body["state"]["slots"], [None, None])
        self.assertEqual(_coil_states(body["state"])["1-9"], "used")
        self.assertEqual(meisai_service.find_output(self.conn, "L5160Z0", 2).keys,
                         ["1-9", "2-7"])

    def test_修正中は条番号の再表示と丈数の変更を断る(self):
        _post(self.client, "/api/edit-output", {"no": 1})
        for path, data in (("/api/strands", {"confirm": True}), ("/api/jou-su", {"value": "3"})):
            with self.subTest(path=path):
                res = _post(self.client, path, data)
                self.assertEqual(res.status_code, 409)
                self.assertEqual(res.get_json()["error"]["code"], "editing")

    def test_修正中にもう1枚は断る(self):
        _post(self.client, "/api/edit-output", {"no": 1})
        res = _post(self.client, "/api/edit-output", {"no": 2})
        self.assertEqual(res.status_code, 409)

    def test_無いNo_番号でないものは断る(self):
        self.assertEqual(_post(self.client, "/api/edit-output", {"no": 9}).status_code, 400)
        self.assertEqual(_post(self.client, "/api/edit-output", {"no": "x"}).status_code, 400)

    def test_修正中に落ちて開き直しても前の紙のコイルは使用済み(self):
        _post(self.client, "/api/edit-output", {"no": 1})
        _post(self.client, "/api/unstack", {"slot_no": 1})
        _post(self.client, "/api/discard", {"key": "1-1"})       # 途中経過が書かれる
        session.clear()
        body = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"}).get_json()
        states = _coil_states(body["state"])
        self.assertEqual(states["1-10"], "used")
        self.assertEqual(states["2-8"], "used")
        self.assertIsNone(body["state"]["edit"])

    def test_条番号を再表示しても出力済みのコイルは使用済みのまま(self):
        """No1〜No3 のコイルを No4 へまた積めないように(VER 0.13.23)。"""
        body = _post(self.client, "/api/strands", {}).get_json()
        states = _coil_states(body["state"])
        for key in ("1-10", "2-8", "1-9", "2-7", "1-8", "2-6"):
            self.assertEqual(states[key], "used", key)
        self.assertEqual(states["1-1"], "available")


class ClearAllTest(unittest.TestCase):
    """全クリア(VER 0.13.23)。履歴も含めて消して、初めからやり直す。"""

    def setUp(self):
        _three_outputs(self)
        _post(self.client, "/api/discard", {"key": "1-1"})

    def test_確かめるまでは何も消さない(self):
        body = _post(self.client, "/api/clear-all", {}).get_json()
        self.assertEqual(body["confirm"], "clear_all")
        self.assertIn("3枚", body["message"])
        self.assertEqual(len(meisai_service.list_outputs(self.conn, "L5160Z0")), 3)

    def test_出力済み_副番履歴_連番_途中経過が消える(self):
        body = _post(self.client, "/api/clear-all", {"confirm": True}).get_json()
        self.assertTrue(body["ok"])
        state = body["state"]
        self.assertEqual(state["outputs"], [])
        self.assertEqual(state["current_seq"], 0)
        self.assertEqual(state["fuban_count"], 0)
        self.assertEqual(state["strands"], [])          # 条番号表示からやり直す
        self.assertEqual(state["lot_no"], "L5160Z0")    # ロットは開いたまま
        self.assertTrue(state["printed"])
        self.assertIsNone(history_repo.load_snapshot(self.conn, "L5160Z0"))
        # 次の出力は No1 から。前のコイルも使える
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        state = _post(self.client, "/api/strands", {}).get_json()["state"]
        self.assertEqual(_coil_states(state)["1-10"], "available")
        self.assertEqual(_coil_states(state)["1-1"], "available")   # 廃棄も消える
        _post(self.client, "/api/stack-max", {"value": "1"})
        _post(self.client, "/api/stack", {"key": "1-10"})
        body = _post(self.client, "/api/output", {}).get_json()
        self.assertEqual(body["seq_no"], 1)                         # 重複の確認も出ない

    def test_明細の履歴は消さずに取り消しにする(self):
        _post(self.client, "/api/clear-all", {"confirm": True})
        states = [r[0] for r in self.conn.execute("SELECT 状態 FROM 明細履歴")]
        self.assertEqual(states, ["取り消し"] * 3)

    def test_修正中でも全クリアできる(self):
        _post(self.client, "/api/edit-output", {"no": 1})
        body = _post(self.client, "/api/clear-all", {"confirm": True}).get_json()
        self.assertIsNone(body["state"]["edit"])
        self.assertEqual(body["state"]["outputs"], [])

    def test_ロットを開いていなければ断る(self):
        _post(self.client, "/api/clear-lot", {})
        self.assertEqual(_post(self.client, "/api/clear-all", {"confirm": True}).status_code, 409)


class IrregularClearTest(unittest.TestCase):
    """全クリア → 出し直し → 出力済み修正 → もう一度全クリア の残り方(VER 0.13.24・現場の求め)。

    決まり: 全クリアは**その時点で効いている紙**(出力済みの一覧にある紙)だけを「取り消し」にする。
    すでに作り直された紙は「作り直し」のまま(そのときに起きたことがそれだから)。前の回で
    取り消した紙には触らない(二重に印を付けない)。
    """

    def setUp(self):
        self.client, self.conn = _client(self)
        _seed(self.conn)
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        _post(self.client, "/api/strands", {})

    def _out(self, keys):
        _post(self.client, "/api/stack-max", {"value": str(len(keys))})
        for key in keys:
            _post(self.client, "/api/stack", {"key": key})
        body = _post(self.client, "/api/output", {"confirm": True}).get_json()
        self.assertTrue(body.get("ok"), body)
        return body["seq_no"]

    def _edit(self, no, drop_slot, add_key):
        _post(self.client, "/api/edit-output", {"no": no})
        _post(self.client, "/api/unstack", {"slot_no": drop_slot})
        _post(self.client, "/api/stack", {"key": add_key})
        body = _post(self.client, "/api/output", {"confirm": True}).get_json()
        self.assertTrue(body.get("ok"), body)

    def _clear(self):
        body = _post(self.client, "/api/clear-all", {"confirm": True}).get_json()
        self.assertTrue(body.get("ok"), body)
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        _post(self.client, "/api/strands", {})

    def _papers(self):
        rows = self.conn.execute(
            "SELECT 送信ID, No, 状態, 変更回数, 本数 FROM 明細履歴 ORDER BY 出力日時, rowid").fetchall()
        coils = {}
        for r in self.conn.execute("SELECT 送信ID, 副番 FROM 明細履歴_副番 ORDER BY 順"):
            coils.setdefault(r[0], []).append(r[1])
        return [(r["No"], r["状態"], " ".join(coils.get(r["送信ID"], [])), r["変更回数"],
                 r["送信ID"]) for r in rows]

    def test_全クリアして直したものを修正し_もう一度全クリア(self):
        # 1回目: No1〜No3 を出して全クリア
        for keys in (("1-12", "2-12"), ("1-11", "2-11"), ("1-10", "2-10")):
            self._out(keys)
        self._clear()
        round1 = {p[4]: p[3] for p in self._papers()}
        # 2回目: No1・No2 を出し、No1 を2回修正して、もう一度全クリア
        self.assertEqual(self._out(("1-12", "2-12")), 1)      # 1回目のコイルも使える・No1から
        self.assertEqual(self._out(("1-11", "2-11")), 2)
        self._edit(1, 2, "2-9")                                 # 1-12 2-9
        self._edit(1, 1, "1-9")                                 # 1-9 2-9
        self._clear()
        # 3回目: No1 を出す
        self.assertEqual(self._out(("1-12", "2-12")), 1)

        papers = self._papers()
        self.assertEqual([(no, st, c) for no, st, c, _, _ in papers], [
            (1, "取り消し", "1-12 2-12"), (2, "取り消し", "1-11 2-11"), (3, "取り消し", "1-10 2-10"),
            (1, "作り直し", "1-12 2-12"), (2, "取り消し", "1-11 2-11"),
            (1, "作り直し", "1-12 2-9"), (1, "取り消し", "1-9 2-9"),
            (1, "有効", "1-12 2-12"),
        ])
        # 前の回で取り消した紙には、2回目の全クリアで触らない(変更回数が増えない)
        for sid, count in round1.items():
            self.assertEqual([p[3] for p in papers if p[4] == sid], [count])
        # 画面: 3回目の No1 だけが出力済み。連番は1
        self.assertEqual(history_repo.current_seq(self.conn), 1)
        self.assertEqual([o.seq_no for o in meisai_service.list_outputs(self.conn, "L5160Z0")], [1])

        # 集計(CSV): 数えるのは有効の1枚だけ。作り直し2・取り消し5
        from modules.packing_details.meisai import slip_history
        counts = {"有効": 0, "作り直し": 0, "取り消し": 0}
        for _, st, _, _, _ in papers:
            counts[st] += 1
        self.assertEqual(counts, {"有効": 1, "作り直し": 2, "取り消し": 5})
        self.assertEqual(slip_history.STATE_CANCELLED, "取り消し")

    def test_呼び出した紙面の断り書き(self):
        self._out(("1-12", "2-12"))
        self._edit(1, 2, "2-9")
        self._clear()
        ids = ",".join(p[4] for p in self._papers())
        res = self.client.get(f"/report/history?ids={ids}&t={TOKEN}", headers=HEADERS)
        self.assertEqual(res.status_code, 200, res.get_data(as_text=True)[:300])
        html = res.get_data(as_text=True)
        self.assertIn("このあと作り直されています（同じNoの新しい紙が履歴にあります）", html)
        self.assertIn("このあと全クリアで取り消されています", html)

    def test_ロットを切り替えて戻った紙は全クリアで取り消さない(self):
        """ロット切替で片付いた紙は、刷って使った紙(有効のまま)。戻って全クリアしても触らない。"""
        self._out(("1-12", "2-12"))
        _seed(self.conn, lot="L7110P0")
        _post(self.client, "/api/lot", {"lot_no": "L7110P0", "confirm": True})
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0", "confirm": True})
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        _post(self.client, "/api/strands", {})
        self._out(("1-11", "2-11"))
        self._clear()
        self.assertEqual([(no, st) for no, st, _, _, _ in self._papers()],
                         [(1, "有効"), (1, "取り消し")])

    def test_修正中に全クリア(self):
        self._out(("1-12", "2-12"))
        _post(self.client, "/api/edit-output", {"no": 1})
        _post(self.client, "/api/unstack", {"slot_no": 1})
        self._clear()
        self.assertEqual([(no, st) for no, st, _, _, _ in self._papers()], [(1, "取り消し")])
        state = _post(self.client, "/api/stack-max", {"value": "1"}).get_json()["state"]
        self.assertIsNone(state["edit"])


class JouSuLockTest(unittest.TestCase):
    """積み上げを始めたロットは丈数を変えられない(VER 0.13.24・現場の求め)。変えるときは全クリア。"""

    def setUp(self):
        self.client, self.conn = _client(self)
        _seed(self.conn)
        _post(self.client, "/api/lot", {"lot_no": "L5160Z0"})
        _post(self.client, "/api/weights", {"weights": ["250", "248"]})
        _post(self.client, "/api/strands", {})
        _post(self.client, "/api/stack-max", {"value": "2"})

    def _jou(self, value="3"):
        return _post(self.client, "/api/jou-su", {"value": value})

    def test_積む前なら変えられる(self):
        body = self._jou().get_json()
        self.assertTrue(body["ok"])
        self.assertFalse(body["state"]["jou_su_locked"])
        self.assertEqual(body["state"]["jou_su"], 3)

    def test_1本でも積んだら断る(self):
        state = _post(self.client, "/api/stack", {"key": "1-12"}).get_json()["state"]
        self.assertTrue(state["jou_su_locked"])
        res = self._jou()
        self.assertEqual(res.status_code, 409)
        self.assertEqual(res.get_json()["error"]["code"], "jou_su_locked")
        self.assertIn("全クリア", res.get_json()["error"]["message"])
        self.assertEqual(session.current().jou_su, 2)

    def test_積んだものを外しても_出力済みがあれば断る(self):
        _post(self.client, "/api/stack", {"key": "1-12"})
        _post(self.client, "/api/stack", {"key": "2-12"})
        _post(self.client, "/api/output", {"confirm": True})
        self.assertEqual(self._jou().status_code, 409)

    def test_開き直して途中経過から戻しても断る(self):
        _post(self.client, "/api/stack", {"key": "1-12"})
        _post(self.client, "/api/stack", {"key": "2-12"})
        _post(self.client, "/api/output", {"confirm": True})
        session.clear()
        state = _post(self.client, "/api/lot", {"lot_no": "L5160Z0"}).get_json()["state"]
        self.assertTrue(state["jou_su_locked"])
        self.assertEqual(self._jou().status_code, 409)

    def test_全クリアすれば変えられる(self):
        _post(self.client, "/api/stack", {"key": "1-12"})
        _post(self.client, "/api/stack", {"key": "2-12"})
        _post(self.client, "/api/output", {"confirm": True})
        state = _post(self.client, "/api/clear-all", {"confirm": True}).get_json()["state"]
        self.assertFalse(state["jou_su_locked"])
        self.assertTrue(self._jou().get_json()["ok"])

    def test_廃棄だけなら変えられる(self):
        _post(self.client, "/api/discard", {"key": "1-1"})
        self.assertTrue(self._jou().get_json()["ok"])


class SettingsTest(unittest.TestCase):
    """設定。"""

    def setUp(self):
        self.client, self.conn = _client(self)

    def test_設定が読める(self):
        body = self.client.get("/api/settings", headers=HEADERS).get_json()
        self.assertIn("lot_db_dir", body)
        self.assertEqual(len(body["stamps"]), len(config.ALL_SOURCE_FILES))

    def test_一覧に無い設定は変えられない(self):
        res = _post(self.client, "/api/settings",
                    {"key": "admin_password", "value": "x"})
        self.assertEqual(res.status_code, 400)
        self.assertEqual(res.get_json()["error"]["code"], "not_listed")


class ManualTest(unittest.TestCase):
    """説明書 (VBA `cmdManual_Click`)。"""

    def setUp(self):
        self.client, self.conn = _client(self)

    def test_トークン無しで開ける(self):
        """**業務データを含まない。**

        トークンを要求すると、切れた画面から説明書に辿り着けなくなる
        ── 困っている人ほど読めない、という形になる。
        """
        res = self.client.get("/docs")
        self.assertEqual(res.status_code, 200)
        html = res.get_data(as_text=True)
        self.assertIn("コイル梱包明細打ち出しシステム", html)

    def test_見出しと表が組まれる(self):
        html = self.client.get("/docs").get_data(as_text=True)
        self.assertIn("<h2>", html)
        self.assertIn("<table>", html)
        self.assertIn("<blockquote>", html)

    def test_中身がそのまま出ない(self):
        """Markdown の記号が生で出ていないか。"""
        html = self.client.get("/docs").get_data(as_text=True)
        body = html.split("<body>", 1)[1]
        self.assertNotIn("## ", body)
        self.assertNotIn("**", body)

    def test_変更履歴も同じ経路で読める(self):
        """**版を見せるなら「何が変わったか」まで辿れないと半分。**"""
        res = self.client.get("/docs/変更履歴")
        self.assertEqual(res.status_code, 200)
        self.assertIn("変更履歴", res.get_data(as_text=True))

    def test_変更履歴も記号が生で出ない(self):
        """版のバッジから開く。説明書だけ見ていて、こちらが崩れていた。"""
        import re
        html = self.client.get("/docs/変更履歴").get_data(as_text=True)
        body = html.split("<body>", 1)[1]
        # コードの中の `**` は記号そのものを見せている(記法ではない)
        body = re.sub(r"<pre>.*?</pre>|<code>.*?</code>", "", body, flags=re.S)
        self.assertNotIn("**", body)

    def test_コードの中の記号は強調にしない(self):
        """`` `**` `` を離れた `**` と組にすると、間が丸ごと太字になる。"""
        from modules.packing_details.app.routes.settings import _inline
        self.assertEqual(_inline("記号 `**` のまま。**強調**"),
                         "記号 <code>**</code> のまま。<strong>強調</strong>")
        self.assertEqual(_inline("**強調の中の `コード`**"),
                         "<strong>強調の中の <code>コード</code></strong>")

    def test_箇条書きの続きの行は同じ項目(self):
        """字下げした続きの行を段落にしない。以前は

            - 前の画面は、次に操作した時点で「…」と出て
              止まります

        の2行目が箇条書きの外へ出て、別の段落になっていた。
        """
        from modules.packing_details.app.routes.settings import _render_markdown
        html = _render_markdown("- 前の画面は、**次に操作した\n  時点で**止まります\n- 次")
        self.assertEqual(
            html, "<ul><li>前の画面は、<strong>次に操作した 時点で</strong>止まります</li>"
                  "<li>次</li></ul>")

    def test_字下げしていない行で箇条書きは終わる(self):
        from modules.packing_details.app.routes.settings import _render_markdown
        self.assertEqual(_render_markdown("- 項目\n段落"),
                         "<ul><li>項目</li></ul><p>段落</p>")

    def test_引用の続きの行は同じ段落(self):
        """`>` だけの行で段落を分ける。以前は1行ごとに段落になり、
        行をまたいだ強調が `**` のまま出ていた。
        """
        from modules.packing_details.app.routes.settings import _render_markdown
        html = _render_markdown("> **見出し**\n>\n> 本文の**1行目\n> 2行目**です")
        self.assertEqual(
            html, "<blockquote><p><strong>見出し</strong></p>"
                  "<p>本文の<strong>1行目 2行目</strong>です</p></blockquote>")

    def test_一覧に無い名前は断る(self):
        """受け取った文字列でパスを組むと、アプリの外まで読める。"""
        for name in ("秘密", "../config/app.json", "..%2fapp.json"):
            with self.subTest(name=name):
                res = self.client.get(f"/docs/{name}")
                self.assertEqual(res.status_code, 404)
                self.assertNotIn("app_id", res.get_data(as_text=True))


class StackSpinTest(unittest.TestCase):
    """積み条数にも −/＋ を置く(現場の指摘。梱包明細 0.13.5)。"""

    def test_積み条数の横にスピンボタン(self):
        client, _ = _client(self)
        html = client.get(f"/meisai?t={TOKEN}").data.decode("utf-8")
        at = html.index('id="stackMax"')
        near = html[at:at + 600]
        self.assertIn('data-stack="-1"', near)
        self.assertIn('data-stack="1"', near)
        self.assertIn(f'data-max="{config.STACK_LIMIT}"', near)


if __name__ == "__main__":
    unittest.main()
