"""エラーの後追い(統合 1.0.12。現場の指摘)。

    エラー等の後追いが現状できないと感じている。ログを残し、なぜなぜで分析できるように
    しておいてほしい。ログ出力は設定でパス指定できるようにする。

- 止まった処理は**エラーの記録**(1件1ファイル)を作り、画面へ**エラー番号**を返す
- 記録はなぜなぜ分析の順(現象・直接の原因・エラーまでの流れ・状態・記入欄)
- 要求ごとに印(`[要求 R-…]`)を付け、断り・失敗は画面に出した文言も残す
- 画面(ブラウザ)のエラーも記録する(`static/js/error_report.js`)
- ログの出力先・残す日数を設定できる(上の帯の「ログ」。このPCに保存)
"""
from __future__ import annotations

import logging
import os
import re
import shutil
import tempfile
import unittest
from datetime import date, timedelta
from pathlib import Path
from unittest import mock

from flask import Blueprint, jsonify

from .test_app import TOKEN, _make_app

H = {"X-Tool-Token": TOKEN}
ID = re.compile(r"E\d{8}-\d{6}-[0-9A-F]{4}")


def _client(case, *, extra=True):
    app = _make_app(case)
    if extra:
        bp = Blueprint("trail_test", __name__)

        @bp.post("/material/api/boom")
        def boom():
            logging.getLogger("coil_tool.calc").info("計算を始めます: ロット A123456")
            raise KeyError("包装仕様")

        @bp.get("/details/boom")
        def boom_page():
            raise ValueError("<b>こわれた</b>")

        @bp.post("/details/api/refuse")
        def refuse():
            return jsonify({"ok": False, "message": "ロットが見つかりません"}), 409

        app.register_blueprint(bp)
    return app.test_client()


class ServerErrorTest(unittest.TestCase):
    def setUp(self):
        self.client = _client(self)

    def test_api_failure_returns_an_error_number(self):
        res = self.client.post("/material/api/boom?t=" + TOKEN,
                               json={"password": "himitsu", "合言葉": "aikotoba", "lot": "A123456"})
        self.assertEqual(res.status_code, 500)
        body = res.get_json()
        self.assertFalse(body["ok"])
        self.assertRegex(body["error_id"], ID)
        self.assertIn(body["error_id"], body["message"])
        self.assertIn(body["error_id"], body["error"]["message"], "梱包明細の画面はこちらを読む")

    def test_the_record_follows_the_why_why_order(self):
        from common import incidents
        body = self.client.post("/material/api/boom?t=" + TOKEN,
                                json={"password": "himitsu", "lot": "A123456"}).get_json()
        text = incidents.read(body["error_id"])
        for head in ("## 1. 何が起きたか(現象)", "## 2. 直接の原因", "## 3. エラーまでの流れ",
                     "## 4. そのときの状態", "## 5. なぜなぜ分析(記入用)"):
            self.assertIn(head, text)
        for n in range(1, 6):
            self.assertIn(f"- なぜ{n}: ", text)
        self.assertIn("- 再発防止(仕組みで防ぐこと): ", text)
        self.assertIn("KeyError: '包装仕様'", text)
        self.assertIn("- 機能: 資材計算", text)
        self.assertIn(body["error_id"], text.split("画面に出した文言: ")[1].split("\n")[0])
        # 伏せる: パスワード・トークン
        self.assertNotIn("himitsu", text)
        self.assertNotIn(TOKEN, text)
        self.assertIn('"password": "***"', text)
        # この操作の中の行には印(>>)
        self.assertRegex(text, r">> .*計算を始めます: ロット A123456 \[要求 R-[0-9A-F]{6}\]")

    def test_page_failure_shows_the_number_and_escapes(self):
        res = self.client.get("/details/boom")
        self.assertEqual(res.status_code, 500)
        html = res.get_data(as_text=True)
        self.assertRegex(html, "エラー番号 " + ID.pattern)
        self.assertNotIn("<b>こわれた</b>", html)

    def test_http_errors_are_left_alone(self):
        res = self.client.get("/nope")
        self.assertEqual(res.status_code, 404)
        self.assertEqual(res.get_json()["error"]["code"], "not_found")


class RequestTrailTest(unittest.TestCase):
    def setUp(self):
        self.client = _client(self)

    def test_refusals_keep_the_message_the_user_saw(self):
        with self.assertLogs("coil_packing_tools.app", level="WARNING") as cm:
            self.client.post("/details/api/refuse", json={})
        self.assertTrue(any("断り・失敗 409 POST /details/api/refuse: ロットが見つかりません" in l
                            for l in cm.output), cm.output)

    def test_each_line_carries_the_request_mark(self):
        from common import logging_utils
        self.client.post("/details/api/refuse", json={})
        lines = [l for l in logging_utils.recent_lines(50) if "/details/api/refuse" in l]
        self.assertTrue(lines)
        self.assertTrue(all(re.search(r"\[要求 R-[0-9A-F]{6}\]$", l) for l in lines), lines)

    def test_heartbeats_stay_quiet(self):
        with self.assertLogs("coil_packing_tools.app", level="DEBUG") as cm:
            self.client.get("/details/api/health")
            self.client.post("/details/api/refuse", json={})
        health = [l for l in cm.output if "/details/api/health" in l]
        self.assertTrue(health and all(l.startswith("DEBUG:") for l in health), health)


class ClientLogTest(unittest.TestCase):
    def setUp(self):
        import app as app_module
        app_module._client_log_times.clear()
        self.client = _client(self, extra=False)

    def post(self, body, headers=H):
        return self.client.post("/api/client-log", json=body, headers=headers)

    def test_needs_the_token(self):
        self.assertEqual(self.post({"kind": "error"}, headers={}).status_code, 403)

    def test_screen_errors_become_records_and_repeat_into_one(self):
        from common import incidents
        body = {"kind": "error", "message": "x is not defined", "module": "material",
                "source": "/material/static/js/views/calc.js", "line": 12, "col": 3,
                "stack": "ReferenceError: x is not defined\n at calc.js:12", "page": "/material/?t=" + TOKEN}
        first = self.post(body).get_json()
        again = self.post(body).get_json()
        self.assertRegex(first["error_id"], ID)
        self.assertEqual(first["error_id"], again["error_id"], "同じものは1件にまとめる")
        text = incidents.read(first["error_id"])
        self.assertIn("画面(ブラウザ)のエラー", text)
        self.assertIn("- 機能: 資材計算", text)
        self.assertIn("/material/static/js/views/calc.js:12:3", text)
        self.assertIn("同じエラーがもう一度", text)
        self.assertNotIn(TOKEN, text)

    def test_unreached_requests_are_a_line_not_a_record(self):
        with self.assertLogs("coil_packing_tools.app", level="WARNING") as cm:
            res = self.post({"kind": "offline", "source": "/details/api/lot", "count": 3,
                             "message": "10:00 〜 10:02"})
        self.assertEqual(res.get_json(), {"ok": True})
        self.assertIn("サーバに届かなかった通信 /details/api/lot(3 回", "\n".join(cm.output))

    def test_limits(self):
        import app as app_module
        big = {"kind": "error", "message": "x" * (app_module.CLIENT_LOG_MAX_BYTES + 10)}
        self.assertEqual(self.post(big).status_code, 413)
        with mock.patch.object(app_module, "CLIENT_LOG_PER_MIN", 2):
            codes = [self.post({"kind": "offline", "source": f"/a{i}"}).status_code
                     for i in range(3)]
        self.assertEqual(codes, [200, 200, 429])


class InjectionTest(unittest.TestCase):
    def setUp(self):
        self.client = _client(self, extra=False)

    def test_reporter_is_in_every_screen(self):
        for path in (f"/details/meisai?t={TOKEN}", f"/material/?t={TOKEN}", "/pena/",
                     "/pena/labels/print?ob=5", "/"):
            with self.subTest(path=path):
                html = self.client.get(path, follow_redirects=True).get_data(as_text=True)
                self.assertEqual(html.count("js/error_report.js"), 1, path)
                self.assertIn(f'data-token="{TOKEN}"', html)

    def test_fragments_are_left_alone(self):
        html = self.client.get("/pena/tare?pane=1").get_data(as_text=True)
        self.assertNotIn("error_report.js", html)

    def test_log_page_and_apis(self):
        page = self.client.get("/log?embed=1").get_data(as_text=True)
        self.assertIn("エラーの記録", page)
        self.assertIn("出力先の設定", page)
        for path in ("/api/log/status", "/api/log/incidents", "/api/log/recent"):
            with self.subTest(path=path):
                self.assertEqual(self.client.get(path).status_code, 403)
                self.assertEqual(self.client.get(path, headers=H).status_code, 200)
        self.assertEqual(self.client.get("/api/log/incidents/../../etc", headers=H).status_code,
                         404)

    def test_shell_has_the_log_button(self):
        html = self.client.get("/?go=1").get_data(as_text=True)
        self.assertIn('id="logBtn"', html)
        self.assertIn('id="logDialog"', html)


class LogDirSettingTest(unittest.TestCase):
    """ログの出力先を設定で変える(このPCに保存。保存するとすぐ切り替わる)。"""

    def setUp(self):
        from common import logging_utils
        self.tmp = Path(tempfile.mkdtemp(prefix="cpt-logdir-"))
        env = {k: v for k, v in os.environ.items() if k != logging_utils.ENV_LOG_DIR}
        env["COIL_PACKING_TOOLS_SETTINGS_PATH"] = str(self.tmp / "settings.json")
        patcher = mock.patch.dict(os.environ, env, clear=True)
        patcher.start()
        self.addCleanup(patcher.stop)
        self.addCleanup(logging_utils.apply_settings)        # 元の場所へ戻す(後の試験のため)
        logging_utils.apply_settings()
        self.client = _client(self, extra=False)

    def send(self, **body):
        return self.client.post("/api/log/settings", json=body, headers=H)

    def test_save_switches_at_once_into_a_pc_folder(self):
        from common import logging_utils
        target = self.tmp / "share"
        res = self.send(action="save", log_dir=str(target), keep_days=90)
        self.assertEqual(res.status_code, 200, res.get_json())
        body = res.get_json()
        # 統合 1.2.53: <出力パス>\<ツールID>(日報複合ツールの決まり。PC名はファイル名に入る)
        self.assertEqual(body["dir"], str(target / "coilpacking"))
        self.assertEqual(body["keep_days"], 90)
        logging.getLogger("coil_packing_tools.x").info("出力先を変えたあとの行")
        today = target / "coilpacking" / f"coilpacking_{logging_utils.pc_name()}_{date.today():%Y%m%d}.log"
        self.assertIn("出力先を変えたあとの行", today.read_text(encoding="utf-8"))
        status = self.client.get("/api/log/status", headers=H).get_json()
        self.assertEqual(status["configured"], str(target))
        # 既定に戻す
        back = self.send(action="reset").get_json()
        self.assertEqual(back["dir"], str(logging_utils.default_log_dir()))
        self.assertEqual(back["configured"], "")

    def test_check_does_not_save(self):
        res = self.send(action="check", log_dir=str(self.tmp / "x")).get_json()
        self.assertTrue(res["ok"])
        status = self.client.get("/api/log/status", headers=H).get_json()
        self.assertEqual(status["configured"], "")

    def test_unwritable_place_is_refused(self):
        blocker = self.tmp / "file.txt"
        blocker.write_text("", encoding="utf-8")
        res = self.send(action="save", log_dir=str(blocker))
        self.assertEqual(res.status_code, 400)
        self.assertIn("書けません", res.get_json()["message"])

    def test_keep_days_range(self):
        self.assertEqual(self.send(action="save", log_dir="", keep_days=3).status_code, 400)
        self.assertEqual(self.send(action="save", log_dir="", keep_days="あ").status_code, 400)

    def test_a_lost_share_falls_back_to_this_pc(self):
        """出力先に書けなくなったら、このPCの既定の場所へ切り替えて書き続ける(黙って捨てない)。"""
        from common import logging_utils
        target = self.tmp / "share"
        self.send(action="save", log_dir=str(target))
        handler = logging_utils.DailyFileHandler(date.today(), encoding="utf-8", delay=True)
        handler.setFormatter(logging.Formatter(logging_utils.LINE_FORMAT))
        record = logging.LogRecord("x", logging.INFO, __file__, 0, "共有が切れた直後の行", (), None)
        record.ctx = ""
        handler.handleError(record)
        handler.close()
        self.assertIn("書けなくなりました", logging_utils.log_problem())
        fallback = logging_utils.default_log_dir() / f"{logging_utils.file_prefix()}_{date.today():%Y%m%d}.log"
        text = fallback.read_text(encoding="utf-8")
        self.assertIn("共有が切れた直後の行", text)
        self.assertIn("ログの出力先に書けなくなりました", text)


class LogPlaceRuleTest(unittest.TestCase):
    """ログの置き場所と名前の決まり(統合 1.2.53。日報複合ツール all-tools に合わせる)

    1. <出力パス>\\<ツールID>\\ に書く(出力パスの直下に置かない。ツールIDは半角英小文字)
    2. ファイル名は <ツールID>_ で始め、PC名を必ず入れる(coilpacking_<PC>_YYYYMMDD.log・
       incidents\\coilpacking_なぜなぜ_<PC>_E….md)
    3. 出力パスを決めていない・届かないときは、この端末の logs\\coilpacking\\ に同じ形で書く
    4. 前の場所・前の名前のログは移さない・もう書かない。読む・片付けるときは前の場所の自分のログも見る
    5. ほかのツールのファイル・フォルダは読まない・消さない。消すのは自分のツールの・この端末の物だけ
    """

    PC = "LINE-PC2"

    def setUp(self):
        from common import logging_utils
        self.lu = logging_utils
        self.tmp = Path(tempfile.mkdtemp(prefix="cpt-rule-"))
        self.addCleanup(shutil.rmtree, self.tmp, True)
        self.share = self.tmp / "share"
        self.cur = self.share / "coilpacking"
        self.local = self.tmp / "local_logs"
        for patcher in (mock.patch.dict(os.environ, {"COMPUTERNAME": self.PC}),
                        mock.patch.object(logging_utils, "log_dir", return_value=self.cur),
                        mock.patch.object(logging_utils, "local_logs_root", return_value=self.local),
                        mock.patch.object(logging_utils, "_env_base", return_value=None),
                        mock.patch.object(logging_utils, "configured_log_dir", return_value=str(self.share))):
            patcher.start()
            self.addCleanup(patcher.stop)

    def put(self, path: Path, text="x"):
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_text(text, encoding="utf-8")
        return path

    # 1・2
    def test_1_2_place_and_names(self):
        from common import incidents
        self.assertEqual(self.lu.tool_id(), "coilpacking")
        self.assertRegex(self.lu.tool_id(), r"^[a-z][a-z0-9]*$")
        self.assertEqual(self.lu.target_for(str(self.share)), self.share / "coilpacking")
        self.lu.log_dir.return_value = self.lu.target_for(str(self.share))
        self.assertEqual(self.lu.log_path_for(date(2026, 10, 24)),
                         self.share / "coilpacking" / "coilpacking_LINE-PC2_20261024.log")
        self.assertEqual(incidents.file_name("E20261024-101010-ABCD"),
                         "coilpacking_なぜなぜ_LINE-PC2_E20261024-101010-ABCD.md")
        self.assertEqual(incidents.folder(), self.share / "coilpacking" / "incidents")
        self.assertEqual(self.lu.default_log_dir(), self.local / "coilpacking", "3. この端末でも同じ形")

    def test_tool_id_must_be_lowercase_ascii(self):
        for bad in ("CoilPacking", "コイル", "coil-packing", "1coil", ""):
            with self.subTest(bad=bad), mock.patch.object(self.lu.app_config, "load",
                                                          return_value={"tool_id": bad}):
                self.assertEqual(self.lu.tool_id(), "coilpacking")
        with mock.patch.object(self.lu.app_config, "load", return_value={"tool_id": "coil2"}):
            self.assertEqual(self.lu.tool_id(), "coil2")

    def test_env_override_also_gets_the_tool_folder(self):
        with mock.patch.object(self.lu, "_env_base", return_value=self.tmp / "env"):
            self.assertEqual(self.lu.resolve_log_dir(), (self.tmp / "env" / "coilpacking", ""))

    # 4: 前の場所の自分の記録も読む・新しくは書かない
    def test_4_new_records_go_to_the_new_place_and_old_ones_are_read(self):
        from common import incidents
        try:
            raise ValueError("新しい決まりの試験")
        except ValueError as exc:
            eid = incidents.record("server", "新しい決まり", exc=exc)
        self.assertEqual([p.name for p in (self.cur / "incidents").iterdir()], [incidents.file_name(eid)])
        old_150 = self.put(self.share / self.PC / "incidents" / "E20261001-101010-AAAA.md", "- 要約: 1.2.50\n")
        old_151 = self.put(self.share / "コイル梱包ツール" / self.PC / "incidents"
                           / f"コイル梱包ツール_{self.PC}_E20261002-101010-BBBB.md", "- 要約: 1.2.51\n")
        old_local = self.put(self.local / "incidents" / "E20261003-101010-CCCC.md", "- 要約: 端末\n")
        # ほかのツール・ほかの PC の前の物は読まない
        self.put(self.share / "nippou" / "incidents" / "E20261004-101010-DDDD.md")
        self.put(self.share / "コイル梱包ツール" / self.PC / "incidents"
                 / "コイル梱包ツール_OTHER-PC_E20261005-101010-EEEE.md")
        ids = [x["id"] for x in incidents.list_recent()]
        self.assertEqual(ids, [eid, "E20261003-101010-CCCC", "E20261002-101010-BBBB", "E20261001-101010-AAAA"])
        self.assertIn("1.2.50", incidents.read("E20261001-101010-AAAA"))
        self.assertEqual(incidents.path_for("E20261002-101010-BBBB"), old_151)
        self.assertIsNone(incidents.read("E20261004-101010-DDDD"))
        # 同じエラーの「もう一度」は前の場所に書かない
        incidents._append_repeat("E20261001-101010-AAAA")
        self.assertEqual(old_150.read_text(encoding="utf-8"), "- 要約: 1.2.50\n")
        self.assertTrue(old_local.exists())

    def test_4_today_view_reads_todays_old_file_too(self):
        today = date.today()
        old = self.put(self.share / self.PC / f"coil_packing_tools_{today:%Y%m%d}.log", "朝の行\n")
        self.put(self.local / f"コイル梱包ツール_{self.PC}_{today:%Y%m%d}.log", "端末の行\n")
        self.put(self.local / f"コイル梱包ツール_OTHER-PC_{today:%Y%m%d}.log", "ほかの PC\n")
        got = self.lu.legacy_log_paths_for(today)
        self.assertIn(old, got)
        self.assertEqual(len(got), 2)

    # 4・5: 片付けるのは自分のツールの・この端末の物だけ(前の場所の自分の物も)
    def test_4_5_cleanup_only_this_tool_and_this_pc(self):
        from common import incidents
        today = date(2026, 10, 24)
        old = today - timedelta(days=200)
        mine_old = self.put(self.cur / f"coilpacking_{self.PC}_{old:%Y%m%d}.log")
        mine_new = self.put(self.cur / f"coilpacking_{self.PC}_{today:%Y%m%d}.log")
        other_pc = self.put(self.cur / f"coilpacking_OTHER-PC_{old:%Y%m%d}.log")
        kind_named = self.put(self.cur / f"coilpacking_記録_{self.PC}_{old:%Y%m%d}.log")
        memo = self.put(self.cur / "人が置いたメモ.log")
        other_tool = self.put(self.share / "nippou" / f"nippou_{self.PC}_{old:%Y%m%d}.log")
        root_other = self.put(self.share / f"dbviewer_{self.PC}_{old:%Y%m%d}.log")
        leg150 = self.put(self.share / self.PC / f"coil_packing_tools_{old:%Y%m%d}.log")
        leg150_other = self.put(self.share / self.PC / f"nippou_{old:%Y%m%d}.log")
        leg151 = self.put(self.share / "コイル梱包ツール" / self.PC / f"コイル梱包ツール_{self.PC}_{old:%Y%m%d}.log")
        leg_local = self.put(self.local / f"coil_packing_tools_{old:%Y%m%d}.log")
        leg_local_new = self.put(self.local / f"coil_packing_tools_{today:%Y%m%d}.log")
        inc_mine = self.put(self.cur / "incidents" / f"coilpacking_なぜなぜ_{self.PC}_E{old:%Y%m%d}-101010-ABCD.md")
        inc_other = self.put(self.cur / "incidents" / f"coilpacking_なぜなぜ_OTHER-PC_E{old:%Y%m%d}-101010-ABCD.md")
        inc_leg = self.put(self.share / self.PC / "incidents" / f"E{old:%Y%m%d}-101010-ABCD.md")
        inc_memo = self.put(self.local / "incidents" / "メモ.md")
        removed = self.lu.cleanup_old(180, today=today)
        removed_inc = incidents.cleanup_old(180, today=today)
        gone = [mine_old, leg150, leg151, leg_local, inc_mine, inc_leg]
        kept = [mine_new, other_pc, kind_named, memo, other_tool, root_other, leg150_other, leg_local_new,
                inc_other, inc_memo]
        for p in gone:
            self.assertFalse(p.exists(), p)
        for p in kept:
            self.assertTrue(p.exists(), p)
        self.assertEqual(len(removed) + len(removed_inc), len(gone))


class PenaIncidentTest(unittest.TestCase):
    def test_pena_api_failure_has_a_number_and_a_record(self):
        from common import incidents
        from modules.packing_pena_label.app.routes.api import ApiRoutes
        with mock.patch.object(ApiRoutes, "_materials", side_effect=RuntimeError("資材が読めない")):
            client = _make_app(self).test_client()
            res = client.post("/pena/api/materials", json={}, headers=H)
        self.assertEqual(res.status_code, 500)
        body = res.get_json()
        self.assertRegex(body["errorId"], ID)
        self.assertIn(f"エラー番号 {body['errorId']}", body["message"])
        text = incidents.read(body["errorId"])
        self.assertIn("RuntimeError: 資材が読めない", text)
        self.assertIn("- 機能: ペナラベル", text)


if __name__ == "__main__":
    unittest.main()
