"""出力・副番の重複・連番・スナップショット (RULE-06〜08)"""
from __future__ import annotations

import unittest

from modules.packing_details.meisai import (config, history_repo, meisai_service, report,
                                           slip_history, strand_service)
from modules.packing_details.meisai.history_repo import Snapshot
from modules.packing_details.meisai.meisai_service import (CONFIRM_DUPLICATE, CONFIRM_EDIT,
                                   REFUSE_INCOMPLETE, REFUSE_NO_WEIGHT, Output)
from modules.packing_details.meisai.strand_service import Board, RefusedError

from . import _db


def board(lot="L5160Z0", zen=24, jou=2, stack=6, weights=(250, 248)) -> Board:
    b = Board(lot_no=lot, zen_kotei=zen, jou_su=jou, weights=list(weights))
    b.build_strands()
    b.set_stack_max(stack)
    return b


def filled(lot="L5160Z0", keys=("1-10", "1-5", "2-8", "2-11", "2-12", "2-9")) -> Board:
    b = board(lot=lot, stack=len(keys))
    for key in keys:
        b.stack(key)
    return b


class WeightTest(unittest.TestCase):
    """RULE-06: 重量の読み取り。"""

    def test_整数として読む(self):
        # 入力欄が数字以外を弾く(VBA clsWeightCtrl)ので必ず整数
        self.assertEqual(meisai_service.read_weights(["250", "248"], 2), [250, 248])

    def test_未入力は断る(self):
        for values in ([""], ["250", ""], ["250"]):
            with self.subTest(values=values), self.assertRaises(RefusedError) as cm:
                meisai_service.read_weights(values, 2)
            self.assertEqual(cm.exception.reason, REFUSE_NO_WEIGHT)

    def test_数値でなければ断る(self):
        with self.assertRaises(RefusedError) as cm:
            meisai_service.read_weights(["abc", "248"], 2)
        self.assertEqual(cm.exception.reason, REFUSE_NO_WEIGHT)
        self.assertIn("丈1", cm.exception.message)

    def test_小数は黙って切り捨てずに断る(self):
        """統合 1.2.6: 日本語入力から入った 2502.7 を 2502 にして紙・履歴へ書いていた。"""
        for bad in ("2502.7", "2502.5", "inf", "nan"):
            with self.subTest(bad=bad), self.assertRaises(RefusedError) as cm:
                meisai_service.read_weights([bad, "248"], 2)
            self.assertEqual(cm.exception.reason, REFUSE_NO_WEIGHT)
            self.assertIn("整数", cm.exception.message)
        # 同じ数なら通す(小数点以下が 0)
        self.assertEqual(meisai_service.read_weights(["250.0", "248"], 2), [250, 248])

    def test_余分な欄は無視する(self):
        self.assertEqual(meisai_service.read_weights(["250", "248", "999"], 2),
                         [250, 248])


class OutputTest(unittest.TestCase):
    """出力そのもの。"""

    def setUp(self):
        self.conn = _db.case_db(self)

    def test_積み上げが埋まっていなければ断る(self):
        result = meisai_service.output(self.conn, board())
        self.assertFalse(result.ok)
        self.assertEqual(result.refuse, REFUSE_INCOMPLETE)

    def test_出力すると連番が1から始まる(self):
        result = meisai_service.output(self.conn, filled())
        self.assertTrue(result.ok)
        self.assertEqual(result.seq_no, 1)

    def test_出力のたびに連番が進む(self):
        self.assertEqual(meisai_service.output(self.conn, filled()).seq_no, 1)
        self.assertEqual(meisai_service.output(self.conn, filled(), confirm=True).seq_no, 2)

    def test_出力すると印刷済フラグがおりる(self):
        history_repo.set_printed(self.conn, True)
        meisai_service.output(self.conn, filled())
        self.assertFalse(history_repo.is_printed(self.conn))

    def test_副番が履歴へ入る(self):
        meisai_service.output(self.conn, filled())
        self.assertEqual(history_repo.fuban_count(self.conn), 6)
        self.assertTrue(history_repo.is_duplicate(self.conn, "L5160Z0-1-10"))

    def test_行数が紙を超えたら断る(self):
        # 積み条数の上限が15なのでふつうは起きないが、念のため
        b = board(zen=40, jou=2, stack=config.STACK_LIMIT)
        for key in b.all_keys[:config.STACK_LIMIT]:
            b.stack(key)
        b.slots.append("1-20")          # 無理やり16本目
        result = meisai_service.output(self.conn, b)
        self.assertFalse(result.ok)
        self.assertEqual(result.refuse, meisai_service.REFUSE_TOO_MANY)


class DuplicateTest(unittest.TestCase):
    """RULE-07: 副番の重複。"""

    def setUp(self):
        self.conn = _db.case_db(self)
        meisai_service.output(self.conn, filled())   # 1回目

    def test_重複すると確認を返す(self):
        result = meisai_service.output(self.conn, filled())
        self.assertFalse(result.ok)
        self.assertTrue(result.needs_confirm)
        self.assertEqual([c.kind for c in result.confirms], [CONFIRM_DUPLICATE])

    def test_確認の内訳は積み上げ順(self):
        b = filled(keys=("2-8", "1-10"))
        result = meisai_service.output(self.conn, b)
        self.assertEqual(result.confirms[0].detail, ["2-8", "1-10"])

    def test_確認しなければ何も書かれない(self):
        before = history_repo.fuban_count(self.conn)
        meisai_service.output(self.conn, filled())
        self.assertEqual(history_repo.fuban_count(self.conn), before)
        self.assertEqual(history_repo.current_seq(self.conn), 1)

    def test_了解すれば出力できる(self):
        result = meisai_service.output(self.conn, filled(), confirm=True)
        self.assertTrue(result.ok)
        self.assertEqual(result.seq_no, 2)

    def test_同じ副番は二重に登録しない(self):
        # VBA: If Not IsDuplicate Then RegisterFuban
        meisai_service.output(self.conn, filled(), confirm=True)
        self.assertEqual(history_repo.fuban_count(self.conn), 6)

    def test_No指定は受け付けない(self):
        """No指定は外した(VER 0.13.23)。出力はいつも次の連番で、重複も必ず確かめる。"""
        with self.assertRaises(TypeError):
            meisai_service.output(self.conn, filled(), specify_no=1)   # noqa
        result = meisai_service.output(self.conn, filled())
        self.assertEqual([c.kind for c in result.confirms], [CONFIRM_DUPLICATE])

    def test_一部だけ重複していればその分だけ出す(self):
        b = filled(keys=("1-10", "1-1"))       # 1-10 は既出、1-1 は新しい
        result = meisai_service.output(self.conn, b)
        self.assertEqual(result.confirms[0].detail, ["1-10"])


class SeqTest(unittest.TestCase):
    """RULE-08: 連番(No指定は VER 0.13.23 で外した)。"""

    def setUp(self):
        self.conn = _db.case_db(self)

    def test_連番のリセット(self):
        meisai_service.output(self.conn, filled())
        history_repo.reset_seq(self.conn)
        self.assertEqual(history_repo.current_seq(self.conn), 0)

    def test_副番履歴を消しても連番は戻さない(self):
        # VBA: 「seqNo は副番履歴クリアでリセットしない」。画面の副番履歴クリアは
        # VER 0.13.23 で外したが、ロット切替・全クリアの片付け(`close_lot`)が使う
        meisai_service.output(self.conn, filled())
        history_repo.clear_fuban(self.conn)
        self.assertEqual(history_repo.current_seq(self.conn), 1)
        self.assertEqual(history_repo.fuban_count(self.conn), 0)
        self.assertTrue(history_repo.is_printed(self.conn))


def three_outputs(conn) -> Board:
    """No1〜No3 を出した盤面(画面と同じ流れ: 積む → 出力 → 積み上げだけ空にする)。"""
    b = board(stack=2)
    for keys in (("1-1", "1-2"), ("1-3", "1-4"), ("2-1", "2-2")):
        b.set_stack_max(len(keys))
        for key in keys:
            b.stack(key)
        result = meisai_service.output(conn, b, confirm=True)
        assert result.ok, result.message
        meisai_service.reset_stack(b)
    return b


class EditOutputTest(unittest.TestCase):
    """出力済みの修正(VER 0.13.23・現場の求め)。"""

    def setUp(self):
        self.conn = _db.case_db(self)
        self.b = three_outputs(self.conn)

    def _keys(self, no):
        return meisai_service.find_output(self.conn, "L5160Z0", no).keys

    def test_修正に入るとそのNoのコイルが積み上げへ戻る(self):
        meisai_service.begin_edit(self.conn, self.b, 1)
        self.assertTrue(self.b.editing)
        self.assertEqual(self.b.slots, ["1-1", "1-2"])
        self.assertEqual(self.b.stack_max, 2)
        # ほかのNoのコイルは使用済みのまま
        for key in ("1-3", "1-4", "2-1", "2-2"):
            self.assertEqual(self.b.state_of(key), "used", key)

    def test_ほかのNoのコイルは積めない(self):
        """No1〜No3 があってNo1を直すとき、No2・No3 のコイルは使えない。"""
        meisai_service.begin_edit(self.conn, self.b, 1)
        self.b.unstack(2)
        for key in ("1-3", "2-2"):
            with self.subTest(key=key), self.assertRaises(RefusedError) as cm:
                self.b.stack(key)
            self.assertEqual(cm.exception.reason, strand_service.REFUSE_USED)
        self.b.stack("1-5")                     # 未使用なら積める
        self.assertEqual(self.b.slots, ["1-1", "1-5"])

    def test_最後の砦でもほかのNoのコイルは断る(self):
        meisai_service.begin_edit(self.conn, self.b, 1)
        self.b.slots[1] = "1-3"                 # 画面を通さずに無理やり
        result = meisai_service.output_edit(self.conn, self.b, confirm=True)
        self.assertFalse(result.ok)
        self.assertEqual(result.refuse, meisai_service.REFUSE_USED_ELSEWHERE)
        self.assertEqual(self._keys(1), ["1-1", "1-2"])

    def test_いつも確かめて_確かめるまでは何も書かない(self):
        meisai_service.begin_edit(self.conn, self.b, 1)
        self.b.unstack(2)
        self.b.stack("1-5")
        result = meisai_service.output_edit(self.conn, self.b)
        self.assertEqual([c.kind for c in result.confirms], [CONFIRM_EDIT])
        detail = result.confirms[0].detail
        self.assertIn("前: 1-1 1-2", detail)
        self.assertIn("後: 1-1 1-5", detail)
        self.assertIn("外すコイル: 1-2", detail)
        self.assertIn("加えるコイル: 1-5", detail)
        self.assertEqual(self._keys(1), ["1-1", "1-2"])

    def test_同じNoのまま出し直す(self):
        meisai_service.begin_edit(self.conn, self.b, 1)
        self.b.unstack(2)
        self.b.stack("1-5")
        history_repo.set_printed(self.conn, True)
        result = meisai_service.output_edit(self.conn, self.b, confirm=True)
        self.assertTrue(result.ok, result.message)
        self.assertEqual(result.seq_no, 1)
        self.assertEqual(self._keys(1), ["1-1", "1-5"])
        self.assertEqual([o.seq_no for o in meisai_service.list_outputs(self.conn, "L5160Z0")],
                         [1, 2, 3])
        self.assertEqual(history_repo.current_seq(self.conn), 3)   # 連番は進まない
        self.assertFalse(history_repo.is_printed(self.conn))       # 刷り直しが要る

    def test_副番履歴は外れた分を外し加わった分を足す(self):
        meisai_service.begin_edit(self.conn, self.b, 1)
        self.b.unstack(2)
        self.b.stack("1-5")
        meisai_service.output_edit(self.conn, self.b, confirm=True)
        self.assertFalse(history_repo.is_duplicate(self.conn, "L5160Z0-1-2"))
        self.assertTrue(history_repo.is_duplicate(self.conn, "L5160Z0-1-5"))
        self.assertTrue(history_repo.is_duplicate(self.conn, "L5160Z0-1-3"))   # No2 のまま
        self.assertEqual(history_repo.fuban_count(self.conn), 6)

    def test_修正を終えると外したコイルは空き_加えたコイルは使用済み(self):
        meisai_service.begin_edit(self.conn, self.b, 1)
        self.b.unstack(2)
        self.b.stack("1-5")
        meisai_service.output_edit(self.conn, self.b, confirm=True)
        self.b.end_edit(committed=True)
        self.assertFalse(self.b.editing)
        self.assertEqual(self.b.state_of("1-2"), "available")
        self.assertEqual(self.b.state_of("1-5"), "used")
        self.assertFalse(any(self.b.slots))

    def test_やめると前のまま(self):
        meisai_service.begin_edit(self.conn, self.b, 2)
        self.b.unstack(1)
        self.b.end_edit(committed=False)
        self.assertFalse(self.b.editing)
        for key in ("1-3", "1-4"):
            self.assertEqual(self.b.state_of(key), "used")
        self.assertEqual(self._keys(2), ["1-3", "1-4"])

    def test_重量はそのNoを出したときの値で出し直し_終われば戻す(self):
        # No1 を出したあとで重量を変えた
        self.b.weights = [300, 299]
        meisai_service.begin_edit(self.conn, self.b, 1)
        self.assertEqual(self.b.weights, [250, 248])
        meisai_service.output_edit(self.conn, self.b, confirm=True)
        self.assertEqual(meisai_service.find_output(self.conn, "L5160Z0", 1).weights, [250, 248])
        self.b.end_edit(committed=True)
        self.assertEqual(self.b.weights, [300, 299])

    def test_書き足しは引き継ぐ(self):
        meisai_service.save_hand_edits(self.conn, "L5160Z0", {
            f"{meisai_service.find_output(self.conn, 'L5160Z0', 1).row_id}.size": "0.3x50",
            f"{meisai_service.find_output(self.conn, 'L5160Z0', 1).row_id}.lotno": "AB12"})
        meisai_service.begin_edit(self.conn, self.b, 1)
        meisai_service.output_edit(self.conn, self.b, confirm=True)
        found = meisai_service.find_output(self.conn, "L5160Z0", 1)
        self.assertEqual((found.hand_size, found.hand_lotno), ("0.3x50", "AB12"))

    def test_前の紙の履歴は作り直し_新しい紙は有効(self):
        old_id = meisai_service.find_output(self.conn, "L5160Z0", 1).history_id
        meisai_service.begin_edit(self.conn, self.b, 1)
        meisai_service.output_edit(self.conn, self.b, confirm=True)
        new_id = meisai_service.find_output(self.conn, "L5160Z0", 1).history_id
        self.assertNotEqual(old_id, new_id)
        states = dict(self.conn.execute("SELECT 送信ID, 状態 FROM 明細履歴").fetchall())
        self.assertEqual(states[old_id], slip_history.STATE_REPLACED)
        self.assertEqual(states[new_id], slip_history.STATE_VALID)

    def test_断る場面(self):
        with self.subTest("無いNo"), self.assertRaises(RefusedError) as cm:
            meisai_service.begin_edit(self.conn, self.b, 9)
        self.assertEqual(cm.exception.reason, meisai_service.REFUSE_NO_OUTPUT)
        with self.subTest("積み上げ中"):
            self.b.set_stack_max(1)
            self.b.stack("1-6")
            with self.assertRaises(RefusedError) as cm:
                meisai_service.begin_edit(self.conn, self.b, 1)
            self.assertEqual(cm.exception.reason, strand_service.REFUSE_STACKING)
            self.b.unstack(1)
        with self.subTest("修正中にもう1枚"):
            meisai_service.begin_edit(self.conn, self.b, 1)
            with self.assertRaises(RefusedError) as cm:
                meisai_service.begin_edit(self.conn, self.b, 2)
            self.assertEqual(cm.exception.reason, strand_service.REFUSE_EDITING)
        with self.subTest("修正中でなければ修正の出力はできない"):
            self.b.end_edit(committed=False)
            result = meisai_service.output_edit(self.conn, self.b, confirm=True)
            self.assertEqual(result.refuse, meisai_service.REFUSE_NOT_EDITING)

    def test_修正中の途中経過は前の紙のコイルも使用済みに数える(self):
        """修正中にタブが落ちて開き直しても、前のNoのコイルが空いて見えない。"""
        meisai_service.begin_edit(self.conn, self.b, 1)
        self.b.unstack(1)
        self.b.unstack(2)
        self.assertTrue({"1-1", "1-2"} <= self.b.committed_used)
        self.assertNotIn("1-1", self.b.used)


class CancelLotHistoryTest(unittest.TestCase):
    """全クリアで紙の履歴は消さずに「取り消し」にする(VER 0.13.23)。"""

    def test_そのロットの紙を取り消しにする(self):
        conn = _db.case_db(self)
        three_outputs(conn)
        self.assertEqual(meisai_service.cancel_lot_history(conn, "L5160Z0"), 3)
        states = [r[0] for r in conn.execute("SELECT 状態 FROM 明細履歴")]
        self.assertEqual(states, [slip_history.STATE_CANCELLED] * 3)
        self.assertEqual(meisai_service.cancel_lot_history(conn, "XXXXXXX"), 0)


class SnapshotTest(unittest.TestCase):
    """途中経過の保存と復元。"""

    def setUp(self):
        self.conn = _db.case_db(self)

    def _snap(self, lot="L5160Z0") -> Snapshot:
        return Snapshot(lot_no=lot, zen_kotei=24, jou_su=2,
                        weights=[250, 248], strands=[12, 12],
                        used=["1-10", "2-8"], discarded=["2-3"])

    def test_保存して復元すると全項目が一致する(self):
        history_repo.save_snapshot(self.conn, self._snap())
        got = history_repo.load_snapshot(self.conn, "L5160Z0")
        self.assertEqual(got, self._snap())

    def test_無ければNone(self):
        self.assertIsNone(history_repo.load_snapshot(self.conn, "XXXXXXX"))

    def test_同じロットは上書きされる(self):
        history_repo.save_snapshot(self.conn, self._snap())
        second = self._snap()
        second.jou_su = 3
        second.strands = [8, 8, 8]
        history_repo.save_snapshot(self.conn, second)
        self.assertEqual(history_repo.load_snapshot(self.conn, "L5160Z0").jou_su, 3)
        self.assertEqual(len(history_repo.snapshot_lots(self.conn)), 1)

    def test_直近5件だけ残る(self):
        for i in range(7):
            history_repo.save_snapshot(self.conn, self._snap(lot=f"LOT{i:04d}"))
        self.assertEqual(len(history_repo.snapshot_lots(self.conn)),
                         config.SNAPSHOT_KEEP)

    def test_1ロットだけ消せる(self):
        history_repo.save_snapshot(self.conn, self._snap())
        history_repo.delete_snapshot(self.conn, "L5160Z0")
        self.assertIsNone(history_repo.load_snapshot(self.conn, "L5160Z0"))

    def test_全消去(self):
        history_repo.save_snapshot(self.conn, self._snap())
        history_repo.clear_snapshots(self.conn)
        self.assertEqual(history_repo.snapshot_lots(self.conn), [])

    def test_空の並びは空の一覧になる(self):
        snap = self._snap()
        snap.used = []
        snap.discarded = []
        history_repo.save_snapshot(self.conn, snap)
        got = history_repo.load_snapshot(self.conn, "L5160Z0")
        self.assertEqual(got.used, [])
        self.assertEqual(got.discarded, [])


class ReportTest(unittest.TestCase):
    """帳票。"""

    def _output(self, keys=("1-10", "1-5", "2-8"), weights=(250, 248)) -> Output:
        return Output(lot_no="L5160Z0", seq_no=1, keys=list(keys),
                      weights=list(weights))

    def test_シート名はVBAと同じ形(self):
        self.assertEqual(self._output().sheet_name, "L5160Z0-No1")

    def test_行の重量は丈から引かれる(self):
        rows = self._output().rows()
        self.assertEqual(rows, [("1-10", "250.0Kg"), ("1-5", "250.0Kg"),
                                ("2-8", "248.0Kg")])

    def test_読めないキーは重量を空にして行を残す(self):
        # VBA WriteData は不正キーを飛ばしつつ行だけ進める
        rows = self._output(keys=("1-10", "こわれた", "2-8")).rows()
        self.assertEqual(rows[1], ("こわれた", ""))
        self.assertEqual(len(rows), 3)

    def test_紙は必ず15行(self):
        html = report.render([self._output()])
        self.assertEqual(html.count('class="r-data"'), config.SHEET_ROWS)

    def test_用紙はA4横で余白0(self):
        html = report.render([self._output()])
        self.assertIn("size: A4 landscape", html)
        self.assertIn("margin: 0mm", html)

    def test_見出しの文言(self):
        html = report.render([self._output()])
        for text in ("L5160Z0-No1  梱包明細表", "NLM.NAGOYA.QA",
                     "コイル副番", "重量(Kg)", "小ラベル添付"):
            self.assertIn(text, html)

    def test_0行でも紙は成立する(self):
        html = report.render([self._output(keys=())])
        self.assertEqual(html.count('class="r-data"'), config.SHEET_ROWS)

    # --- 用紙: A4横の左半分に1枚、右半分は空ける ---------------------
    def _outs(self, *nos):
        outs = []
        for no in nos:
            out = self._output()
            out.seq_no = no
            outs.append(out)
        return outs

    def test_明細1枚は左半分で右半分は空ける(self):
        """**現場は右半分を切り取って使う。** VBA の紙面も同じ
        (`FitToPages` は縮めるだけで引き伸ばさないので、1面は等倍で
        左半分に出る)。

        **右半分を何かで埋めない。** 別の明細を並べる運用はしておらず、
        同じ明細を写すと同じ梱包の紙が2枚になる。
        """
        html = report.render([self._output()])
        self.assertEqual(html.count('class="page"'), 1)
        self.assertEqual(html.count('class="meisai"'), 1)
        self.assertIn('<div class="half blank"></div></div>', html)

    def test_まとめて刷っても1枚ずつ別の用紙(self):
        """**A4 1枚に2枚並べる運用はしていない**(現場に確認済み)。"""
        for count in (2, 3, 4):
            with self.subTest(count=count):
                html = report.render(self._outs(*range(1, count + 1)))
                self.assertEqual(html.count('class="page"'), count)
                self.assertEqual(html.count('class="meisai"'), count)
                self.assertEqual(html.count('class="half blank"'), count)

    def test_どの用紙も左に明細_右は空(self):
        """紙ごとに「明細 → 空」の順。右に明細が来る紙を作らない。"""
        html = report.render(self._outs(1, 2, 3))
        pages = html.split('<div class="page">')[1:]
        for page in pages:
            left, right = page.split('<div class="half blank">', 1)
            self.assertIn('class="meisai"', left)
            self.assertNotIn('class="meisai"', right.split('</div></div>', 1)[0])

    def test_右半分に線を引かない(self):
        """VBA の紙面の右半分は何も無い(切り取る側)。"""
        self.assertNotIn(".half + .half", report._sheet_css())
        self.assertNotIn("dashed", report._sheet_css())

    # --- 刷れる範囲と文字の寄せ ---------------------------------------
    @staticmethod
    def _rule(selector: str) -> str:
        """`_sheet_css()` の中の1つの規則の中身。"""
        css = report._sheet_css()
        start = css.index(selector + "{")
        return css[start:css.index("}", start)]

    def test_紙の端から5mm内に置く(self):
        """プリンタは紙の縁から約4mmを刷れない。**刷ると左と上が欠けて
        いた**(現場の実機。プレビューでは収まって見える)。

        余白は中身(`.half` の内側)で取る ── 現場は「余白なし」で刷る
        ので、`@page` の余白は上書きされて効かない。右も同じだけ空ける
        (切り目が紙の端になる)。
        """
        self.assertGreaterEqual(report.SAFE_MARGIN_MM, 5.0)
        half = self._rule(".half")
        # 線の太さぶん足してある(表の外枠の線は箱から少しはみ出す)
        self.assertIn(f"padding:calc({report.SAFE_MARGIN_MM:g}mm + 1px);", half)
        self.assertIn("overflow:hidden", half)

    def test_副番と重量は中央(self):
        """VBA は `MergePut` で結合して `xlCenter` にしてから値を入れる。
        以前は副番を左・重量を右に寄せていて、VBA と違っていた。
        """
        self.assertIn("text-align:center", self._rule(".meisai .r-data td"))
        css = report._sheet_css()
        self.assertNotIn("text-align:right", css)
        self.assertNotIn("text-align:left", css)

    def test_書き足したサイズとLOTNOも中央(self):
        self.assertIn("text-align:center", self._rule(".meisai .r-label .hand .hl"))

    def test_どの文字も端に寄せない(self):
        """見出し・小ラベル添付・左の見出しも中央(VBA はどれも `xlCenter`)。"""
        for selector in (".meisai .r-title td", ".meisai .r-label .cap",
                         ".meisai .r-label .paste", ".meisai .r-head td"):
            with self.subTest(selector=selector):
                self.assertIn("text-align:center", self._rule(selector))

    def test_寸法は実ブックの比率(self):
        """行の合計が 595.5pt = A4の高さ。列は1面ぶん(=A4横の半分)で100%。"""
        self.assertAlmostEqual(sum(report._col_percents()), 100.0, places=3)
        title, label, head, data = report._row_percents()
        self.assertAlmostEqual(title + label + head + data * config.SHEET_ROWS,
                               100.0, places=3)


class OutputListTest(unittest.TestCase):
    """出力の一覧と削除。"""

    def setUp(self):
        self.conn = _db.case_db(self)

    def test_No順に並ぶ(self):
        meisai_service.output(self.conn, filled())
        meisai_service.output(self.conn, filled(), confirm=True)
        got = meisai_service.list_outputs(self.conn, "L5160Z0")
        self.assertEqual([o.seq_no for o in got], [1, 2])

    def test_ロットが違えば混ざらない(self):
        meisai_service.output(self.conn, filled())
        meisai_service.output(self.conn, filled(lot="AAA0001"))
        self.assertEqual(len(meisai_service.list_outputs(self.conn, "L5160Z0")), 1)

    def test_ロットごと消せる(self):
        meisai_service.output(self.conn, filled())
        meisai_service.output(self.conn, filled(), confirm=True)
        self.assertEqual(meisai_service.delete_lot_outputs(self.conn, "L5160Z0"), 2)
        self.assertEqual(meisai_service.list_outputs(self.conn, "L5160Z0"), [])

    def test_出力の中身が読み戻せる(self):
        meisai_service.output(self.conn, filled())
        got = meisai_service.find_output(self.conn, "L5160Z0", 1)
        self.assertEqual(got.keys, ["1-10", "1-5", "2-8", "2-11", "2-12", "2-9"])
        self.assertEqual(got.weights, [250, 248])

    def test_出力後に積み上げだけ空になる(self):
        # VBA: ClearGimmickC のみ。条番号のキャンバスは残る
        b = filled()
        meisai_service.output(self.conn, b)
        meisai_service.reset_stack(b)
        self.assertEqual(b.slots, [None] * 6)
        self.assertEqual(b.strands, [12, 12])

    def test_出力したコイルは使用済みのまま残る(self):
        """VBA「Cスロットのみ。Bは出力後もそのまま残す」。

        **同じコイルを次の梱包へ二重に積ませない**ための仕掛け。
        ここが消えると、現物と紙が食い違う。
        """
        b = filled()
        keys = b.stacked_keys()
        meisai_service.output(self.conn, b)
        meisai_service.reset_stack(b)

        self.assertEqual(b.used, set(keys))
        for key in keys:
            self.assertEqual(b.state_of(key), "used")

    def test_出力済みのコイルは積み直せない(self):
        b = filled()
        meisai_service.output(self.conn, b)
        meisai_service.reset_stack(b)
        with self.assertRaises(RefusedError) as cm:
            b.stack("1-10")
        self.assertEqual(cm.exception.reason, "used")

    def test_条番号を作り直せば使用済みは消える(self):
        # VBA: BuildGimmickB がラベルを作り直すので IsUsed も消える
        b = filled()
        meisai_service.output(self.conn, b)
        meisai_service.reset_stack(b)
        b.build_strands()
        self.assertEqual(b.used, set())


if __name__ == "__main__":
    unittest.main()
