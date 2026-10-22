# -*- coding: utf-8 -*-
"""記録ファイルの書き方(common/record_file.py。統合 1.2.44・現場の指摘: .jsonl は普通は開けない)

- 既定は Excel で開ける CSV(UTF-8 BOM・CRLF・見出しは 1 回だけ)。JSONL・両方も選べる
- 決まった列(日時・種類・内容・モード・PC名・ライン・利用者・版)と、全部を入れた「詳しい中身」
- Excel で開いていて書けないときは「追記待ち」に足し、次に書けたときに本体へ書き足して消す
"""
import datetime
import json
import os
import shutil
import tempfile
import unittest
from unittest import mock

from common import record_file as RF

WHEN = datetime.datetime(2026, 10, 22, 9, 0, 0)


def _entry(kensa="N6X0945", **kw):
    e = {"at": "2026-10-22 09:00:00", "kind": "計算", "title": "風袋計算", "kensaNo": kensa,
         "coils": {"1": {"本数": "11", "NW": "110.0"}}, "material": {"label": "SQLite", "stale": False},
         "version": {"app": "1.2.44", "pena": "1.5.44", "stamp": "abc123"},
         "who": {"pc": "LINE-PC2", "line": "2ライン", "user": "taro", "mode": "テストモード"}}
    e.update(kw)
    return e


class RecordFileTest(unittest.TestCase):
    def setUp(self):
        self.dir = tempfile.mkdtemp(prefix="rf_")

    def tearDown(self):
        shutil.rmtree(self.dir, ignore_errors=True)

    def _append(self, entry, fmt="csv"):
        return RF.append(self.dir, "ペナラベル記録", entry, fmt, feature_version_key="pena", when=WHEN)

    def test_既定はCSV_Excelで開けるBOMとCRLF_見出しは1回(self):
        self.assertEqual(RF.normalize_format(""), "csv")
        self.assertEqual(RF.normalize_format("へん"), "csv")
        self.assertEqual(self._append(_entry()), "")
        self.assertEqual(self._append(_entry("H670129")), "")
        self.assertEqual(sorted(os.listdir(self.dir)), ["ペナラベル記録_2026-10.csv"])
        path = os.path.join(self.dir, "ペナラベル記録_2026-10.csv")
        with open(path, "rb") as f:
            raw = f.read()
        self.assertTrue(raw.startswith(b"\xef\xbb\xbf"))
        self.assertEqual(raw.count(b"\xef\xbb\xbf"), 1)
        self.assertIn(b"\r\n", raw)
        rows = RF.read_csv(path)
        self.assertEqual(len(rows), 2, "見出しは 1 回だけ")
        self.assertEqual(list(rows[0].keys()), list(RF.CSV_COLUMNS))
        r = rows[0]
        self.assertEqual((r["日時"], r["種類"], r["モード"], r["PC名"], r["ライン"], r["利用者"]),
                         ("2026-10-22 09:00:00", "計算", "テストモード", "LINE-PC2", "2ライン", "taro"))
        self.assertEqual((r["統合の版"], r["機能の版"], r["中身の指紋"]), ("1.2.44", "1.5.44", "abc123"))
        self.assertIn("画面: 風袋計算", r["内容"])
        self.assertIn("検査番号: N6X0945", r["内容"])
        self.assertIn("本数=11", r["内容"])
        self.assertIn("古い内容=いいえ", r["内容"])
        self.assertNotIn("LINE-PC2", r["内容"], "決まった列に入れたものは内容に出さない")
        self.assertEqual(json.loads(r["詳しい中身"]), _entry(), "全部を入れてあり、何も落とさない")
        self.assertEqual(rows[1]["内容"].count("H670129"), 1)

    def test_JSONLと両方(self):
        self._append(_entry(), "jsonl")
        self.assertEqual(os.listdir(self.dir), ["ペナラベル記録_2026-10.jsonl"])
        self._append(_entry("K6X0132"), "both")
        self.assertEqual(sorted(os.listdir(self.dir)),
                         ["ペナラベル記録_2026-10.csv", "ペナラベル記録_2026-10.jsonl"])
        with open(os.path.join(self.dir, "ペナラベル記録_2026-10.jsonl"), encoding="utf-8") as f:
            lines = [json.loads(x) for x in f.read().splitlines()]
        self.assertEqual([x["kensaNo"] for x in lines], ["N6X0945", "K6X0132"])

    def test_Excelで開いていて書けない間は追記待ちに足し_閉じたら書き足す(self):
        self._append(_entry("N6X0001"))
        path = os.path.join(self.dir, "ペナラベル記録_2026-10.csv")
        real = RF._append_text

        def locked(p, text, header=""):
            if p == path:
                raise PermissionError(13, "Permission denied")
            return real(p, text, header=header)

        with mock.patch.object(RF, "_append_text", side_effect=locked):
            msg = self._append(_entry("N6X0002"))
            msg2 = self._append(_entry("N6X0003"))
        self.assertIn("開かれていて書けません", msg)
        self.assertIn("追記待ち", msg)
        self.assertIn("閉じたあと次の記録で書き足します", msg2)
        self.assertTrue(os.path.exists(path + RF.PENDING_SUFFIX))
        self.assertEqual(len(RF.read_csv(path)), 1, "本体はそのまま")
        # 閉じた → 次の記録で、待っていた 2 件と合わせて書き足す
        self.assertEqual(self._append(_entry("N6X0004")), "")
        self.assertFalse(os.path.exists(path + RF.PENDING_SUFFIX))
        kensa = [json.loads(r["詳しい中身"])["kensaNo"] for r in RF.read_csv(path)]
        self.assertEqual(kensa, ["N6X0001", "N6X0002", "N6X0003", "N6X0004"], "順も落とさない")

    def test_追記待ちにも書けなければ理由を返す(self):
        with mock.patch.object(RF, "_append_text", side_effect=PermissionError(13, "denied")):
            msg = self._append(_entry())
        self.assertIn("に書けません", msg)

    def test_月の記録を読む_両方に書いた分は1回_同じ中身の別の記録は残す(self):
        self._append(_entry("A111111"), "csv")
        self._append(_entry("B222222"), "both")
        self._append(_entry("B222222"), "both")          # 同じ秒・同じ中身の別の記録
        self._append(_entry("C333333"), "jsonl")
        got = [x["kensaNo"] for x in RF.read_month(self.dir, "ペナラベル記録", WHEN.date())]
        self.assertEqual(sorted(got), ["A111111", "B222222", "B222222", "C333333"])

    def test_選択肢と見出し(self):
        self.assertEqual([v for v, _ in RF.FORMAT_CHOICES], ["csv", "jsonl", "both"])
        self.assertEqual(RF.format_label("both"), "CSV と JSONL の両方")
        self.assertEqual(RF.file_names("梱包明細記録", WHEN, "both"),
                         ["梱包明細記録_2026-10.csv", "梱包明細記録_2026-10.jsonl"])


if __name__ == "__main__":
    unittest.main()
