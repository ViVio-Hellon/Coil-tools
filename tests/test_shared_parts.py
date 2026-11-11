"""共通の部品: 「読めない」の親・共有フォルダの時計(統合 1.2.72・構造点検の改善案3)"""
from __future__ import annotations

import os
import shutil
import tempfile
import time
import unittest
from pathlib import Path
from unittest import mock

from common import shared_clock, unreadable


class UnreadableTest(unittest.TestCase):
    """機能ごとに7通りあった「読めない」の例外は、どれも `common.unreadable.Unreadable` を親に持つ。"""

    def test_7つとも同じ親を持ち_前の親もそのまま(self):
        from common import local_settings, remote_stop
        from modules.packing_details.meisai import user_settings as details_settings
        from modules.packing_material_calculation.coil_tool import order_history
        from modules.packing_material_calculation.coil_tool import user_settings as material_settings
        from modules.packing_pena_label.app import config as pena_config
        classes = {
            local_settings.Unreadable: OSError, details_settings.SettingsUnreadable: OSError,
            pena_config.LocalConfigUnreadable: OSError,
            remote_stop.Unreadable: Exception, remote_stop.RulesUnreadable: Exception,
            order_history.HistoryUnreadable: Exception, material_settings.SettingsUnreadable: Exception,
        }
        for cls, old_parent in classes.items():
            with self.subTest(cls=cls.__qualname__):
                self.assertTrue(issubclass(cls, unreadable.Unreadable))
                self.assertTrue(issubclass(cls, old_parent), "いままでの except で受けられなくなった")

    def test_1つの名前で受けられる(self):
        from modules.packing_details.meisai import user_settings as details_settings
        with self.assertRaises(unreadable.Unreadable):
            raise details_settings.SettingsUnreadable("壊れている")


class SharedClockTest(unittest.TestCase):
    def setUp(self):
        self.folder = Path(tempfile.mkdtemp(prefix="cpt-clock-"))
        self.addCleanup(shutil.rmtree, self.folder, True)

    def test_共有の時計を測り_試しのファイルを残さない(self):
        at = shared_clock.now(self.folder)
        self.assertIsNotNone(at)
        self.assertLess(abs(at - time.time()), 60)
        self.assertLess(abs(shared_clock.offset(self.folder)), 60)
        self.assertEqual(os.listdir(self.folder), [])

    def test_PCの時計がずれていれば_そのずれを返す(self):
        real = time.time
        with mock.patch.object(shared_clock.time, "time", side_effect=lambda: real() + 300):
            self.assertAlmostEqual(shared_clock.offset(self.folder), -300, delta=5)

    def test_書けなければNone(self):
        self.assertIsNone(shared_clock.now(self.folder / "無いフォルダ"))
        self.assertIsNone(shared_clock.offset(self.folder / "無いフォルダ"))


if __name__ == "__main__":
    unittest.main()


class Ls4PickRowsTest(unittest.TestCase):
    """実績の形式の行の選び方(統合 1.2.80・現場の指示: 月を問わず、今の月に近いほうから)。"""

    def test_月を問わずロットごとにいちばん新しい行(self):
        from datetime import date
        from common import ls4_source
        rows = [{"ﾛｯﾄ番号": "A", "BOX実績_作業開始_日付": "2026-09-30 23:00:00", "v": "先々月"},
                {"ﾛｯﾄ番号": "A", "BOX実績_作業開始_日付": "2026/10/01 01:00:00", "v": "先月"},
                {"ﾛｯﾄ番号": "B", "BOX実績_作業開始_日付": "", "v": "日付なし"},
                {"ﾛｯﾄ番号": "B", "BOX実績_作業開始_日付": "2026-08-01 00:00:00", "v": "8月"}]
        got = {r["ﾛｯﾄ番号"]: r["v"] for r in ls4_source.pick_rows(rows, date(2026, 11, 11))}
        self.assertEqual(got, {"A": "先月", "B": "8月"})

    def test_絞るときは何か月前まで(self):
        from datetime import date
        from common import ls4_source
        rows = [{"ﾛｯﾄ番号": "A", "BOX実績_作業開始_日付": "2026-11-02"},
                {"ﾛｯﾄ番号": "B", "BOX実績_作業開始_日付": "2026-10-31"},
                {"ﾛｯﾄ番号": "C", "BOX実績_作業開始_日付": "2026-09-30"}]
        pick = lambda n: sorted(r["ﾛｯﾄ番号"] for r in ls4_source.pick_rows(rows, date(2026, 11, 11), months_back=n))
        self.assertEqual(pick(0), ["A"])
        self.assertEqual(pick(1), ["A", "B"])
        self.assertEqual(ls4_source.MONTHS_BACK, None, "既定は月を問わない")
